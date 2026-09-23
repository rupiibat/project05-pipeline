"""
Reassign (or unassign) one or more users to the same calling agent. Runs
inside GitHub Actions, triggered either by the single-user "Reassign Agent"
widget or the "Bulk Reassign Agent" (paste User IDs) widget on the
dashboard's Search User page (via the master-userlist-upload worker's
/reassign-agent endpoint).

Downloads master_userlist.db from R2, upserts (or deletes, for un-assign)
one row per user_id in agent_assignments, and re-uploads -- the same
read-modify-write pattern ci_ingest.py uses for bulk file ingests, just
scoped to a small explicit ID list instead of a whole spreadsheet.

Runs in its own dedicated GitHub Actions concurrency group (not the
hourly pipeline's "db-ingest" group -- see reassign_agent.yml), so it
completes in ~10-30s instead of queuing behind a run that can take
15-20+ minutes. To make that safe, this script re-downloads a fresh copy
of master_userlist.db immediately before its own upload and reapplies
just this run's assignment delta onto it, rather than blindly
re-uploading its original (now several-seconds-stale) download -- keeps
the window during which this run's upload could revert some OTHER
concurrent writer's change (the hourly pipeline updating users' totals,
say) down to a couple of seconds instead of this whole job's runtime.
The mirror-image protection (the hourly pipeline never reverting THIS
script's agent_assignments write) lives in
ci_ingest.refresh_table_from_r2, called from both api_pull_ingest.py and
build_deposit_report.py right before their own master_userlist.db
uploads.

Usage: python3 reassign_agent.py --user-ids 12345 --agent "Sathya (WFH)"
       python3 reassign_agent.py --user-ids 12345,67890,111 --agent "Sathya (WFH)"
       python3 reassign_agent.py --user-ids 12345 --agent ""   (un-assign)
"""
import argparse
import os
import sqlite3
import sys

import boto3

import ci_ingest

BASE = os.path.dirname(os.path.abspath(__file__))
MASTER_DB = os.path.join(BASE, "master_userlist.db")


def r2_client():
    return boto3.client(
        "s3",
        endpoint_url=os.environ["R2_ENDPOINT_URL"],
        aws_access_key_id=os.environ["R2_ACCESS_KEY_ID"],
        aws_secret_access_key=os.environ["R2_SECRET_ACCESS_KEY"],
        region_name="auto",
    )


def apply_assignments(db_path, user_ids, agent):
    """Upserts (or deletes, for un-assign) one agent_assignments row per
    user_id in `user_ids` against whatever users already exist in
    `db_path`. Returns (assigned, missing)."""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS agent_assignments (user_id INTEGER PRIMARY KEY, agent_name TEXT)")
    assigned, missing = [], []
    for user_id in user_ids:
        exists = cur.execute("SELECT 1 FROM users WHERE user_id = ?", (user_id,)).fetchone()
        if not exists:
            missing.append(user_id)
            continue
        if agent:
            cur.execute(
                "INSERT OR REPLACE INTO agent_assignments (user_id, agent_name) VALUES (?, ?)",
                (user_id, agent),
            )
        else:
            cur.execute("DELETE FROM agent_assignments WHERE user_id = ?", (user_id,))
        assigned.append(user_id)
    conn.commit()
    conn.close()
    return assigned, missing


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--user-ids", required=True, help="Comma-separated user IDs")
    ap.add_argument("--agent", required=True, help="Agent name, or empty string to un-assign")
    args = ap.parse_args()

    user_ids = [int(x.strip()) for x in args.user_ids.split(",") if x.strip()]
    if not user_ids:
        print("FATAL: no user IDs given", file=sys.stderr)
        sys.exit(1)

    bucket = os.environ["R2_BUCKET"]
    s3 = r2_client()

    try:
        s3.download_file(bucket, "master_userlist.db", MASTER_DB)
    except Exception as e:
        print(f"FATAL: could not download master_userlist.db from R2: {e}", file=sys.stderr)
        sys.exit(1)

    agent = args.agent.strip()
    assigned, missing = apply_assignments(MASTER_DB, user_ids, agent)

    if missing:
        print(f"Skipped {len(missing)} user_id(s) not found in users table: {missing}", file=sys.stderr)
    label = agent or "Un-Assigned"
    print(f"Assigned {len(assigned)} user(s) -> {label}: {assigned}")
    if not assigned:
        print("FATAL: no valid user IDs to reassign", file=sys.stderr)
        sys.exit(1)

    # Re-fetch a fresh copy right before uploading and reapply just this
    # run's delta onto it, instead of re-uploading the copy downloaded
    # above -- shrinks the window during which this upload could revert
    # some other concurrent writer's change (e.g. the hourly pipeline
    # updating users' recharge totals) from this whole job's runtime down
    # to the time between this re-download and the upload below. Falls
    # back to the original download if the refresh itself fails, so a
    # transient R2 hiccup here doesn't turn a successful reassignment into
    # a failure.
    fresh_path = MASTER_DB + ".fresh"
    try:
        ci_ingest.download_with_retry(s3, bucket, "master_userlist.db", fresh_path, attempts=2)
        apply_assignments(fresh_path, assigned, agent)
        os.replace(fresh_path, MASTER_DB)
    except Exception as e:
        print(f"WARNING: could not refresh before upload, uploading original download instead: {e}", file=sys.stderr)
        if os.path.exists(fresh_path):
            os.remove(fresh_path)

    s3.upload_file(MASTER_DB, bucket, "master_userlist.db")
    print("Uploaded updated master_userlist.db")


if __name__ == "__main__":
    main()
