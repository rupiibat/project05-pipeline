"""
Creates a new agent name with zero users assigned, so it shows up in the
Reassign Agent dropdown (and everywhere else agent_list.json is read --
Agent Logins, Performance page, etc.) before anyone's actually been moved
to them.

Without this, an agent name only exists implicitly: agent_list is derived
purely from DISTINCT agent_name values in agent_assignments (see
build_deposit_report.py), so a brand-new agent with no users yet had no way
to appear anywhere, and the Reassign Agent dropdown -- populated from that
same list -- had no way to offer them as a target.

Stores known agent names in a new `agent_roster` table in
master_userlist.db, independent of agent_assignments. build_deposit_report.py
unions this into agent_list going forward, so a roster entry keeps an agent
visible even if every one of their users is later reassigned away (a
returning/relaunched agent shouldn't have to be "recreated").

Triggered by the "Create New Agent" widget on the master-userlist-upload
worker's page (via its /create-agent endpoint) -- same pattern as
reassign_agent.py / rename_agent.py.

Usage: python3 create_agent.py --agent-name "Priya (WFH)"
"""
import argparse
import os
import sys
from datetime import datetime, timedelta

import boto3
import sqlite3

BASE = os.path.dirname(os.path.abspath(__file__))
MASTER_DB = os.path.join(BASE, "master_userlist.db")
DAILY_DB = os.path.join(BASE, "daily_records.db")


def r2_client():
    return boto3.client(
        "s3",
        endpoint_url=os.environ["R2_ENDPOINT_URL"],
        aws_access_key_id=os.environ["R2_ACCESS_KEY_ID"],
        aws_secret_access_key=os.environ["R2_SECRET_ACCESS_KEY"],
        region_name="auto",
    )


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--agent-name", required=True)
    args = ap.parse_args()

    agent_name = args.agent_name.strip()
    if not agent_name:
        print("FATAL: agent name is empty", file=sys.stderr)
        sys.exit(1)
    if agent_name == "Un-Assigned":
        print('FATAL: "Un-Assigned" is a reserved label, not a real agent name', file=sys.stderr)
        sys.exit(1)

    bucket = os.environ["R2_BUCKET"]
    s3 = r2_client()
    try:
        s3.download_file(bucket, "master_userlist.db", MASTER_DB)
        # Not modified here, but build_deposit_report.py (run right after
        # this script, in the same job workspace, since create_agent.yml
        # -- unlike reassign_agent.yml -- refreshes the report
        # synchronously) needs it present locally. Same pattern as
        # ban_user.py.
        s3.download_file(bucket, "daily_records.db", DAILY_DB)
    except Exception as e:
        print(f"FATAL: could not download DBs from R2: {e}", file=sys.stderr)
        sys.exit(1)

    conn = sqlite3.connect(MASTER_DB)
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS agent_roster (agent_name TEXT PRIMARY KEY, created_at TEXT)")

    already_in_roster = cur.execute("SELECT 1 FROM agent_roster WHERE agent_name = ?", (agent_name,)).fetchone()
    already_assigned = cur.execute(
        "SELECT 1 FROM agent_assignments WHERE agent_name = ? LIMIT 1", (agent_name,)
    ).fetchone()
    if already_in_roster or already_assigned:
        conn.close()
        print(f"NOTE: agent '{agent_name}' already exists (already in roster or already has assigned users) -- nothing to do")
        return

    now = (datetime.utcnow() + timedelta(hours=5, minutes=30)).strftime("%Y-%m-%d %H:%M:%S")
    cur.execute("INSERT INTO agent_roster (agent_name, created_at) VALUES (?, ?)", (agent_name, now))
    conn.commit()
    conn.close()

    s3.upload_file(MASTER_DB, bucket, "master_userlist.db")
    print(f"Created agent '{agent_name}' at {now} -- will appear in the Reassign Agent dropdown after the next report refresh")


if __name__ == "__main__":
    main()
