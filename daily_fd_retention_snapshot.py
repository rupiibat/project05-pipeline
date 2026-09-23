"""Runs daily (see .github/workflows/daily_fd_retention_snapshot.yml) to
compute yesterday's FD Users Retention numbers (same definition as
fd_users_retention_report() in build_deposit_report.py) and commit them to
reports/fd_retention_snapshot.json -- a stable, permanent file (overwritten
each run, not a one-off diagnostic) that a separate scheduled process reads
to push the report as an image without needing R2 credentials of its own.
"""
import datetime as dt
import json
import os
import subprocess
import sys

import boto3

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_deposit_report import fd_users_retention_report

BASE = os.path.dirname(os.path.abspath(__file__))
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
    bucket = os.environ["R2_BUCKET"]
    s3 = r2_client()
    s3.download_file(bucket, "daily_records.db", DAILY_DB)

    now_ist = dt.datetime.utcnow() + dt.timedelta(hours=5, minutes=30)
    result = fd_users_retention_report(DAILY_DB, [], [], now_ist)

    out_path = os.path.join(BASE, "reports")
    os.makedirs(out_path, exist_ok=True)
    with open(os.path.join(out_path, "fd_retention_snapshot.json"), "w") as f:
        json.dump(result, f, indent=2, default=str)

    subprocess.run(["git", "config", "user.email", "pipeline@bot.local"], check=True)
    subprocess.run(["git", "config", "user.name", "pipeline-bot"], check=True)
    subprocess.run(["git", "add", "reports/fd_retention_snapshot.json"], check=True)
    subprocess.run(["git", "commit", "-m", "FD retention snapshot for " + result["date"], "--allow-empty"], check=True)

    # This repo sees other pushes land around the same time (manual
    # diagnostic commits, feature commits), and a plain `git push` fails
    # outright if origin/main moved since checkout -- confirmed as the
    # cause of the 2026-09-13 run's silent failure (no new snapshot
    # committed that day). Retry with a rebase onto the latest origin/main
    # instead of giving up on the first race.
    for attempt in range(5):
        push = subprocess.run(["git", "push"], capture_output=True, text=True)
        if push.returncode == 0:
            break
        print(f"push attempt {attempt + 1} failed, rebasing and retrying:\n{push.stderr}")
        subprocess.run(["git", "pull", "--rebase", "origin", "main"], check=True)
    else:
        raise RuntimeError("git push failed after 5 rebase-and-retry attempts")

    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
