import json, os, boto3
from collections import Counter

def r2_client():
    return boto3.client(
        "s3",
        endpoint_url=os.environ["R2_ENDPOINT_URL"],
        aws_access_key_id=os.environ["R2_ACCESS_KEY_ID"],
        aws_secret_access_key=os.environ["R2_SECRET_ACCESS_KEY"],
        region_name="auto",
    )

s3 = r2_client()
bucket = os.environ["R2_BUCKET"]
obj = s3.get_object(Bucket=bucket, Key="reports/deposit_report.json")
data = json.loads(obj["Body"].read())

result = {}
result["top_level_keys"] = sorted(data.keys())
bc = data.get("bonus_claims", {})
result["bonus_claims_keys"] = sorted(bc.keys())
cd = bc.get("wallet_claim_details", [])
result["claim_count"] = len(cd)
dates = Counter(str(r.get("claimed_time"))[:10] for r in cd)
result["claimed_time_date_counts"] = dict(dates)
result["sample_rows"] = cd[:3]
result["generated_at"] = data.get("generated_at")

print("=== DIAG_JSON_START ===")
print(json.dumps(result, default=str))
print("=== DIAG_JSON_END ===")
