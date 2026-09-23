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
ac = data.get("action_center", {})

def band_count(key, lo, hi):
    rows = ac.get(key, {}).get("rows", [])
    return sum(1 for r in rows if r.get("inactive_days") is not None and lo <= r["inactive_days"] <= hi)

result = {}
result["active_low_0_7"] = band_count("active_low", 0, 7)
result["active_low_8_12"] = band_count("active_low", 8, 12)
result["active_low_13_15"] = band_count("active_low", 13, 15)
result["active_high_0_7"] = band_count("active_high", 0, 7)
result["active_high_8_11"] = band_count("active_high", 8, 11)
result["active_high_12_15"] = band_count("active_high", 12, 15)

result["inactive_low_16_30"] = band_count("inactive_low", 16, 30)
result["inactive_low_31_60"] = band_count("inactive_low", 31, 60)
result["inactive_low_61_90"] = band_count("inactive_low", 61, 90)
result["inactive_low_91_180"] = band_count("inactive_low", 91, 180)

result["inactive_high_16_30"] = band_count("inactive_high", 16, 30)
result["inactive_high_31_60"] = band_count("inactive_high", 31, 60)
result["inactive_high_61_90"] = band_count("inactive_high", 61, 90)
result["inactive_high_91_180"] = band_count("inactive_high", 91, 180)

result["near_upgrade_low_total"] = ac.get("near_upgrade_low", {}).get("total_matching", 0)
result["near_upgrade_high_total"] = ac.get("near_upgrade_high", {}).get("total_matching", 0)

claim_details = data.get("bonus_claims", {}).get("wallet_claim_details", [])
result["bonus_claimed_total"] = len(claim_details)
result["bonus_not_yet_redeposited"] = sum(1 for r in claim_details if r.get("deposited_after") == "No")
result["bonus_already_redeposited"] = sum(1 for r in claim_details if r.get("deposited_after") == "Yes")

result["profit_users_count"] = len(data.get("profit_users", []))
result["weekly_cashback_shield_count"] = len(data.get("weekly_cashback_shield", {}).get("rows", []))

print("=== BUCKETS_JSON_START ===")
print(json.dumps(result, default=str))
print("=== BUCKETS_JSON_END ===")
