"""
QuickBase Debug - Check field mappings and test API directly
"""

import json
import os
import requests

# Load settings
with open("local.settings.json", "r") as f:
    settings = json.load(f)
    for key, value in settings.get("Values", {}).items():
        os.environ[key] = str(value)

# QB Config
realm = os.environ.get('QB_REALM')
user_token = os.environ.get('QB_USER_TOKEN')
app_id = os.environ.get('QB_APP_ID')
table_id = os.environ.get('QB_TICKETS_TABLE_ID')

headers = {
    'QB-Realm-Hostname': realm,
    'Authorization': f'QB-USER-TOKEN {user_token}',
    'Content-Type': 'application/json'
}

print(f"Realm: {realm}")
print(f"App ID: {app_id}")
print(f"Table ID: {table_id}")
print(f"Token: {user_token[:10]}..." if user_token else "NO TOKEN!")

# ============================================================================
# Step 1: Get table fields to see correct field IDs
# ============================================================================
print("\n" + "="*60)
print("STEP 1: Getting table fields...")
print("="*60)

fields_url = f"https://api.quickbase.com/v1/fields?tableId={table_id}"
response = requests.get(fields_url, headers=headers)

print(f"Status: {response.status_code}")

if response.status_code == 200:
    fields = response.json()
    print(f"\n📋 Table has {len(fields)} fields:\n")
    print(f"{'ID':<6} {'Type':<15} {'Label'}")
    print("-"*60)
    for field in fields:
        fid = field.get('id')
        ftype = field.get('fieldType', 'unknown')
        label = field.get('label', 'no label')
        print(f"{fid:<6} {ftype:<15} {label}")
else:
    print(f"Error: {response.text}")

# ============================================================================
# Step 2: Try creating a ticket with verbose output
# ============================================================================
print("\n" + "="*60)
print("STEP 2: Testing ticket creation...")
print("="*60)

# You need to update these field IDs based on Step 1 output!
# These are GUESSES - check the output above for correct IDs
record_data = {
    "to": table_id,
    "data": [{
        # UPDATE THESE FIELD IDs based on Step 1 output!
        "6": {"value": "TEST-9999"},           # Ticket Number (or auto-number?)
        "7": {"value": "Test from debug.py"},  # Subject
        "8": {"value": "Debug test ticket"},   # Description
        "9": {"value": "Low"},                 # Priority
        "10": {"value": "General Support"},    # Category
        "11": {"value": "New"},                # Status
    }]
}

print(f"\nSending to: https://api.quickbase.com/v1/records")
print(f"Payload:\n{json.dumps(record_data, indent=2)}")

response = requests.post(
    "https://api.quickbase.com/v1/records",
    headers=headers,
    json=record_data
)

print(f"\nStatus: {response.status_code}")
print(f"Response:\n{json.dumps(response.json(), indent=2)}")

if response.status_code in [200, 201]:
    print("\n✅ SUCCESS!")
else:
    print("\n❌ FAILED - Check field IDs above and update the record_data")
    print("\nCommon issues:")
    print("  - Field IDs don't match your table")
    print("  - Required field missing")
    print("  - Field type mismatch (e.g., sending string to number field)")
    print("  - Ticket number might be auto-generated (don't send it)")