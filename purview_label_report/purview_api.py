import requests
import msal

# Azure AD app registration details (replace with yours)
CLIENT_ID = "YOUR_CLIENT_ID"
TENANT_ID = "YOUR_TENANT_ID"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = [
    "https://graph.microsoft.com/.default",
    "User.Read",
    "Group.Read.All",
    "Directory.Read.All",
    "InformationProtectionPolicy.Read.All"
]

def get_token():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception("Failed to create device flow")
    print(flow["message"])  # Shows prompt to user
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception("Failed to obtain token: " + str(result))

def get_labels_with_policies():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    # 1. Get all sensitivity labels
    label_url = "https://graph.microsoft.com/v1.0/informationProtection/sensitivityLabels"
    response = requests.get(label_url, headers=headers)
    labels = response.json().get("value", [])
    output = []
    for label in labels:
        label_name = label.get("name")
        description = label.get("description", "")
        # 2. Get the publishing policy (scoping groups/units)
        policies = []
        for scope in label.get("policy", {}).get("scopes", []):
            if scope["type"] == "group":
                group_id = scope["id"]
                group_url = f"https://graph.microsoft.com/v1.0/groups/{group_id}"
                group = requests.get(group_url, headers=headers).json()
                count_url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/members/$count"
                user_count = requests.get(count_url, headers=headers).text
                policies.append({
                    "type": "Group",
                    "name": group.get("displayName", group_id),
                    "user_count": user_count
                })
            elif scope["type"] == "administrativeUnit":
                unit_id = scope["id"]
                unit_url = f"https://graph.microsoft.com/v1.0/administrativeUnits/{unit_id}"
                unit = requests.get(unit_url, headers=headers).json()
                count_url = f"https://graph.microsoft.com/v1.0/administrativeUnits/{unit_id}/members/$count"
                user_count = requests.get(count_url, headers=headers).text
                policies.append({
                    "type": "Administrative Unit",
                    "name": unit.get("displayName", unit_id),
                    "user_count": user_count
                })
        output.append({
            "label_name": label_name,
            "description": description,
            "policies": policies
        })
    return output
