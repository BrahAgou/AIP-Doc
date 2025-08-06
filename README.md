# Purview Label PowerPoint Exporter

Exports Microsoft Purview sensitivity labels, publishing policies, and user counts to a styled PowerPoint (.pptx) file.

## Features

- Interactive authentication (OAuth consent popup).
- Retrieves label, policy, and user count data from Microsoft Graph API.
- Advanced slide styling: color, fonts, bullet lists, etc.

## Setup

1. Register an Azure AD App in your tenant for Graph API access.
2. Paste your `CLIENT_ID` and `TENANT_ID` in `purview_api.py`.
3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
4. Run:
   ```
   python main.py
   ```

## Permissions

App must have:
- InformationProtectionPolicy.Read.All
- Group.Read.All
- Directory.Read.All
- User.Read

## Output

- One slide per label, listing description and publishing policy (groups/units and user counts).

## Notes

- The structure of label policies may vary by tenant/configuration. You may need to adjust the API call in `purview_api.py` to reflect your setup.
- For more advanced layouts (tables, icons, branding), see [python-pptx documentation](https://python-pptx.readthedocs.io/en/latest/).
