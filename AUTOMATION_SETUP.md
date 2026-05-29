# BU Dashboard Data Automation - Setup

This repo regenerates `bu_dashboard_data.json` automatically from the source
Excel workbook on the PCMA Team SharePoint site, using a GitHub Action that
runs every 30 minutes (and on demand).

```
Excel on Team site  ->  GitHub Action (Python + openpyxl)  ->  bu_dashboard_data.json  ->  dashboard polls every 30s
```

Why a GitHub Action and not Power Automate: the dashboard groups business units
by the FILL COLOR of header rows. Power Automate's Excel connector cannot read
cell fill colors; openpyxl can. So this is the only approach that reproduces the
grouping correctly.

## Files

- `.github/workflows/update-bu-data.yml` - the scheduled workflow
- `scripts/build_bu_data.py` - downloads the workbook and builds the JSON
- `scripts/requirements.txt` - Python dependencies

## One-time setup

### 1. Source file

The source workbook is `2026 Business Unit Numbers.xlsx` on the PCMA Team site
(currently in the `2026-Business-Unit-Dashboard` library folder, sheet
"Business Unit"). The pipeline addresses it by its STABLE item id, so it can be
moved between folders without breaking. This is the file to edit going forward;
edits to any old OneDrive copy will NOT reach the dashboard.

- Drive id: `b!_609OR8fLU6wjK_Inp1_4O2-L5JDgrxHumtyBqVrpD8I85ePiYNhQITFvKReaukz`
- Item id: `01WUVNJBVHHYXW4LBSZRGYE3CTJMQBZ33J`

### 2. Create an Entra app registration

In Entra admin center (entra.microsoft.com) > App registrations > New
registration:

- Name: `PCMA-BU-Dashboard-Reader`
- Supported account types: single tenant
- No redirect URI needed

Record the Application (client) ID and Directory (tenant) ID.
Tenant ID for PCMA is `758ab235-b480-4c59-8e56-30144f4893ce`.

### 3. Grant Graph permission (least privilege: Sites.Selected)

In the app > API permissions > Add a permission > Microsoft Graph >
Application permissions > add `Sites.Selected` > Grant admin consent.

Then authorize this app to read ONLY the PCMA Team site. Run once (Graph
Explorer or PowerShell) as an admin:

```
POST https://graph.microsoft.com/v1.0/sites/393dadff-1f1f-4e2d-b08c-afc89e9d7fe0/permissions
{
  "roles": ["read"],
  "grantedToIdentities": [
    { "application": { "id": "<APPLICATION_CLIENT_ID>", "displayName": "PCMA-BU-Dashboard-Reader" } }
  ]
}
```

Simpler but broader fallback: grant `Sites.Read.All` (application) with admin
consent and skip the site-scoping call. Sites.Selected is recommended given
PCMA's data-classification posture.

### 4. Create the client secret and add repo secrets

App > Certificates & secrets > New client secret. Copy the VALUE immediately.

In GitHub: repo Settings > Secrets and variables > Actions > New repository
secret, add all three:

- `AZURE_TENANT_ID` = `758ab235-b480-4c59-8e56-30144f4893ce`
- `AZURE_CLIENT_ID` = the application (client) id
- `AZURE_CLIENT_SECRET` = the secret value

(The workflow commits with the built-in `GITHUB_TOKEN`, so no GitHub PAT is
needed.)

### 5. Parser calibration  (DONE)

The CONFIG block in `scripts/build_bu_data.py` (sheet name, columns, and
`COLOR_MAP` of fill color -> group) is calibrated and validated against the live
workbook: 4 groups, 24 subgroups, 68 items. If the workbook's brand fill colors
or column order change, update only that CONFIG block. To re-validate locally
without Graph: `BU_LOCAL_XLSX=/path/to/file.xlsx python scripts/build_bu_data.py`.

### 6. Run it

Actions tab > "Update BU Dashboard Data" > Run workflow. Confirm a commit to
`bu_dashboard_data.json` appears and the dashboard's "updated" timestamp moves.

## Notes

- Secret expiry: client secrets expire (max 24 months). Set a calendar reminder
  to rotate `AZURE_CLIENT_SECRET` before it lapses, or the flow goes silent.
- The script refuses to overwrite the JSON if parsing yields zero groups, so a
  transient read error cannot blank the dashboard.
- Scheduled GitHub Actions can be delayed a few minutes under load; the 30s
  dashboard polling still picks up changes promptly once committed.
