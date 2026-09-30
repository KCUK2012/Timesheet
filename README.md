# KC Timesheet: Outlook add-in, version 1.1 (Option D)

Updated 30 September 2026. Draft for IT review.

In version 1.1 the client list is no longer published with the add-in. It is held in a SharePoint list and read with each member of staff's own Microsoft 365 sign-in, so only people who already have access to that list can see it. The files in this folder contain **no client names** and are safe to host publicly.

## What changed from version 1.0

| | Version 1.0 (pilot) | Version 1.1 |
|---|---|---|
| Client list | `clients.json`, publicly readable | SharePoint list, read with the user's own sign-in |
| Sign-in | None | Silent, using the Outlook account (Microsoft nested app authentication) |
| Updating clients | Edit a file on GitHub | Edit the SharePoint list |
| New files | | `config.js`, `data.js`, `redirect.html`, `lib/msal-browser.min.js` (Microsoft's sign-in library, version 4.30.0, MIT licence) |

## Who does what

| Part | Who | Needs |
|---|---|---|
| A. Remove the public pilot | Repository owner | GitHub access |
| B. Create the SharePoint list | Client list owner | Owner rights on the chosen SharePoint site |
| C. Register the app in Entra ID | Microsoft 365 administrator | Rights to create app registrations and grant admin consent |
| D. Configure and host the files | Repository owner or IT | GitHub organisation or Azure access |
| E. Test | Pilot users | Classic Outlook version 2409 or later on Microsoft 365 |
| F. Deploy | Microsoft 365 administrator | Microsoft 365 admin centre |

---

## Part A: Remove the public pilot

The version 1.0 repository contains the client list and its file history.

1. In the old repository, go to **Settings > Pages** and click **Unpublish site**.
2. Go to **Settings > General**, scroll to **Danger Zone** and choose **Delete this repository**.
3. In Outlook, remove the sideloaded version 1.0 add-in (**Get Add-ins > My add-ins**, then remove **KC Timesheet**).

Deleting only `clients.json` is not enough, because earlier versions stay in the repository history.

## Part B: Create the SharePoint list

1. Choose a SharePoint site that only Kaleidoscope staff can access, with no guest or external users.
2. On that site, choose **New > List > From CSV** and upload the import file supplied separately (`20260930_Timesheet clients SharePoint import_DPO.csv`). **Do not put that file in GitHub.**
3. Name the list, for example **Timesheet clients**.
4. Check the columns:
   - **Title**: the client name exactly as it should appear in the subject
   - **Aliases**: other recognised names, separated by semicolons (a single line or multiple lines of text column)
5. Under **List settings**, check the internal column names. If they are not `Title` and `Aliases`, change `nameColumn` and `aliasColumn` in `config.js` to match.
6. Keep the list's permissions inherited from the site, or restrict them to staff. Everyone who uses the add-in needs at least **Read** access.
7. Note the site address, for example `https://kaleidoscopeconsultants.sharepoint.com/sites/KCcorpserv/KCOps`. The part after `.com` is the **site path**.

## Part C: Register the app in Microsoft Entra ID

1. Sign in to the Microsoft Entra admin centre and go to **App registrations > New registration**.
2. **Name:** KC Timesheet add-in.
3. **Supported account types:** accounts in this organisational directory only (single tenant).
4. **Redirect URI:** choose **Single-page application (SPA)** and enter `brk-multihub://<add-in domain>`. Use the domain only, with no `https://` and no path. For example, if the add-in is hosted at `https://kc-timesheet.github.io/Timesheet`, enter `brk-multihub://kc-timesheet.github.io`.
5. Click **Register**.
6. Under **Authentication**, add a second **SPA** redirect URI: `https://<full hosting address>/redirect.html`. This is used only if silent sign-in is not available.
7. Under **API permissions**, choose **Add a permission > Microsoft Graph > Delegated permissions > Sites.Read.All**, then **Grant admin consent for Kaleidoscope Consultants**.
   - Because the permission is delegated, the add-in can only read what each person can already read in SharePoint.
   - A tighter alternative is **Sites.Selected** (delegated), limited to the one site. It needs an extra site-level grant, which IT can set up with Microsoft Graph or PowerShell.
8. From **Overview**, copy the **Directory (tenant) ID** and **Application (client) ID**.

These two IDs are identifiers, not secrets. It is normal for them to appear in a web add-in's code. No client secret is created or needed.

## Part D: Configure and host the files

1. Create a repository under a **firm-owned** GitHub organisation (GitHub Free is enough, because the files hold no client data). Alternatively, use Azure Static Web Apps (Free plan) in the firm's Azure subscription.
2. Open `config.js` and fill in the four **FILL IN** values:
   - `tenantId` and `clientId` from Part C
   - `sitePath` from Part B, for example `/sites/KCcorpserv/KCOps` (or `""` for the root site)
   - `listName`, for example `Timesheet clients`
3. Upload the contents of this folder to the top level of the repository, including the `lib` folder. Do **not** upload the CSV.
4. Go to **Settings > Pages**, set the source to the **main** branch and the **/ (root)** folder, and save. Leave **Custom domain** empty and tick **Enforce HTTPS**.
5. Open `https://<hosting address>/taskpane.html` in a browser to check it loads.
6. Edit `manifest.xml`:
   - replace every `https://REPLACE-WITH-YOUR-HOST` with the hosting address, with no trailing slash
   - change the **first** `<AppDomain>` to the domain only, for example `https://kc-timesheet.github.io`
   - leave the second `<AppDomain>` (`https://login.microsoftonline.com`) as it is
7. Check that the redirect URIs in Part C match the final hosting address.

## Part E: Test

1. In classic Outlook, check the version under **File > Office Account > About Outlook**. It must be **2409 or later** for silent sign-in.
2. Sideload the new `manifest.xml` (**Home > Get Add-ins > My add-ins > Add a custom add-in > Add from File**) and restart Outlook.
3. Create an appointment and click **Timesheet**. The panel should show "Loading client list from SharePoint...", then the full list.
4. Run the checks in the version 1.0 pilot checklist (search, apply, ACTUAL, pre-fill, ten-hour warning).
5. Extra checks for version 1.1:

| Test | Expected result |
|---|---|
| Add a test client to the SharePoint list, then reopen the panel | The new client appears |
| A pilot user without access to the list opens the panel | "You do not have access to the client list..." |
| Open `https://<hosting address>/` and browse the files | No client names appear anywhere |
| New Outlook and Outlook on the web | Same behaviour as classic Outlook |

## Part F: Deploy

1. In the Microsoft 365 admin centre, go to **Settings > Integrated apps**.
2. If version 1.0 was deployed, remove it, or use **Update** with the new manifest. The version number is now 1.1.0.0.
3. Upload the new `manifest.xml` and assign it to staff. It can take up to 24 hours to appear.

## Maintenance

| Change | What to do |
|---|---|
| Client added, renamed or removed | Edit the SharePoint list. Staff see the change the next time they open the panel. |
| Internal codes or hints change | Edit `config.js` in the repository |
| Hosting address changes | Update `manifest.xml` and both Entra redirect URIs, raise `<Version>`, then redeploy |
| Sign-in library update | Replace `lib/msal-browser.min.js` with a newer 4.x release after testing |

## Troubleshooting

| Message or symptom | Likely cause |
|---|---|
| "The add-in has not been set up yet..." | `config.js` still contains FILL IN values |
| "You do not have access to the client list..." | The user lacks Read access, or `sitePath` is wrong |
| 'The SharePoint list "..." was not found' | `listName` does not match the list's name |
| Sign-in window appears and fails | The redirect URIs in Entra do not match the hosting address, or admin consent was not granted |
| Panel blank in classic Outlook | Outlook is older than version 2409, or the hosting address in the manifest is wrong |

## Security summary

- **Mailbox access:** the add-in can only read and edit the open calendar item (`ReadWriteItem`).
- **Client list:** read at run time over HTTPS from SharePoint, using the user's own delegated access. It is never stored in the hosted files.
- **Credentials:** no passwords or secrets are held in the code. Sign-in is handled by Microsoft.
- **Data leaving Outlook:** none, other than the request to SharePoint for the client list.

## Testing done so far

- Code checked for syntax errors. Manifest checked as well-formed XML.
- Panel tested in a Chromium browser with **simulated** sign-in and SharePoint responses. The following worked: listing clients across more than one page of results, searching by alias, writing the subject, and the "no access" message.
- **Not yet tested** against real Microsoft Entra ID or SharePoint. Part E covers this.

## References

- Microsoft Learn, "Enable single sign-on in an Office Add-in with nested app authentication"
- Microsoft Learn, "Nested app auth requirement sets" (minimum Outlook versions)
- Microsoft Learn, Microsoft Graph: get a site by path, list items, and the Sites.Read.All and Sites.Selected permissions
