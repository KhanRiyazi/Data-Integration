# Backend setup: Google Sheet as your database (about 10 minutes)

Your page stays on GitHub Pages. A Google Sheet stores the data, and a small Apps Script web app
sits between them. Only you can write. Visitors can only read the projects you mark as published.

## 1. Create the Sheet and add the script
1. Create a new Google Sheet named **Portfolio Database**.
2. **Extensions > Apps Script**. Delete the default code and paste all of `Code.gs`.
3. **Project Settings** (gear icon) > tick **Show "appsscript.json" manifest file in editor**.
   Open `appsscript.json` in the editor and replace its contents with the file from this folder.
4. Choose the function **setup** in the toolbar and press **Run**. Approve the permissions.
   You now have the tabs `tasks`, `projects`, `kv` and a `dashboard` tab full of Sheets formulas.

## 2. Create a Google sign-in (OAuth client ID)
This is what keeps strangers out of your data.
1. Open https://console.cloud.google.com, create a project (any name).
2. **APIs & Services > OAuth consent screen**: user type External, fill the app name and your email.
   Add your own Google account under **Test users**. You do not need to publish the app.
3. **APIs & Services > Credentials > Create credentials > OAuth client ID > Web application**.
4. Under **Authorized JavaScript origins** add your site origin exactly, with no trailing slash:
   `https://khanriyazi.github.io` (and `https://khanriyazi.in` if you use that domain).
5. Copy the **Client ID**.

## 3. Tell the script who you are
In Apps Script: **Project Settings > Script properties > Add script property**
| Property | Value |
|---|---|
| `OWNER_EMAIL` | your Google account email |
| `CLIENT_ID` | the Client ID from step 2 |

## 4. Deploy the web app
**Deploy > New deployment > type: Web app**
- Execute as: **Me**
- Who has access: **Anyone**

Press **Deploy** and copy the **Web app URL** (it ends with `/exec`).
Test it: open `<URL>?action=ping` in a browser. You should see `{"ok":true,...}`.

After you change `Code.gs` later: **Deploy > Manage deployments > Edit > Version: New version > Deploy**.

## 5. Connect the page
In `index.html`, find `backend:` inside `CONFIG` and fill it in:
```js
backend: {
  url: "https://script.google.com/macros/s/XXXX/exec",
  clientId: "XXXX.apps.googleusercontent.com",
  sheetUrl: "https://docs.google.com/spreadsheets/d/XXXX/edit",
  autoSync: true
},
```
Commit the change to GitHub.

## 6. First sync
1. Open your page and press the **☁️** button in the top bar (Sync tab).
2. **Sign in with Google** using the account from `OWNER_EMAIL`.
3. Press **Add this page's projects to the sheet**, then **Sync now**. Open the Sheet and check the `projects` tab.
4. From now on, tasks, projects, learning progress, readiness and formula practice sync after every change.
   Manage projects in the **Projects** tab of the Workspace window (add, edit, delete, publish).

## How it works
- Everything saves on your device first, so the page works offline.
- Sync uses last-write-wins per item. Deletions are kept for 30 days so they reach every device.
- The sign-in token lasts about an hour. When it ends, sign in again.
- Every cell is stored as plain text, so nothing typed on the page can run as a spreadsheet formula.

## Troubleshooting
| Message | Cause and fix |
|---|---|
| Could not reach the backend | The URL is wrong, or the deployment is not set to **Anyone**. |
| Backend not configured | Add `OWNER_EMAIL` and `CLIENT_ID` in Script properties. |
| This Google account is not allowed | You signed in with an account that is not `OWNER_EMAIL`. |
| Sign in is not allowed for this app | `CLIENT_ID` in Script properties differs from `clientId` on the page. |
| Google button does not appear | The page origin is missing from Authorized JavaScript origins. |
| Missing tab 'tasks' | Run `setup` once from the Apps Script editor. |