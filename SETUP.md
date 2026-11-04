# Shift Report app: setup guide (about 30 minutes)

You will do 3 things: (A) prepare the Google Sheet, (B) put the app online, (C) install it on the supervisors' phones.

## A. Google Sheet and backend

1. Go to **sheets.google.com** and create a blank spreadsheet. Name it **Shift Reports**.
2. Open **File > Settings** and set the **Time zone** to your country (for example Nairobi). Save.
3. Open **Extensions > Apps Script**. Delete everything in the editor, then paste the full contents of **Code.gs** (from the `apps-script` folder). Click **Save**.
4. In the function drop-down at the top, choose **setupSheets** and click **Run**. Google asks for permission: choose your account, **Advanced > Go to project (unsafe) > Allow**. This is normal for your own script.
5. Go back to the spreadsheet. It now has 3 tabs:
   - **Users**: one row per supervisor: Name, PIN (4 to 6 digits), Active (Yes/No). Replace the sample rows with your 8 supervisors.
   - **Equipment**: one row per machine: Equipment, Area, Supervisor. The Supervisor must match a Name in Users exactly. If two supervisors share a machine, write both names separated by a comma.
   - **Jobs**: filled by the app. Do not type here.
6. In Apps Script click **Deploy > New deployment**, click the gear and choose **Web app**. Set **Execute as: Me** and **Who has access: Anyone**. Click **Deploy** and **copy the Web app URL** (it ends with `/exec`).

Changing a PIN or adding a machine later: just edit the Users or Equipment tab. No redeploy needed.

## B. Put the app online

1. Open **config.js** in a text editor and replace `PASTE_YOUR_WEB_APP_URL_HERE` with the URL you copied. Save.
2. Go to **app.netlify.com/drop**, sign up for a free account, and drag the whole app folder (the one containing `index.html`) onto the page. Netlify gives you a link such as `https://something.netlify.app`.
3. Open that link in Chrome on your own phone and sign in with one of the supervisor names and PINs to test.

## C. Install on the supervisors' phones

1. Send the link to each supervisor on WhatsApp, with their name and PIN.
2. On Android open it in **Chrome**, tap the **three dots menu > Install app** (or **Add to Home screen**).
3. Sign in once while there is internet. After that the app opens without signal.

## How supervisors use it

- **Report tab:** choose the date and shift, tap **Add job** for each breakdown, maintenance or concern, then **Submit shift**. With no signal it saves on the phone and uploads by itself when the signal returns (watch the badge on the Sent tab).
- **My machines tab:** each machine shows open backlog, breakdowns and downtime for the last 30 or 60 days. Tap a machine to see the jobs, and tap **Update status** on a backlog job when it is finished or waiting for parts.
- **Sent tab:** shows what is uploaded and what is still waiting.

## Your analysis

The **Jobs** tab in the Google Sheet has one row per job, with the Date as a real date, so pivot tables work. To use Excel: **File > Download > Microsoft Excel (.xlsx)**. You can also add charts directly in Google Sheets.

## Updating the app later

If you change `index.html`, also change the `VERSION` text at the top of `sw.js` (for example `shift-report-v2`) before uploading, so phones pick up the new version. If you change **Code.gs**, use **Deploy > Manage deployments > edit > New version > Deploy**; the URL stays the same.

## Good to know

- PINs are stored in plain text in your Users tab, so keep the sheet private (do not share it with the supervisors).
- Five wrong PIN attempts lock that name for 10 minutes.
- If a supervisor loses the phone, set **Active** to **No** for them in the Users tab and they are signed out at the next upload.
- Photos are not included in this version.
