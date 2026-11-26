# Rimons IELTS class mail

A private web page, made for the phone, that sends each class email to a whole batch from your own Gmail and keeps a record of every class. Free.

## What it does

**Send**
1. Pick the batch.
2. Pick **Live class** or **Recorded class**.
   - **Live class:** type the title, pick the day, time and length. When you tap **Open in Gmail**, the Zoom meeting is created on your Zoom account (it shows in the Zoom app on your Mac and phone) and its link, meeting ID and passcode go into the email.
   - **Recorded class:** tap the recording from your recent Zoom cloud recordings. Or tap **Paste a link instead** for a YouTube, Drive or other link.
   - Every class gets an automatic title from its batch and the next class number, such as **Batch 82 · Class 14** or **Batch 82 · Recording 6**. Add to it if you like (": Writing Task 2"); clearing it brings the automatic title back.
3. Check the email preview and tap **Open in Gmail**. The page saves the designed email as a **Gmail draft** with every student in BCC (nobody sees the other emails). Nothing is sent yet.
   - **On a computer:** Gmail opens in a new tab with that draft open. Check it and press **Send**.
   - **On a phone:** tap **Open Gmail app**, open **Drafts**, tap the email and press **Send**. (A phone link can open the Gmail app but not one particular draft, so the Drafts step stays.)

**If something goes wrong**
- Under every finished email there are three buttons: **Copy email design**, **Copy subject** and **Copy student emails**. Paste them into a new Gmail email (design into the message, subject into the subject, emails into **Bcc**) and send it by hand.
- If Gmail cannot save the draft, the page says so and shows those steps. The Zoom meeting and the class history are kept.
- If Zoom is not working, the page says why in plain words at the top of the class box. Make the meeting in the Zoom app, tap **Paste a Zoom invitation instead**, and paste the Zoom invitation. Everything else works the same.
- The banner in a copied design is loaded from the repository's `main` branch (`apps-script/assets/`), because Gmail drops pictures pasted as data. It shows once this folder is on `main`. Drafts made by the page carry the banner inside the email and do not depend on this.

**Batches**
- One list of all batches, most recently used first, with a search box. One **New batch** button, and **Details** and **Edit** on each batch.
- Details, Edit and New batch slide up from the bottom and cover most of the screen. The **×** at the top goes back one step (Edit → Details → closed); tapping the dimmed area closes it.
- On **Send**, the 6 most recently used batches show as buttons; search finds any other. The day is picked from a row of day cards (or **Other** for the phone's calendar), the time from one field: hour : minutes AM/PM.
- **Edit:** paste the emails in any form: one per line, commas, spaces, semicolons, copied from Gmail or WhatsApp, even stuck together with no space. While you paste, the page says how many emails it found and lists any it could not read, so you can fix them before saving. If a save fails, your text stays in the box. The saved list shows one email per line, and **Copy all emails** copies them with commas, ready for Gmail or anywhere else.
- **Details:** how many classes the batch has had, how many live and recorded, and the total time. Below that, every class with its date and length. For live classes the real length and how many joined come from Zoom after the class ends. Upcoming classes are listed but not counted yet.

Everything is stored in one Google Sheet the page makes in your Drive, called **Rimons IELTS class mail data**. You never need to open it, but it is there as a backup and you can look at it any time.

The code is in [`apps-script/`](apps-script). The old copy-paste pages (`index.html`, `live-*.html`, `record-*.html`) are no longer needed. Their images were hosted on `media-hosting.imagekit.io`, which no longer exists, so they show broken images now.

## One-time setup (about 15 minutes, easier on a computer)

Do this while signed in to the Gmail account the emails should come **from**.

1. Go to <https://script.google.com> and click **New project**. Name it `Class mail`.
2. Replace everything in `Code.gs` with the contents of [`apps-script/Code.gs`](apps-script/Code.gs).
3. Click **+** next to Files, choose **HTML**, name it `Index` (exactly), and paste in [`apps-script/Index.html`](apps-script/Index.html).
4. Click **+** next to Files again, choose **Script**, name it `Images` (exactly), and paste in [`apps-script/Images.gs`](apps-script/Images.gs). It is one long block of letters: that is the two banner pictures packed into the script, so they travel inside each email and never need hosting.
5. Click **Save**.
6. Click **Deploy → New deployment**. Click the gear and pick **Web app**.
   - Execute as: **Me**
   - Who has access: **Only myself**
   - Click **Deploy**.
7. Google asks for permission: create Gmail drafts as you, and make and edit the data sheet in your Drive. Click **Authorize access** and pick your account. Google shows "Google hasn't verified this app" because you wrote it yourself. Click **Advanced → Go to Class mail**, then **Allow**.
8. Copy the **Web app URL**. Open it on your phone (signed in to the same Google account) and use **Add to Home screen** so it opens like an app.

## Connect Zoom (one time, free)

Until this is done, both class types work by pasting the Zoom invitation or link. Do it while signed in to the Zoom account that hosts the classes.

1. Go to <https://marketplace.zoom.us> → **Develop → Build App** → **Server-to-Server OAuth App** → **Create**. Name it `Class mail`.
2. On **App Credentials**, copy the **Account ID**, **Client ID** and **Client Secret**.
3. On **Information**, fill in the required name and email fields.
4. On **Scopes**, click **Add Scopes** and tick these three (names can differ slightly on older accounts):
   - create meetings: `meeting:write:meeting:admin`
   - see how long a past meeting ran: `meeting:read:past_meeting:admin`
   - list cloud recordings: `cloud_recording:read:list_user_recordings:admin`
5. On **Activation**, click **Activate your app**.
6. In the Apps Script project, open **Project Settings** (gear) → **Script Properties** → **Add script property** four times:
   - `ZOOM_ACCOUNT_ID` = the Account ID
   - `ZOOM_CLIENT_ID` = the Client ID
   - `ZOOM_CLIENT_SECRET` = the Client Secret
   - `ZOOM_HOST_EMAIL` = the email you sign in to Zoom with (the account that hosts the classes). Zoom's rules for this kind of app need the host named here.
7. Reload the page. Live class now shows title, day, time and length, and Recorded class lists your recordings. If anything is wrong (a wrong key, a missing scope, no host email), an orange note at the top of the class box says exactly what to fix. The first use asks Google for one more permission ("connect to an external service"), which is the call to Zoom.

The keys stay in Script Properties, not in the code, so the code can be shared safely. Meeting settings (waiting room, recording, passcode rules) follow the Zoom account's own defaults. Times are Bangladesh time; change `TIMEZONE` and `UTC_OFFSET_MIN` at the top of `Code.gs` if that ever changes.

If the length or "joined" count of a live class stays empty, the past-meeting scope is missing. The class still counts, with its planned length.

Zoom limits to know: the recordings list needs a paid Zoom plan (Pro or higher) with cloud recording on; Zoom allows 100 meetings to be created per day per host.

**Not yet tried against a real Zoom account.** The calls follow Zoom's documentation, and every failure falls back to pasting the invitation. Do one test class first: pick a time, tap **Open in Gmail**, check the meeting appears in the Zoom app, then delete the draft and the test meeting.

## Limits

- The final Send happens in Gmail, so normal Gmail limits apply: about 500 recipients a day and 500 per email on a free Gmail account (2,000 on Google Workspace). The page itself sends nothing.
- The class is added to the batch's history when the draft is made. If you delete the draft instead of sending it, the class still shows in the history (and a live class's Zoom meeting still exists).
- A batch can hold about 1,700 emails.
- Recorded class lists cloud recordings from the last 30 days. Older ones: paste the link.
- Only classes sent from this page are tracked. A meeting made straight in Zoom is not linked to a batch.

## Updating the code later

Paste the new code into the same project, then go to **Deploy → Manage deployments → edit (pencil) → Version: New version → Deploy**. The URL stays the same, and your batches and class history stay in the sheet.

## Checking the code

In the Apps Script editor, pick `test_all` in the function menu and click **Run**. It should log `all checks passed`. It reads no data and sends nothing.
