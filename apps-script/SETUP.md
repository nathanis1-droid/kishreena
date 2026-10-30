# Setting up the Kishreena session log (about 15 minutes)

Do this signed in to **kishreenafoundation@gmail.com**, not a school account. School accounts often block public web apps, and the data should belong to Kishreena.

1. Go to **sheets.new**. Name the sheet **Kishreena Session Log**.
2. Click **Extensions → Apps Script**. Delete the code in the editor, paste in everything from `Code.gs`, and click **Save**.
3. In the dropdown next to **Run**, choose **setup**, then click **Run**. Google asks for permission:
   **Review permissions → choose your account → Advanced → Go to (project) → Allow.** This is normal for a script you wrote yourself.
   Your sheet now has four tabs: Mentors, Logs, Hours, Pairs.
4. In the **Mentors** tab, add one row per mentor: email, name (e.g. "Dhriti N."), a 4-digit code, and **Y**.
5. Back in Apps Script, click **Deploy → New deployment**. Click the gear icon, choose **Web app**, and set:
   - Execute as: **Me**
   - Who has access: **Anyone**
   Click **Deploy**, then copy the **Web app URL** (it ends in `/exec`).
6. Send that URL to Claude. It goes into `SCRIPT_URL` at the bottom of `mentors.html`.

**Giving mentors access:** email each mentor their personal code, with a link to kishreena.org/mentors.html#log.

**If you change Code.gs later:** Deploy → Manage deployments → edit (pencil) → Version: New version → Deploy. The URL stays the same.
