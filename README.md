# Fantasy Lighting Time Card

A time card app for one shared iPad, in English or Spanish. People tap their name, enter their PIN, then enter the date, start time, and end time. Hours are saved in a Google Sheet.

Breaks and lunch are paid, so every minute between the start and end time counts. Nothing is subtracted.

## Files

- `index.html`: the app
- `icon.png`: the home screen icon
- `google-sheet-script.gs`: the small program that goes inside the Google Sheet

## Set up the Google Sheet (once)

Google's screens may look a little different from these steps.

1. Sign in to the Google account that should own the hours, and go to [sheets.new](https://sheets.new) to make a new Sheet. Name it "Fantasy Lighting Time Card".
2. In the Sheet, click **Extensions**, then **Apps Script**.
3. Delete what's in the editor. Paste in everything from `google-sheet-script.gs`, then click the save icon.
4. At the top, choose `setup` from the function list and click **Run**.
5. Google asks for permission. Your own script hasn't been reviewed by Google, so it shows a warning. Click **Advanced**, then **Go to** your project, then **Allow**.
6. Click **Deploy**, then **New deployment**. Click the gear icon and choose **Web app**.
7. Set **Execute as** to **Me**, and **Who has access** to **Anyone**. Click **Deploy**.
8. Copy the **Web app URL**. It ends in `/exec`. This address goes into `index.html` as `SCRIPT_URL`.

"Anyone" lets the iPad reach the Sheet without a Google sign-in. Nothing can be read or changed without the access code that step 4 created.

## Add names and PINs

Open the **People** tab of the Sheet. Put each person on their own row, with their name in the first column and a 4-digit PIN in the second. Anyone who can open this Sheet can see the PINs, so don't share it widely, and don't use a PIN someone uses for their bank.

To remove someone, delete their row. Their past hours stay on the Hours tab.

## Connect the iPad

1. Put the app online so it has a web address, then open that address in Safari.
2. Tap Share, then **Add to Home Screen**.
3. Open the app from that home screen icon.
4. Enter the access code from the **Setup** tab of the Sheet. This only happens once.

## Fixing mistakes

- People can edit or delete their own entries from **Weekly hours** using their PIN.
- A manager can change anything directly on the **Hours** tab of the Sheet. The app picks up the change.
- If a row can't be read, for example because of a mistyped time, **Weekly hours** warns you which row it is. That row isn't counted until it's fixed.

## Each season

At the end of each season, download a copy of the Sheet using **File**, then **Download**, then **Microsoft Excel**.

## If the script is changed later

After pasting in a new version of `google-sheet-script.gs`, click **Deploy**, then **Manage deployments**, then the pencil icon. Choose **New version** and click **Deploy**. Without this, the old version keeps running.
