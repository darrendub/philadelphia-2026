# Philadelphia 2026: setup

The site is plain files on GitHub Pages. The Google Sheet is the database, and a small Apps Script lets the site read and save to it.

## 1. Connect the Google Sheet (about 5 minutes, easiest on a laptop)

1. Open the trip Google Sheet.
2. Choose **Extensions > Apps Script**.
3. Delete the starter code and paste in everything from `apps-script/Code.gs`. Save.
4. Choose **Deploy > New deployment**. Click the gear and pick **Web app**.
5. Set **Execute as** to **Me** and **Who has access** to **Anyone**. Click **Deploy**.
6. Approve the permission prompts. If Google says the app is not verified, choose **Advanced**, then **Go to project**. It is your own script.
7. Copy the **Web app URL**. It ends in `/exec`.
8. Open `config.js` in this repo and paste the URL between the quotes. Commit the change.

Open the URL in a browser once. You should see JSON, and a new **Itinerary** tab appears in your sheet with the starting plan filled in.

## 2. Publish on GitHub Pages

1. In the repo, open **Settings > Pages**.
2. Under **Build and deployment**, set **Source** to **Deploy from a branch**, branch **main**, folder **/ (root)**. Save.
3. After a minute the site is live at `https://YOUR-USERNAME.github.io/REPO-NAME/`.

## 3. Add it to each iPhone

Open the site in Safari, tap **Share**, then **Add to Home Screen**. It opens full screen like an app and keeps the last copy of the plan if the signal drops.

## Notes

- Family can also edit rows directly in the **Itinerary** tab. Keep the column headers in row 1. A row without an id gets one automatically.
- The site checks the sheet every 20 seconds and right after anyone saves.
- Anyone with the site link can edit the plan. Share the link only with family.
- If you change `Code.gs` later, use **Deploy > Manage deployments > Edit > New version** so the URL stays the same.
