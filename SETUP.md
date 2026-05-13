# LumpiaPOS Setup Guide

## 1. Google Sheets Setup

1. Create a new Google Spreadsheet for each branch
2. Copy the Spreadsheet ID from the URL: `https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit`
3. The system will auto-create all required tabs (Menu, Ingredients, Recipes, Orders, OrderItems, Ledger, StockLog, Settings) on first login

## 2. Google Apps Script Backend

1. Go to [script.google.com](https://script.google.com) and create a new project
2. Delete the default `Code.gs` content and paste the contents of `Code.gs` from this repo
3. Edit the `BRANCHES` config at the top of `Code.gs`:

```javascript
const BRANCHES = {
  'Lumpia Pusat': { password: 'your_password_here', sheetId: 'YOUR_SPREADSHEET_ID' },
  'Lumpia Cabang 1': { password: 'another_password', sheetId: 'ANOTHER_SPREADSHEET_ID' },
};
```

4. Deploy as Web App:
   - Click **Deploy → New deployment**
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy**
   - Copy the Web App URL (ends with `/exec`)

5. If you update `Code.gs`, create a **new deployment** (Deploy → Manage deployments → New version) for changes to take effect

## 3. Frontend Setup

### Option A: Open Locally
Simply open `index.html` in a browser. All features work locally.

### Option B: GitHub Pages

1. Create a GitHub repository
2. Push `index.html` to the repo
3. Create `.github/workflows/pages.yml`:

```yaml
name: Deploy to GitHub Pages
on:
  push:
    branches: [main]
permissions:
  contents: read
  pages: write
  id-token: write
jobs:
  deploy:
    runs-on: ubuntu-latest
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    steps:
      - uses: actions/checkout@v4
      - uses: actions/configure-pages@v4
      - uses: actions/upload-pages-artifact@v3
        with:
          path: '.'
      - id: deployment
        uses: actions/deploy-pages@v4
```

4. In repo Settings → Pages, set source to **GitHub Actions**
5. Push and wait for the workflow to complete

## 4. First Login

1. Open the app (local file or GitHub Pages URL)
2. Go to **Settings** and paste the Apps Script Web App URL
3. Select your branch and enter the password
4. Click **Sign In** — the app will sync with Google Sheets

## 5. Image Storage (Optional)

Menu item images can be stored on GitHub for cross-device access:

1. Create a GitHub Personal Access Token (PAT):
   - Go to GitHub → Settings → Developer settings → Personal access tokens → Fine-grained tokens
   - Create token with **Contents: Read and write** permission for your repo
2. In app **Settings**, enter the repo (`username/repo`) and token
3. Images upload to `images/` folder in the repo and are served via `raw.githubusercontent.com`

> **Note:** The GitHub token is stored locally per device and never synced to the cloud.

## 6. Branch Configuration

To add more branches:
1. Create a new Google Spreadsheet
2. Add the branch to `BRANCHES` in `Code.gs`
3. Redeploy the Apps Script web app
4. Add the branch name to the `branches` array in `initLogin()` in `index.html`

## 7. Architecture Notes

- **Local-first**: All data saves locally first for instant UX, then syncs to Google Sheets in the background
- **CORS workaround**: All API requests use GET with `?payload=` parameter (Apps Script doesn't handle OPTIONS preflight)
- **Timezone**: Frontend uses local date functions (not UTC) to avoid date-shift bugs in UTC+7
- **Order IDs**: Generated on frontend so both local and cloud share the same ID (critical for cancellation)
- **Images**: Base64 thumbnails stored in a dedicated localStorage key (`mb_menu_images_{branch}`), separate from menu data to prevent overwrite during sync

## 8. Payment Methods

| Method      | Pool  |
|-------------|-------|
| Cash        | Cash  |
| DANA        | QRIS  |
| GoFood      | QRIS  |
| ShopeeFood  | QRIS  |
| GoPay       | QRIS  |
| Grab        | QRIS  |

Each menu item can have different prices per payment method. Reports split totals by Cash vs QRIS.

## 9. Troubleshooting

- **Dates shifted by 1 day**: The backend converts Google Sheets Date objects using `Session.getScriptTimeZone()`. If dates are still wrong, check that the Apps Script project timezone is set correctly (File → Project settings).
- **Sync fails**: Check that the Web App URL ends with `/exec` (not `/dev`). After updating `Code.gs`, create a new deployment version.
- **Images not showing on other devices**: Ensure GitHub repo and token are configured. Images are stored locally as base64 but only URLs are synced via Google Sheets.
- **Double orders**: The app uses an `_ordering` flag and disables the button during processing. If this still happens, check network latency.
