# NFPC Supply Chain System

HTML / SharePoint version of `NFPC Supply Chain Prog_2022.accdb`. It has the same layout, settings and engine as the RAB Sales & Expenses System and NFPC SC Mngt.
Every Access form is a page, every Access table is a SharePoint list, and every Access report is on the Reports page.

## Files
| File | What it is |
|---|---|
| `index.html` | The page: login, side panel and top bar |
| `css/app.css` | All styles: base, the SC Mngt theme layer and the NFPC additions |
| `js/00-config.js` | **Edit this.** Tenant ID, Client ID, Site URL, and the system name/logo for every user |
| `js/00b-nfpc-access-config.js` | Generated from the .accdb: 122 tables (lists and columns), 165 forms (pages, fields, subforms) and 80 reports |
| `js/01-core.js` | Helpers, and builds the SharePoint list schema from the config |
| `js/02-data.js` | Microsoft sign-in, Microsoft Graph, SharePoint lists, the demo data store, users (LoginUserList) and roles |
| `js/03-ui.js` | Form engine, grids, Excel import/export and charts |
| `js/04-logic.js` | Business rules (FEFO stock posting, line totals, numbering, LPO/repair/GL actions, alerts) and the demo data |
| `js/05-pages-main.js` | Dashboard, Stock Balance, LOT Expiry & FEFO, Fleet Compliance and Staff Document Expiry |
| `js/06-print.js` | Branded A4 documents: logo, company header, watermark and footer |
| `js/07-pages-access-forms.js` | One smart page per Access form, with subforms as editable line grids |
| `js/08-reports.js` | The Reports page (the 80 Access reports plus 5 built-in reports) |
| `js/09-admin-boot.js` | Settings, help/form map, search, switchboard and start-up |
| `.nojekyll` | Tells GitHub Pages to serve the files as they are |

Keep the folder structure as it is. `index.html` loads `css/` and `js/` by relative path, and the scripts must load in the order listed in `index.html`.

## SharePoint lists
- **8 existing lists** on `sites/nfpc_WLOps_APP`, the same ones SC Mngt uses: LoginUserList, Inventory_Item_List, Main_Inventory_Movements_tbl, UOM_tbl, Items_WHs_Location_tbl, Facilities_Work_Scop_tbl, Trls_Daily_Activity_tbl and Vehicles_Dtls_tbl.
  - The system reads and writes these lists as they are. Columns are matched by name and are never created or changed.
- **114 new lists** come from the Access back-end tables. They are created on the *New lists site* (blank means the global site) with **Settings → Create / repair all lists**, which needs the `Sites.Manage.All` permission.
- Sign-in uses the **LoginUserList** list: `User_Name` or `Email_ID`, plus `Password` and `Access_Level`.
  - Admin or Manager levels can open every page. Other levels can open every module except Admin.

## Publish on GitHub Pages
1. Create a repository in the **same GitHub account as SC Mngt** (for example `nfpc-sc`). Both systems then share the connection settings and the colour theme.
2. Upload everything in this folder, including `.nojekyll`, to the root of the repository.
3. Go to the repository's **Settings → Pages**. Set Source to *Deploy from a branch*, choose `main` / `(root)`, and click Save.
4. In Azure, go to App registrations → your app → **Authentication** → **Single-page application** and add:
   - `https://<account>.github.io/<repo>/`
   - `https://<account>.github.io/<repo>/index.html`

   Add them under Single-page application, not *Web*; under Web you get error AADSTS9002326.
5. Open the site, sign in and go to **Settings**. If SC Mngt is already configured on the same account, the IDs are already filled in.
6. First time only: click **Settings → Check lists**, then **Create / repair all lists**, then import your Access tables from Excel. Import header tables before their line tables.

## Run locally
Microsoft sign-in does not work from a double-clicked file. Instead, start a web server in this folder with `python -m http.server 8080` and open `http://localhost:8080/`.
If you open the page as a plain file, it works with demo data only.

## Demo sign-in
In demo mode (no SharePoint), sign in as `admin` / `admin`, or as `store` / `1234` for a normal user.
