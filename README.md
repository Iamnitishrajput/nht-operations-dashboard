# NHT Operations Dashboard V3

Secure NHT Operations Portal hosted on GitHub Pages.

## V3 features

- Supabase email/password authentication
- Persistent authenticated session
- Sign-out control
- Browser-only Excel processing
- Local persistence of the last processed Excel summary
- Month and location filters
- Monthly total rows such as `June Total`, `July Total`, etc. are excluded from the location-level dataset to prevent duplicate months and double-counting
- KPI cards and management charts
- Location summary table

## Deployment

Upload these files to the root of the GitHub Pages repository:

- `index.html`
- `style.css`
- `script.js`
- `README.md`

Do **not** upload the NHT Excel file to GitHub.

## Supabase configuration

The frontend uses the project's **publishable key**. Supabase documents publishable keys as safe to expose in browser applications; secret keys must never be placed in frontend code.

Authentication redirect URL:

`https://iamnitishrajput.github.io/nht-operations-dashboard/`

## Data flow

`User signs in → browser loads saved dashboard data → user uploads the latest Excel when needed → Excel is parsed locally → processed summary is stored in that browser's localStorage.`

The raw Excel file is not uploaded to Supabase or GitHub by this dashboard.
