# NHT Operations Dashboard V5

## What changed
- Supabase email/password authentication retained.
- NHT summary data is now stored centrally in Supabase instead of browser localStorage.
- Users no longer need to upload the Excel every time they open the dashboard.
- Uploading the latest Excel replaces the shared dashboard dataset.
- Supabase Realtime refreshes connected dashboards when the shared data changes.
- Monthly total / grand total rows are excluded.
- Month and location filters rebuild all charts consistently.
- Throughput and Joining Throughput are displayed as percentages correctly.
- Joining Throughput is calculated using the workbook's source Joining Throughput values weighted by inflow.

## One-time Supabase database setup
1. Open Supabase -> SQL Editor for the NHT Operations Dashboard project.
2. Run the SQL in `supabase_setup.sql`.
3. Then upload the four web files to the GitHub Pages repository.

## Important
- The raw Excel file is processed in the browser; it is not uploaded to GitHub.
- The processed summary rows are stored in the private Supabase database and are available only to authenticated users through RLS.
- Current V4 allows any authenticated dashboard user to upload/replace the shared dataset. Admin-only upload can be added later.


## Branding
- Hi-Tek Syndicate logo is included locally as `hi-tek-syndicate-logo.png`.
- Dashboard and login use a Hi-Tek-inspired navy, blue, red and white palette.
- Chart series use the same company-inspired palette.
- No external logo image URL is required.
