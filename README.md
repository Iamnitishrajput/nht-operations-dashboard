# NHT Operations Portal — V8

Hi-Tek Syndicate branded NHT Operations Dashboard.

## Included
- Supabase email/password authentication
- Role-based access: ADMIN / UPLOADER / VIEWER
- Central Supabase storage for NHT Excel data
- Month-scoped Excel replacement: uploading June again replaces June only
- Excel validation before any live-data replacement
- Monthly total rows excluded
- Throughput and Joining Throughput calculated correctly
- Automatic 5-minute inactivity sign-out
- PNG download button on every chart
- Five live data-analyst insight cards with animated rotation
- Hi-Tek Syndicate logo and corporate colour palette
- Responsive desktop/mobile layout
- Supabase Realtime refresh for shared data

## Required Supabase setup

The central data table should already exist from the earlier database setup.

For role-based upload permissions, run `nht_role_access_migration.sql` once in Supabase SQL Editor.

Then create users in Supabase Authentication > Users and assign each user a role in `public.nht_user_roles`:

- `admin` — can upload/update data
- `uploader` — can upload/update data
- `viewer` — view only

Example:

```sql
insert into public.nht_user_roles (user_id, role)
values ('USER-UUID-HERE', 'viewer')
on conflict (user_id) do update set role = excluded.role;
```

Do not put the Supabase secret key in the website. The publishable key in `script.js` is intended for browser use.

## GitHub Pages files

Upload these website files to the repository root:
- index.html
- style.css
- script.js
- hi-tek-syndicate-logo.png

The SQL files are setup/migration references and do not need to be served by GitHub Pages.

## Excel behaviour

A valid workbook should contain:
Month, Total Batch Conducted, Location, Total Inflow, Total Outflow,
HR Attrition, Training Attrition, Throughput, Total Joined, Joining Throughput.

If an upload contains only June, only June is replaced.
If it contains June, July and August, all three months are replaced.
If validation fails, the existing live dataset is kept.

The raw Excel file is processed in the browser; the workbook itself is not uploaded to GitHub.


## V12 changes
- Fixed role detection by using a secure `get_nht_role()` RPC, avoiding the frontend 403 caused by direct REST reads of `nht_user_roles`.
- Preserved RLS and role-based write controls.
- Rebuilt mobile layout from the supplied phone screenshots: stacked branding, full-width controls, responsive KPI cards, stable chart panels, and a stacked login brand.
- Added cache-busting query strings to CSS/JS so GitHub Pages does not keep serving the previous mobile CSS.
- Run the updated `nht_role_access_migration.sql` once in Supabase after deploying V12. Existing users and their assigned roles do not need to be recreated.


## V12 data logic correction
- Total Batches is treated as a monthly shared metric. One batch can contain employees from multiple locations, so batches are not summed across location rows.
- Location filtering no longer changes the monthly batch count.
- Throughput and joining throughput are calculated from the filtered location/month inflow, outflow and joined data.
- HR and training attrition charts use the filtered location/month attrition values directly.
