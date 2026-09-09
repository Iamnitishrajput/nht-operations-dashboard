NHT Operations Dashboard — Final Chart Fix

This patch keeps the existing Supabase authentication, roles, central data and dashboard design.

Fixes:
- Single selected month now renders throughput as visible bars instead of a one-point line.
- Single selected month now renders HR and training attrition as visible bars.
- Multi-month selections retain line trends.
- Merged Month cells in the source Excel remain supported through forward-fill validation.
- Existing Admin/Uploader/Viewer access is preserved.
