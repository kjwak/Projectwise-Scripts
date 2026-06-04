-- Requires QUOTED_IDENTIFIER ON for indexed tables (pw_folder_cache / audit_events).
SET QUOTED_IDENTIFIER ON;
GO

DELETE FROM pw_folder_cache WHERE resolve_failed = 1;
GO

SELECT COUNT(*) AS remaining_failed FROM pw_folder_cache WHERE resolve_failed = 1;
GO
