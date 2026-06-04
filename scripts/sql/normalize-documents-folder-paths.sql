-- Canonical folder paths: lowercase backslashes, leading documents\
-- Views v_qc_cycle_aging and v_folder_activity read base tables; fix tables only.
-- Requires QUOTED_IDENTIFIER ON on some installs.
SET QUOTED_IDENTIFIER ON;
GO

/* ---- processing_jobs ---- */
UPDATE processing_jobs SET source_folder = LOWER(REPLACE(LTRIM(RTRIM(source_folder)), '/', '\')) WHERE source_folder IS NOT NULL AND LTRIM(RTRIM(source_folder)) <> '';
GO
UPDATE processing_jobs SET source_folder = 'documents\' + source_folder WHERE source_folder IS NOT NULL AND LTRIM(RTRIM(source_folder)) <> '' AND source_folder NOT LIKE 'documents\%' AND source_folder NOT LIKE 'pw:\%';
GO
UPDATE processing_jobs SET source_folder = SUBSTRING(source_folder, CHARINDEX('\documents\', LOWER(source_folder)) + 1, 8000) WHERE source_folder LIKE 'pw:\%' AND CHARINDEX('\documents\', LOWER(source_folder)) > 0;
GO
UPDATE processing_jobs SET source_path = LOWER(REPLACE(LTRIM(RTRIM(source_path)), '/', '\')) WHERE source_path IS NOT NULL AND LTRIM(RTRIM(source_path)) <> '';
GO
UPDATE processing_jobs SET source_path = 'documents\' + source_path WHERE source_path IS NOT NULL AND LTRIM(RTRIM(source_path)) <> '' AND source_path NOT LIKE 'documents\%' AND source_path NOT LIKE 'pw:\%';
GO
UPDATE processing_jobs SET source_path = SUBSTRING(source_path, CHARINDEX('\documents\', LOWER(source_path)) + 1, 8000) WHERE source_path LIKE 'pw:\%' AND CHARINDEX('\documents\', LOWER(source_path)) > 0;
GO

/* ---- document_state_history (feeds v_qc_cycle_aging) ---- */
UPDATE document_state_history SET folder_path = LOWER(REPLACE(LTRIM(RTRIM(folder_path)), '/', '\')) WHERE folder_path IS NOT NULL AND LTRIM(RTRIM(folder_path)) <> '';
GO
UPDATE document_state_history SET folder_path = 'documents\' + folder_path WHERE folder_path IS NOT NULL AND LTRIM(RTRIM(folder_path)) <> '' AND folder_path NOT LIKE 'documents\%' AND folder_path NOT LIKE 'pw:\%';
GO
UPDATE document_state_history SET folder_path = SUBSTRING(folder_path, CHARINDEX('\documents\', LOWER(folder_path)) + 1, 8000) WHERE folder_path LIKE 'pw:\%' AND CHARINDEX('\documents\', LOWER(folder_path)) > 0;
GO

/* ---- transition_events ---- */
UPDATE transition_events SET folder_path = LOWER(REPLACE(LTRIM(RTRIM(folder_path)), '/', '\')) WHERE folder_path IS NOT NULL AND LTRIM(RTRIM(folder_path)) <> '';
GO
UPDATE transition_events SET folder_path = 'documents\' + folder_path WHERE folder_path IS NOT NULL AND LTRIM(RTRIM(folder_path)) <> '' AND folder_path NOT LIKE 'documents\%' AND folder_path NOT LIKE 'pw:\%';
GO
UPDATE transition_events SET folder_path = SUBSTRING(folder_path, CHARINDEX('\documents\', LOWER(folder_path)) + 1, 8000) WHERE folder_path LIKE 'pw:\%' AND CHARINDEX('\documents\', LOWER(folder_path)) > 0;
GO

/* ---- notification_log ---- */
UPDATE notification_log SET folder_path = LOWER(REPLACE(LTRIM(RTRIM(folder_path)), '/', '\')) WHERE folder_path IS NOT NULL AND LTRIM(RTRIM(folder_path)) <> '';
GO
UPDATE notification_log SET folder_path = 'documents\' + folder_path WHERE folder_path IS NOT NULL AND LTRIM(RTRIM(folder_path)) <> '' AND folder_path NOT LIKE 'documents\%' AND folder_path NOT LIKE 'pw:\%';
GO
UPDATE notification_log SET folder_path = SUBSTRING(folder_path, CHARINDEX('\documents\', LOWER(folder_path)) + 1, 8000) WHERE folder_path LIKE 'pw:\%' AND CHARINDEX('\documents\', LOWER(folder_path)) > 0;
GO

/* ---- audit_events (feeds v_folder_activity) ---- */
UPDATE audit_events SET resolved_folder = LOWER(REPLACE(LTRIM(RTRIM(resolved_folder)), '/', '\')) WHERE resolved_folder IS NOT NULL AND LTRIM(RTRIM(resolved_folder)) <> '';
GO
UPDATE audit_events SET resolved_folder = 'documents\' + resolved_folder WHERE resolved_folder IS NOT NULL AND LTRIM(RTRIM(resolved_folder)) <> '' AND resolved_folder NOT LIKE 'documents\%' AND resolved_folder NOT LIKE 'pw:\%';
GO
UPDATE audit_events SET resolved_folder = SUBSTRING(resolved_folder, CHARINDEX('\documents\', LOWER(resolved_folder)) + 1, 8000) WHERE resolved_folder LIKE 'pw:\%' AND CHARINDEX('\documents\', LOWER(resolved_folder)) > 0;
GO

/* ---- sheet_index, document_activity (dashboards / cache) ---- */
UPDATE sheet_index SET folder_path = LOWER(REPLACE(LTRIM(RTRIM(folder_path)), '/', '\')) WHERE folder_path IS NOT NULL AND LTRIM(RTRIM(folder_path)) <> '';
GO
UPDATE sheet_index SET folder_path = 'documents\' + folder_path WHERE folder_path IS NOT NULL AND LTRIM(RTRIM(folder_path)) <> '' AND folder_path NOT LIKE 'documents\%' AND folder_path NOT LIKE 'pw:\%';
GO

UPDATE document_activity SET folder_path = LOWER(REPLACE(LTRIM(RTRIM(folder_path)), '/', '\')) WHERE folder_path IS NOT NULL AND LTRIM(RTRIM(folder_path)) <> '';
GO
UPDATE document_activity SET folder_path = 'documents\' + folder_path WHERE folder_path IS NOT NULL AND LTRIM(RTRIM(folder_path)) <> '' AND folder_path NOT LIKE 'documents\%' AND folder_path NOT LIKE 'pw:\%';
GO

/* ---- collapse documents\documents\ ---- */
DECLARE @t TABLE (n SYSNAME, c SYSNAME);
INSERT INTO @t VALUES
 ('processing_jobs','source_folder'),('processing_jobs','source_path'),
 ('document_state_history','folder_path'),('transition_events','folder_path'),
 ('notification_log','folder_path'),('audit_events','resolved_folder'),
 ('sheet_index','folder_path'),('document_activity','folder_path');

DECLARE @sql NVARCHAR(MAX), @tbl SYSNAME, @col SYSNAME;
DECLARE cur CURSOR LOCAL FAST_FORWARD FOR SELECT n, c FROM @t;
OPEN cur;
FETCH NEXT FROM cur INTO @tbl, @col;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = N'WHILE EXISTS (SELECT 1 FROM ' + QUOTENAME(@tbl) + N' WHERE ' + QUOTENAME(@col) + N' LIKE ''documents\documents\%'')
BEGIN
    UPDATE ' + QUOTENAME(@tbl) + N' SET ' + QUOTENAME(@col) + N' = STUFF(' + QUOTENAME(@col) + N', 1, LEN(''documents\''), '''')
    WHERE ' + QUOTENAME(@col) + N' LIKE ''documents\documents\%'';
END';
    EXEC sp_executesql @sql;
    FETCH NEXT FROM cur INTO @tbl, @col;
END
CLOSE cur; DEALLOCATE cur;
GO

/* ---- verification ---- */
SELECT 'processing_jobs.source_folder' AS metric,
    SUM(CASE WHEN source_folder IS NOT NULL AND source_folder <> '' AND source_folder NOT LIKE 'documents\%' THEN 1 ELSE 0 END) AS not_canonical,
    COUNT(*) AS total FROM processing_jobs;
SELECT 'document_state_history.folder_path' AS metric,
    SUM(CASE WHEN folder_path IS NOT NULL AND folder_path <> '' AND folder_path NOT LIKE 'documents\%' THEN 1 ELSE 0 END) AS not_canonical,
    COUNT(*) AS total FROM document_state_history;
SELECT 'transition_events.folder_path' AS metric,
    SUM(CASE WHEN folder_path IS NOT NULL AND folder_path <> '' AND folder_path NOT LIKE 'documents\%' THEN 1 ELSE 0 END) AS not_canonical,
    COUNT(*) AS total FROM transition_events;
SELECT 'notification_log.folder_path' AS metric,
    SUM(CASE WHEN folder_path IS NOT NULL AND folder_path <> '' AND folder_path NOT LIKE 'documents\%' THEN 1 ELSE 0 END) AS not_canonical,
    COUNT(*) AS total FROM notification_log;
SELECT 'audit_events.resolved_folder' AS metric,
    SUM(CASE WHEN resolved_folder IS NOT NULL AND resolved_folder <> '' AND resolved_folder NOT LIKE 'documents\%' THEN 1 ELSE 0 END) AS not_canonical,
    COUNT(*) AS total FROM audit_events;
GO
