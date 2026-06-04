-- Canonical folder paths: lowercase, forward slashes as backslashes, leading documents\
-- Run after SET QUOTED_IDENTIFIER ON (required for indexed tables on some installs).
SET QUOTED_IDENTIFIER ON;
GO

/* processing_jobs */
UPDATE processing_jobs
SET source_folder = LOWER(REPLACE(LTRIM(RTRIM(source_folder)), '/', '\'))
WHERE source_folder IS NOT NULL AND LTRIM(RTRIM(source_folder)) <> '';
GO

UPDATE processing_jobs
SET source_folder = 'documents\' + source_folder
WHERE source_folder IS NOT NULL
  AND LTRIM(RTRIM(source_folder)) <> ''
  AND source_folder NOT LIKE 'documents\%'
  AND source_folder NOT LIKE 'pw:\%';
GO

UPDATE processing_jobs
SET source_folder = SUBSTRING(source_folder, CHARINDEX('\documents\', LOWER(source_folder)) + 1, 8000)
WHERE source_folder LIKE 'pw:\%' AND CHARINDEX('\documents\', LOWER(source_folder)) > 0;
GO

UPDATE processing_jobs
SET source_path = LOWER(REPLACE(LTRIM(RTRIM(source_path)), '/', '\'))
WHERE source_path IS NOT NULL AND LTRIM(RTRIM(source_path)) <> '';
GO

UPDATE processing_jobs
SET source_path = 'documents\' + source_path
WHERE source_path IS NOT NULL
  AND LTRIM(RTRIM(source_path)) <> ''
  AND source_path NOT LIKE 'documents\%'
  AND source_path NOT LIKE 'pw:\%';
GO

UPDATE processing_jobs
SET source_path = SUBSTRING(source_path, CHARINDEX('\documents\', LOWER(source_path)) + 1, 8000)
WHERE source_path LIKE 'pw:\%' AND CHARINDEX('\documents\', LOWER(source_path)) > 0;
GO

WHILE EXISTS (SELECT 1 FROM processing_jobs WHERE source_folder LIKE 'documents\documents\%')
BEGIN
    UPDATE processing_jobs
    SET source_folder = STUFF(source_folder, 1, LEN('documents\'), '')
    WHERE source_folder LIKE 'documents\documents\%';
END
GO

SELECT 'processing_jobs' AS tbl,
       SUM(CASE WHEN source_folder IS NOT NULL AND source_folder <> '' AND source_folder NOT LIKE 'documents\%' THEN 1 ELSE 0 END) AS folders_not_canonical,
       COUNT(*) AS total_rows
FROM processing_jobs;
GO
