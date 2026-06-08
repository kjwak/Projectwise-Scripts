-- Backfills qc_cycle_completions.sheet_package_id from sheet_documents.
-- Reports unmapped rows and duplicate package logical keys without deleting or merging data.
SET QUOTED_IDENTIFIER ON;
SET NOCOUNT ON;
GO

IF OBJECT_ID('dbo.qc_cycle_completions', 'U') IS NULL
BEGIN
    RAISERROR('qc_cycle_completions table does not exist.', 16, 1);
    RETURN;
END
GO

IF OBJECT_ID('dbo.sheet_documents', 'U') IS NULL
BEGIN
    RAISERROR('sheet_documents table does not exist. Run backfill-sheet-packages.sql first.', 16, 1);
    RETURN;
END
GO

IF OBJECT_ID('dbo.qc_cycle_completion_package_backfill_report', 'U') IS NULL
BEGIN
    CREATE TABLE qc_cycle_completion_package_backfill_report (
        id              BIGINT IDENTITY(1,1) PRIMARY KEY,
        detected_at     DATETIMEOFFSET(3) NOT NULL CONSTRAINT DF_qccc_pkg_bf_report_detected DEFAULT SYSDATETIMEOFFSET(),
        report_type     NVARCHAR(50) NOT NULL,
        completion_id   BIGINT NULL,
        document_guid   UNIQUEIDENTIFIER NULL,
        sheet_package_id UNIQUEIDENTIFIER NULL,
        qc_cycle_id     NVARCHAR(100) NULL,
        qc_review_type  NVARCHAR(100) NULL,
        details         NVARCHAR(MAX) NULL
    );
END
GO

TRUNCATE TABLE qc_cycle_completion_package_backfill_report;
GO

UPDATE c
SET c.sheet_package_id = sd.sheet_package_id
FROM qc_cycle_completions c
INNER JOIN sheet_documents sd
    ON c.document_guid = sd.document_guid
WHERE c.sheet_package_id IS NULL;
GO

INSERT INTO qc_cycle_completion_package_backfill_report
    (report_type, completion_id, document_guid, sheet_package_id, qc_cycle_id, qc_review_type, details)
SELECT
    'unmapped_completion',
    c.id,
    c.document_guid,
    c.sheet_package_id,
    c.qc_cycle_id,
    c.qc_review_type,
    'document_guid not found in sheet_documents and sheet_package_id still NULL'
FROM qc_cycle_completions c
WHERE c.sheet_package_id IS NULL;
GO

INSERT INTO qc_cycle_completion_package_backfill_report
    (report_type, completion_id, document_guid, sheet_package_id, qc_cycle_id, qc_review_type, details)
SELECT
    'document_package_mismatch',
    c.id,
    c.document_guid,
    c.sheet_package_id,
    c.qc_cycle_id,
    c.qc_review_type,
    CONCAT('document_guid maps to package ', CONVERT(NVARCHAR(36), sd.sheet_package_id), ' but completion has ', CONVERT(NVARCHAR(36), c.sheet_package_id))
FROM qc_cycle_completions c
INNER JOIN sheet_documents sd ON c.document_guid = sd.document_guid
WHERE c.sheet_package_id IS NOT NULL
  AND c.sheet_package_id <> sd.sheet_package_id;
GO

INSERT INTO qc_cycle_completion_package_backfill_report
    (report_type, completion_id, document_guid, sheet_package_id, qc_cycle_id, qc_review_type, details)
SELECT
    'duplicate_package_logical_key',
    c.id,
    c.document_guid,
    c.sheet_package_id,
    c.qc_cycle_id,
    c.qc_review_type,
    CONCAT('duplicate count=', CAST(dup.dup_count AS NVARCHAR(20)))
FROM qc_cycle_completions c
INNER JOIN (
    SELECT sheet_package_id, qc_cycle_id, qc_review_type, COUNT(*) AS dup_count
    FROM qc_cycle_completions
    WHERE sheet_package_id IS NOT NULL
    GROUP BY sheet_package_id, qc_cycle_id, qc_review_type
    HAVING COUNT(*) > 1
) dup
    ON c.sheet_package_id = dup.sheet_package_id
   AND c.qc_cycle_id = dup.qc_cycle_id
   AND c.qc_review_type = dup.qc_review_type;
GO

-- Rebuild sheet_packages completion rollups from package-keyed completions.
;WITH bucketed AS (
    SELECT
        sheet_package_id,
        SUM(CASE WHEN qc_review_type IN ('production', 'production_qc', 'production qc', 'qc') THEN 1 ELSE 0 END) AS production_qc_completed_count,
        MAX(CASE WHEN qc_review_type IN ('production', 'production_qc', 'production qc', 'qc') THEN completed_at END) AS production_qc_last_completed_at,
        SUM(CASE WHEN qc_review_type IN ('peer_review', 'peer review', 'peer') THEN 1 ELSE 0 END) AS peer_review_completed_count,
        MAX(CASE WHEN qc_review_type IN ('peer_review', 'peer review', 'peer') THEN completed_at END) AS peer_review_last_completed_at,
        SUM(CASE WHEN qc_review_type IN ('independent_check', 'independent check', 'independent review', 'independent', 'ic') THEN 1 ELSE 0 END) AS independent_check_completed_count,
        MAX(CASE WHEN qc_review_type IN ('independent_check', 'independent check', 'independent review', 'independent', 'ic') THEN completed_at END) AS independent_check_last_completed_at
    FROM qc_cycle_completions
    WHERE sheet_package_id IS NOT NULL
    GROUP BY sheet_package_id
)
UPDATE sp
SET
    sp.production_qc_completed_count = ISNULL(b.production_qc_completed_count, 0),
    sp.production_qc_last_completed_at = b.production_qc_last_completed_at,
    sp.peer_review_completed_count = ISNULL(b.peer_review_completed_count, 0),
    sp.peer_review_last_completed_at = b.peer_review_last_completed_at,
    sp.independent_check_completed_count = ISNULL(b.independent_check_completed_count, 0),
    sp.independent_check_last_completed_at = b.independent_check_last_completed_at,
    sp.last_updated_at = SYSDATETIMEOFFSET()
FROM sheet_packages sp
INNER JOIN bucketed b ON sp.sheet_package_id = b.sheet_package_id;
GO

SELECT report_type, COUNT(*) AS row_count
FROM qc_cycle_completion_package_backfill_report
GROUP BY report_type
ORDER BY report_type;
GO

SELECT
    mapped_completions = (SELECT COUNT(*) FROM qc_cycle_completions WHERE sheet_package_id IS NOT NULL),
    unmapped_completions = (SELECT COUNT(*) FROM qc_cycle_completions WHERE sheet_package_id IS NULL),
    duplicate_package_keys = (SELECT COUNT(*) FROM qc_cycle_completion_package_backfill_report WHERE report_type = 'duplicate_package_logical_key');
GO
