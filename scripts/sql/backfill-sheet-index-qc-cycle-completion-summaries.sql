-- Rebuilds sheet_index QC completion summary columns from qc_cycle_completions.
-- Safe to run repeatedly; does not reconstruct historical completions from workflow history.
SET QUOTED_IDENTIFIER ON;
GO

IF OBJECT_ID('dbo.qc_cycle_completions', 'U') IS NULL
BEGIN
    RAISERROR('qc_cycle_completions table does not exist. Run Initialize-QCDatabaseSchema first.', 16, 1);
    RETURN;
END
GO

IF OBJECT_ID('dbo.sheet_index', 'U') IS NULL
BEGIN
    RAISERROR('sheet_index table does not exist. Run Initialize-QCDatabaseSchema first.', 16, 1);
    RETURN;
END
GO

;WITH bucketed AS (
    SELECT
        document_guid,
        SUM(CASE
            WHEN LOWER(LTRIM(RTRIM(qc_review_type))) IN ('production qc', 'production', 'qc', 'production_qc')
            THEN 1 ELSE 0 END) AS production_qc_completed_count,
        MAX(CASE
            WHEN LOWER(LTRIM(RTRIM(qc_review_type))) IN ('production qc', 'production', 'qc', 'production_qc')
            THEN completed_at END) AS production_qc_last_completed_at,
        SUM(CASE
            WHEN LOWER(LTRIM(RTRIM(qc_review_type))) IN ('peer review', 'peer_review', 'peer')
            THEN 1 ELSE 0 END) AS peer_review_completed_count,
        MAX(CASE
            WHEN LOWER(LTRIM(RTRIM(qc_review_type))) IN ('peer review', 'peer_review', 'peer')
            THEN completed_at END) AS peer_review_last_completed_at,
        SUM(CASE
            WHEN LOWER(LTRIM(RTRIM(qc_review_type))) IN ('independent check', 'independent review', 'independent_check', 'independent', 'ic')
            THEN 1 ELSE 0 END) AS independent_check_completed_count,
        MAX(CASE
            WHEN LOWER(LTRIM(RTRIM(qc_review_type))) IN ('independent check', 'independent review', 'independent_check', 'independent', 'ic')
            THEN completed_at END) AS independent_check_last_completed_at
    FROM qc_cycle_completions
    GROUP BY document_guid
)
UPDATE si
SET
    si.production_qc_completed_count = ISNULL(b.production_qc_completed_count, 0),
    si.production_qc_last_completed_at = b.production_qc_last_completed_at,
    si.peer_review_completed_count = ISNULL(b.peer_review_completed_count, 0),
    si.peer_review_last_completed_at = b.peer_review_last_completed_at,
    si.independent_check_completed_count = ISNULL(b.independent_check_completed_count, 0),
    si.independent_check_last_completed_at = b.independent_check_last_completed_at,
    si.last_updated_at = SYSDATETIMEOFFSET()
FROM sheet_index si
INNER JOIN bucketed b
    ON LOWER(LTRIM(RTRIM(si.document_guid))) = LOWER(LTRIM(RTRIM(CAST(b.document_guid AS NVARCHAR(40)))));
GO

-- Zero summaries for sheet_index rows with no completion rows.
UPDATE si
SET
    si.production_qc_completed_count = 0,
    si.production_qc_last_completed_at = NULL,
    si.peer_review_completed_count = 0,
    si.peer_review_last_completed_at = NULL,
    si.independent_check_completed_count = 0,
    si.independent_check_last_completed_at = NULL,
    si.last_updated_at = SYSDATETIMEOFFSET()
FROM sheet_index si
WHERE NOT EXISTS (
    SELECT 1
    FROM qc_cycle_completions c
    WHERE LOWER(LTRIM(RTRIM(c.document_guid))) = LOWER(LTRIM(RTRIM(si.document_guid)))
);
GO
