-- Validation report for qc_cycle_completions package identity migration.
SET NOCOUNT ON;
GO

PRINT '=== 1. completions missing sheet_package_id ===';
SELECT
    c.id,
    c.document_guid,
    c.document_name,
    c.qc_cycle_id,
    c.qc_review_type,
    c.completed_at
FROM qc_cycle_completions c
WHERE c.sheet_package_id IS NULL;
GO

PRINT '=== 2. duplicate package + cycle + review type records ===';
SELECT
    c.sheet_package_id,
    c.qc_cycle_id,
    c.qc_review_type,
    COUNT(*) AS duplicate_count,
    MIN(c.id) AS min_id,
    MAX(c.id) AS max_id
FROM qc_cycle_completions c
WHERE c.sheet_package_id IS NOT NULL
GROUP BY c.sheet_package_id, c.qc_cycle_id, c.qc_review_type
HAVING COUNT(*) > 1;
GO

PRINT '=== 3. completions where document_guid maps to a different package than sheet_package_id ===';
SELECT
    c.id,
    c.document_guid,
    c.sheet_package_id AS completion_sheet_package_id,
    sd.sheet_package_id AS document_sheet_package_id,
    c.qc_cycle_id,
    c.qc_review_type
FROM qc_cycle_completions c
INNER JOIN sheet_documents sd ON c.document_guid = sd.document_guid
WHERE c.sheet_package_id IS NOT NULL
  AND c.sheet_package_id <> sd.sheet_package_id;
GO

PRINT '=== 4. sheet_packages rollup mismatches vs qc_cycle_completions aggregates ===';
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
SELECT
    sp.sheet_package_id,
    sp.folder_path,
    sp.sheet_stem,
    mismatch = CASE
        WHEN ISNULL(sp.production_qc_completed_count, 0) <> ISNULL(b.production_qc_completed_count, 0) THEN 'production_count'
        WHEN sp.production_qc_last_completed_at <> b.production_qc_last_completed_at
             OR (sp.production_qc_last_completed_at IS NULL AND b.production_qc_last_completed_at IS NOT NULL)
             OR (sp.production_qc_last_completed_at IS NOT NULL AND b.production_qc_last_completed_at IS NULL) THEN 'production_last'
        WHEN ISNULL(sp.peer_review_completed_count, 0) <> ISNULL(b.peer_review_completed_count, 0) THEN 'peer_count'
        WHEN sp.peer_review_last_completed_at <> b.peer_review_last_completed_at
             OR (sp.peer_review_last_completed_at IS NULL AND b.peer_review_last_completed_at IS NOT NULL)
             OR (sp.peer_review_last_completed_at IS NOT NULL AND b.peer_review_last_completed_at IS NULL) THEN 'peer_last'
        WHEN ISNULL(sp.independent_check_completed_count, 0) <> ISNULL(b.independent_check_completed_count, 0) THEN 'independent_count'
        WHEN sp.independent_check_last_completed_at <> b.independent_check_last_completed_at
             OR (sp.independent_check_last_completed_at IS NULL AND b.independent_check_last_completed_at IS NOT NULL)
             OR (sp.independent_check_last_completed_at IS NOT NULL AND b.independent_check_last_completed_at IS NULL) THEN 'independent_last'
        ELSE 'unknown'
    END,
    sp.production_qc_completed_count AS package_production_count,
    b.production_qc_completed_count AS agg_production_count,
    sp.production_qc_last_completed_at AS package_production_last,
    b.production_qc_last_completed_at AS agg_production_last,
    sp.peer_review_completed_count AS package_peer_count,
    b.peer_review_completed_count AS agg_peer_count,
    sp.independent_check_completed_count AS package_independent_count,
    b.independent_check_completed_count AS agg_independent_count
FROM sheet_packages sp
INNER JOIN bucketed b ON sp.sheet_package_id = b.sheet_package_id
WHERE ISNULL(sp.production_qc_completed_count, 0) <> ISNULL(b.production_qc_completed_count, 0)
   OR sp.production_qc_last_completed_at <> b.production_qc_last_completed_at
   OR (sp.production_qc_last_completed_at IS NULL AND b.production_qc_last_completed_at IS NOT NULL)
   OR (sp.production_qc_last_completed_at IS NOT NULL AND b.production_qc_last_completed_at IS NULL)
   OR ISNULL(sp.peer_review_completed_count, 0) <> ISNULL(b.peer_review_completed_count, 0)
   OR sp.peer_review_last_completed_at <> b.peer_review_last_completed_at
   OR (sp.peer_review_last_completed_at IS NULL AND b.peer_review_last_completed_at IS NOT NULL)
   OR (sp.peer_review_last_completed_at IS NOT NULL AND b.peer_review_last_completed_at IS NULL)
   OR ISNULL(sp.independent_check_completed_count, 0) <> ISNULL(b.independent_check_completed_count, 0)
   OR sp.independent_check_last_completed_at <> b.independent_check_last_completed_at
   OR (sp.independent_check_last_completed_at IS NULL AND b.independent_check_last_completed_at IS NOT NULL)
   OR (sp.independent_check_last_completed_at IS NOT NULL AND b.independent_check_last_completed_at IS NULL);
GO

PRINT '=== Summary counts ===';
SELECT metric = 'completions_total', value = COUNT(*) FROM qc_cycle_completions
UNION ALL SELECT 'completions_with_package_id', COUNT(*) FROM qc_cycle_completions WHERE sheet_package_id IS NOT NULL
UNION ALL SELECT 'completions_missing_package_id', COUNT(*) FROM qc_cycle_completions WHERE sheet_package_id IS NULL
UNION ALL SELECT 'duplicate_package_logical_keys', COUNT(*)
    FROM (
        SELECT sheet_package_id, qc_cycle_id, qc_review_type
        FROM qc_cycle_completions
        WHERE sheet_package_id IS NOT NULL
        GROUP BY sheet_package_id, qc_cycle_id, qc_review_type
        HAVING COUNT(*) > 1
    ) d;
GO
