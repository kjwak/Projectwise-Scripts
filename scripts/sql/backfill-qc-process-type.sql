-- Backfill qc_process_type from qc_review_type (non-destructive).
-- Run after schema v1.18 migration.

UPDATE sheet_index
SET qc_process_type = CASE LOWER(LTRIM(RTRIM(ISNULL(qc_review_type, ''))))
    WHEN 'production qc' THEN 'production'
    WHEN 'independent check' THEN 'check'
    WHEN 'peer review' THEN 'review'
    WHEN 'peer_review' THEN 'review'
    WHEN 'independent_check' THEN 'check'
    WHEN 'production' THEN 'production'
    WHEN 'check' THEN 'check'
    WHEN 'review' THEN 'review'
    ELSE NULL
END
WHERE qc_process_type IS NULL AND qc_review_type IS NOT NULL;

UPDATE sheet_packages
SET qc_process_type = CASE LOWER(LTRIM(RTRIM(ISNULL(qc_review_type, ''))))
    WHEN 'production qc' THEN 'production'
    WHEN 'independent check' THEN 'check'
    WHEN 'peer review' THEN 'review'
    WHEN 'peer_review' THEN 'review'
    WHEN 'independent_check' THEN 'check'
    WHEN 'production' THEN 'production'
    WHEN 'check' THEN 'check'
    WHEN 'review' THEN 'review'
    ELSE NULL
END
WHERE qc_process_type IS NULL AND qc_review_type IS NOT NULL;

UPDATE qc_cycle_completions
SET qc_process_type = CASE LOWER(LTRIM(RTRIM(ISNULL(qc_review_type, ''))))
    WHEN 'production qc' THEN 'production'
    WHEN 'production' THEN 'production'
    WHEN 'independent check' THEN 'check'
    WHEN 'independent_check' THEN 'check'
    WHEN 'check' THEN 'check'
    WHEN 'peer review' THEN 'review'
    WHEN 'peer_review' THEN 'review'
    WHEN 'review' THEN 'review'
    ELSE NULL
END
WHERE qc_process_type IS NULL AND qc_review_type IS NOT NULL;
