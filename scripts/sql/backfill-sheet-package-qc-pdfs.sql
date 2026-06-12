-- Backfill sheet_package_qc_pdfs from sheet_packages lane alias columns (v1.19).
-- Idempotent: safe to run multiple times; does not create duplicate active rows.
-- Run after Initialize-QCDatabaseSchema reaches 1.19.0.

SET NOCOUNT ON;

DECLARE @productionUpserts INT = 0;
DECLARE @checkUpserts INT = 0;
DECLARE @reviewUpserts INT = 0;

-- Deactivate stale active rows when package lane GUID changed.
UPDATE q
SET is_active = 0,
    updated_at = SYSDATETIMEOFFSET()
FROM sheet_package_qc_pdfs q
INNER JOIN sheet_packages sp ON sp.sheet_package_id = q.sheet_package_id
WHERE q.is_active = 1
  AND q.qc_process_type = 'production'
  AND sp.qc_pdf_guid IS NOT NULL
  AND q.document_guid <> sp.qc_pdf_guid;

UPDATE q
SET is_active = 0,
    updated_at = SYSDATETIMEOFFSET()
FROM sheet_package_qc_pdfs q
INNER JOIN sheet_packages sp ON sp.sheet_package_id = q.sheet_package_id
WHERE q.is_active = 1
  AND q.qc_process_type = 'check'
  AND sp.qc_chk_pdf_guid IS NOT NULL
  AND q.document_guid <> sp.qc_chk_pdf_guid;

UPDATE q
SET is_active = 0,
    updated_at = SYSDATETIMEOFFSET()
FROM sheet_package_qc_pdfs q
INNER JOIN sheet_packages sp ON sp.sheet_package_id = q.sheet_package_id
WHERE q.is_active = 1
  AND q.qc_process_type = 'review'
  AND sp.qc_rev_pdf_guid IS NOT NULL
  AND q.document_guid <> sp.qc_rev_pdf_guid;

-- Production lane (*-prod.pdf alias: qc_pdf_guid/qc_pdf_name)
MERGE sheet_package_qc_pdfs AS tgt
USING (
    SELECT
        sp.sheet_package_id,
        CAST('production' AS NVARCHAR(32)) AS qc_process_type,
        sp.qc_pdf_guid AS document_guid,
        sp.qc_pdf_name AS document_name,
        sp.folder_path,
        sp.pw_state_name AS current_pw_state
    FROM sheet_packages sp
    WHERE sp.qc_pdf_guid IS NOT NULL
      AND NULLIF(LTRIM(RTRIM(sp.qc_pdf_name)), '') IS NOT NULL
) AS src
ON tgt.sheet_package_id = src.sheet_package_id
   AND tgt.qc_process_type = src.qc_process_type
   AND tgt.document_guid = src.document_guid
   AND tgt.is_active = 1
WHEN MATCHED THEN UPDATE SET
    document_name = src.document_name,
    folder_path = src.folder_path,
    current_pw_state = COALESCE(src.current_pw_state, tgt.current_pw_state),
    last_seen_at = SYSDATETIMEOFFSET(),
    updated_at = SYSDATETIMEOFFSET()
WHEN NOT MATCHED THEN INSERT (
    sheet_package_id, qc_process_type, document_guid, document_name, folder_path,
    current_pw_state, last_seen_at, is_active
) VALUES (
    src.sheet_package_id, src.qc_process_type, src.document_guid, src.document_name, src.folder_path,
    src.current_pw_state, SYSDATETIMEOFFSET(), 1
);

SET @productionUpserts = @@ROWCOUNT;

-- Check lane
MERGE sheet_package_qc_pdfs AS tgt
USING (
    SELECT
        sp.sheet_package_id,
        CAST('check' AS NVARCHAR(32)) AS qc_process_type,
        sp.qc_chk_pdf_guid AS document_guid,
        sp.qc_chk_pdf_name AS document_name,
        sp.folder_path,
        NULL AS current_pw_state
    FROM sheet_packages sp
    WHERE sp.qc_chk_pdf_guid IS NOT NULL
      AND NULLIF(LTRIM(RTRIM(sp.qc_chk_pdf_name)), '') IS NOT NULL
) AS src
ON tgt.sheet_package_id = src.sheet_package_id
   AND tgt.qc_process_type = src.qc_process_type
   AND tgt.document_guid = src.document_guid
   AND tgt.is_active = 1
WHEN MATCHED THEN UPDATE SET
    document_name = src.document_name,
    folder_path = src.folder_path,
    last_seen_at = SYSDATETIMEOFFSET(),
    updated_at = SYSDATETIMEOFFSET()
WHEN NOT MATCHED THEN INSERT (
    sheet_package_id, qc_process_type, document_guid, document_name, folder_path,
    last_seen_at, is_active
) VALUES (
    src.sheet_package_id, src.qc_process_type, src.document_guid, src.document_name, src.folder_path,
    SYSDATETIMEOFFSET(), 1
);

SET @checkUpserts = @@ROWCOUNT;

-- Review lane
MERGE sheet_package_qc_pdfs AS tgt
USING (
    SELECT
        sp.sheet_package_id,
        CAST('review' AS NVARCHAR(32)) AS qc_process_type,
        sp.qc_rev_pdf_guid AS document_guid,
        sp.qc_rev_pdf_name AS document_name,
        sp.folder_path,
        NULL AS current_pw_state
    FROM sheet_packages sp
    WHERE sp.qc_rev_pdf_guid IS NOT NULL
      AND NULLIF(LTRIM(RTRIM(sp.qc_rev_pdf_name)), '') IS NOT NULL
) AS src
ON tgt.sheet_package_id = src.sheet_package_id
   AND tgt.qc_process_type = src.qc_process_type
   AND tgt.document_guid = src.document_guid
   AND tgt.is_active = 1
WHEN MATCHED THEN UPDATE SET
    document_name = src.document_name,
    folder_path = src.folder_path,
    last_seen_at = SYSDATETIMEOFFSET(),
    updated_at = SYSDATETIMEOFFSET()
WHEN NOT MATCHED THEN INSERT (
    sheet_package_id, qc_process_type, document_guid, document_name, folder_path,
    last_seen_at, is_active
) VALUES (
    src.sheet_package_id, src.qc_process_type, src.document_guid, src.document_name, src.folder_path,
    SYSDATETIMEOFFSET(), 1
);

SET @reviewUpserts = @@ROWCOUNT;

SELECT
    @productionUpserts AS production_merge_rows,
    @checkUpserts AS check_merge_rows,
    @reviewUpserts AS review_merge_rows,
    (SELECT COUNT(*) FROM sheet_package_qc_pdfs WHERE is_active = 1 AND qc_process_type = 'production') AS active_production_rows,
    (SELECT COUNT(*) FROM sheet_package_qc_pdfs WHERE is_active = 1 AND qc_process_type = 'check') AS active_check_rows,
    (SELECT COUNT(*) FROM sheet_package_qc_pdfs WHERE is_active = 1 AND qc_process_type = 'review') AS active_review_rows;
