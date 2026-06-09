-- Validation report for sheet package backfill integrity.
-- Read-only: reports issues; does not modify data.
SET NOCOUNT ON;
GO

PRINT '=== 1. sheet_index rows without sheet_package_id (CADD/Sheets scope) ===';
SELECT
    si.document_guid,
    si.document_name,
    si.folder_path,
    si.extension,
    si.source_type
FROM sheet_index si
WHERE REPLACE(LOWER(LTRIM(RTRIM(si.folder_path))), '\', '/') LIKE '%cadd/sheets%'
  AND si.sheet_package_id IS NULL
  AND (
        LOWER(si.document_name) LIKE '%.dgn'
     OR LOWER(si.document_name) LIKE '%.pdf'
  );
GO

PRINT '=== 2. sheet_documents rows without matching sheet_packages ===';
SELECT
    sd.document_guid,
    sd.document_name,
    sd.document_role,
    sd.sheet_package_id
FROM sheet_documents sd
LEFT JOIN sheet_packages sp ON sp.sheet_package_id = sd.sheet_package_id
WHERE sp.sheet_package_id IS NULL;
GO

PRINT '=== 3. sheet_packages with zero documents ===';
SELECT
    sp.sheet_package_id,
    sp.folder_path,
    sp.sheet_stem,
    sp.dgn_guid,
    sp.sheet_pdf_guid,
    sp.qc_pdf_guid
FROM sheet_packages sp
WHERE NOT EXISTS (
    SELECT 1 FROM sheet_documents sd WHERE sd.sheet_package_id = sp.sheet_package_id
);
GO

PRINT '=== 4. packages with duplicate roles in sheet_documents ===';
SELECT
    sd.sheet_package_id,
    sp.folder_path,
    sp.sheet_stem,
    sd.document_role,
    COUNT(*) AS role_count
FROM sheet_documents sd
INNER JOIN sheet_packages sp ON sp.sheet_package_id = sd.sheet_package_id
GROUP BY sd.sheet_package_id, sp.folder_path, sp.sheet_stem, sd.document_role
HAVING COUNT(*) > 1;
GO

PRINT '=== 5. packages missing DGN ===';
SELECT
    sp.sheet_package_id,
    sp.folder_path,
    sp.sheet_stem,
    sp.dgn_guid,
    sp.sheet_pdf_guid,
    sp.qc_pdf_guid
FROM sheet_packages sp
WHERE sp.dgn_guid IS NULL;
GO

PRINT '=== 6. packages missing sheet PDF ===';
SELECT
    sp.sheet_package_id,
    sp.folder_path,
    sp.sheet_stem,
    sp.dgn_guid,
    sp.sheet_pdf_guid,
    sp.qc_pdf_guid
FROM sheet_packages sp
WHERE sp.sheet_pdf_guid IS NULL;
GO

PRINT '=== 7. packages missing QC PDF ===';
SELECT
    sp.sheet_package_id,
    sp.folder_path,
    sp.sheet_stem,
    sp.dgn_guid,
    sp.sheet_pdf_guid,
    sp.qc_pdf_guid
FROM sheet_packages sp
WHERE sp.qc_pdf_guid IS NULL;
GO

PRINT '=== 8. denormalized sheet_packages GUID columns mismatched vs sheet_documents ===';
SELECT
    sp.sheet_package_id,
    sp.folder_path,
    sp.sheet_stem,
    mismatch_type = CASE
        WHEN sp.dgn_guid IS NOT NULL AND (dgn_doc.document_guid IS NULL OR sp.dgn_guid <> dgn_doc.document_guid) THEN 'dgn'
        WHEN sp.sheet_pdf_guid IS NOT NULL AND (pdf_doc.document_guid IS NULL OR sp.sheet_pdf_guid <> pdf_doc.document_guid) THEN 'sheet_pdf'
        WHEN sp.qc_pdf_guid IS NOT NULL AND (qc_doc.document_guid IS NULL OR sp.qc_pdf_guid <> qc_doc.document_guid) THEN 'qc_pdf'
    END,
    sp.dgn_guid AS package_dgn_guid,
    dgn_doc.document_guid AS document_dgn_guid,
    sp.sheet_pdf_guid AS package_sheet_pdf_guid,
    pdf_doc.document_guid AS document_sheet_pdf_guid,
    sp.qc_pdf_guid AS package_qc_pdf_guid,
    qc_doc.document_guid AS document_qc_pdf_guid
FROM sheet_packages sp
LEFT JOIN sheet_documents dgn_doc
    ON dgn_doc.sheet_package_id = sp.sheet_package_id AND dgn_doc.document_role = 'dgn'
LEFT JOIN sheet_documents pdf_doc
    ON pdf_doc.sheet_package_id = sp.sheet_package_id AND pdf_doc.document_role = 'sheet_pdf'
LEFT JOIN sheet_documents qc_doc
    ON qc_doc.sheet_package_id = sp.sheet_package_id AND qc_doc.document_role = 'qc_pdf'
WHERE
    (sp.dgn_guid IS NOT NULL AND (dgn_doc.document_guid IS NULL OR sp.dgn_guid <> dgn_doc.document_guid))
 OR (sp.sheet_pdf_guid IS NOT NULL AND (pdf_doc.document_guid IS NULL OR sp.sheet_pdf_guid <> pdf_doc.document_guid))
 OR (sp.qc_pdf_guid IS NOT NULL AND (qc_doc.document_guid IS NULL OR sp.qc_pdf_guid <> qc_doc.document_guid));
GO

PRINT '=== 9a. document_state_history linkable but sheet_package_id IS NULL ===';
SELECT
    dsh.id,
    dsh.document_guid,
    dsh.document_name,
    dsh.folder_path,
    sd.sheet_package_id AS expected_sheet_package_id
FROM document_state_history dsh
INNER JOIN sheet_documents sd
    ON TRY_CAST(LTRIM(RTRIM(dsh.document_guid)) AS UNIQUEIDENTIFIER) = sd.document_guid
WHERE dsh.sheet_package_id IS NULL;
GO

PRINT '=== 9b. transition_events linkable but sheet_package_id IS NULL ===';
SELECT
    te.id,
    te.document_guid,
    te.document_name,
    te.folder_path,
    sd.sheet_package_id AS expected_sheet_package_id
FROM transition_events te
INNER JOIN sheet_documents sd
    ON TRY_CAST(LTRIM(RTRIM(te.document_guid)) AS UNIQUEIDENTIFIER) = sd.document_guid
WHERE te.sheet_package_id IS NULL;
GO

PRINT '=== 9c. qc_workflow_events linkable but sheet_package_id IS NULL ===';
SELECT
    qwe.event_id,
    qwe.document_id,
    sd.sheet_package_id AS expected_sheet_package_id
FROM qc_workflow_events qwe
INNER JOIN sheet_documents sd
    ON TRY_CAST(LTRIM(RTRIM(qwe.document_id)) AS UNIQUEIDENTIFIER) = sd.document_guid
WHERE qwe.sheet_package_id IS NULL;
GO

PRINT '=== 9d. notification_log linkable but sheet_package_id IS NULL ===';
SELECT
    nl.id,
    nl.document_guid,
    nl.document_name,
    nl.folder_path,
    sd.sheet_package_id AS expected_sheet_package_id
FROM notification_log nl
INNER JOIN sheet_documents sd
    ON TRY_CAST(LTRIM(RTRIM(nl.document_guid)) AS UNIQUEIDENTIFIER) = sd.document_guid
WHERE nl.sheet_package_id IS NULL;
GO

PRINT '=== Backfill conflict report (latest run) ===';
SELECT
    conflict_type,
    folder_path,
    sheet_stem,
    document_role,
    document_guid,
    document_name,
    details,
    detected_at
FROM sheet_package_backfill_conflicts
ORDER BY conflict_type, folder_path, sheet_stem, document_role;
GO

PRINT '=== Summary counts ===';
SELECT metric = 'sheet_packages', value = COUNT(*) FROM sheet_packages
UNION ALL SELECT 'sheet_documents', COUNT(*) FROM sheet_documents
UNION ALL SELECT 'sheet_index_with_package_id', COUNT(*) FROM sheet_index WHERE sheet_package_id IS NOT NULL
UNION ALL SELECT 'sheet_index_missing_package_id', COUNT(*)
    FROM sheet_index
    WHERE sheet_package_id IS NULL
      AND REPLACE(LOWER(LTRIM(RTRIM(folder_path))), '\', '/') LIKE '%cadd/sheets%'
      AND (LOWER(document_name) LIKE '%.dgn' OR LOWER(document_name) LIKE '%.pdf')
UNION ALL SELECT 'backfill_conflicts', COUNT(*) FROM sheet_package_backfill_conflicts;
GO
