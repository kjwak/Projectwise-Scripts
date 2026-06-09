-- Backfills sheet_packages, sheet_documents, and sheet_index.sheet_package_id from legacy sheet_index rows.
-- Also links nullable sheet_package_id on event/history tables via sheet_documents.
-- Idempotent: safe to run repeatedly.
SET QUOTED_IDENTIFIER ON;
SET NOCOUNT ON;
GO

IF OBJECT_ID('dbo.sheet_index', 'U') IS NULL
BEGIN
    RAISERROR('sheet_index table does not exist. Run Initialize-QCDatabaseSchema first.', 16, 1);
    RETURN;
END
GO

IF OBJECT_ID('dbo.sheet_packages', 'U') IS NULL OR OBJECT_ID('dbo.sheet_documents', 'U') IS NULL
BEGIN
    RAISERROR('sheet_packages/sheet_documents tables do not exist. Apply schema v1.15.0 first.', 16, 1);
    RETURN;
END
GO

IF OBJECT_ID('dbo.sheet_package_backfill_conflicts', 'U') IS NULL
BEGIN
    CREATE TABLE sheet_package_backfill_conflicts (
        id              BIGINT IDENTITY(1,1) PRIMARY KEY,
        detected_at     DATETIMEOFFSET(3) NOT NULL CONSTRAINT DF_sheet_pkg_bf_conflicts_detected DEFAULT SYSDATETIMEOFFSET(),
        conflict_type   NVARCHAR(50) NOT NULL,
        folder_path     NVARCHAR(1000) NULL,
        sheet_stem      NVARCHAR(260) NULL,
        document_role   NVARCHAR(50) NULL,
        document_guid   NVARCHAR(50) NULL,
        document_name   NVARCHAR(500) NULL,
        details         NVARCHAR(MAX) NULL
    );
    CREATE INDEX IX_sheet_pkg_bf_conflicts_type ON sheet_package_backfill_conflicts (conflict_type);
END
GO

TRUNCATE TABLE sheet_package_backfill_conflicts;
GO

IF OBJECT_ID('tempdb..#classified') IS NOT NULL DROP TABLE #classified;
CREATE TABLE #classified (
    source_kind                         NVARCHAR(30) NOT NULL,
    folder_path                         NVARCHAR(1000) NOT NULL,
    sheet_stem                          NVARCHAR(260) NOT NULL,
    document_role                       NVARCHAR(50) NOT NULL,
    document_guid                       UNIQUEIDENTIFIER NOT NULL,
    document_name                       NVARCHAR(260) NOT NULL,
    extension                           NVARCHAR(20) NULL,
    source_type                         NVARCHAR(10) NULL,
    pw_state_name                       NVARCHAR(100) NULL,
    designer_email                      NVARCHAR(256) NULL,
    reviewer_email                      NVARCHAR(256) NULL,
    checker_email                       NVARCHAR(256) NULL,
    qc_review_type                      NVARCHAR(100) NULL,
    qc_assigned_to                      NVARCHAR(256) NULL,
    qc_cycle_id                         NVARCHAR(100) NULL,
    qc_cycle_number                     NVARCHAR(16) NULL,
    file_modified_at                    DATETIMEOFFSET(3) NULL,
    production_qc_completed_count         INT NOT NULL DEFAULT 0,
    production_qc_last_completed_at     DATETIME2 NULL,
    peer_review_completed_count         INT NOT NULL DEFAULT 0,
    peer_review_last_completed_at       DATETIME2 NULL,
    independent_check_completed_count   INT NOT NULL DEFAULT 0,
    independent_check_last_completed_at   DATETIME2 NULL
);
GO

;WITH scoped AS (
    SELECT
        si.*,
        LTRIM(RTRIM(si.folder_path)) AS norm_folder_path,
        LTRIM(RTRIM(si.document_name)) AS norm_document_name,
        LTRIM(RTRIM(si.document_guid)) AS norm_document_guid
    FROM sheet_index si
    WHERE REPLACE(LOWER(LTRIM(RTRIM(si.folder_path))), '\', '/') LIKE '%cadd/sheets%'
),
stemmed AS (
    SELECT
        s.*,
        CASE
            WHEN LOWER(RIGHT(base_stem, 3)) = '-qc' THEN LEFT(base_stem, LEN(base_stem) - 3)
            ELSE base_stem
        END AS sheet_stem,
        CASE
            WHEN LOWER(norm_document_name) LIKE '%-qc.pdf' THEN 'qc_pdf'
            WHEN LOWER(norm_document_name) LIKE '%.dgn' THEN 'dgn'
            WHEN LOWER(norm_document_name) LIKE '%.pdf' THEN 'sheet_pdf'
            ELSE 'other'
        END AS document_role
    FROM (
        SELECT
            s.*,
            CASE
                WHEN CHARINDEX('.', norm_document_name) > 0
                    THEN LEFT(norm_document_name, LEN(norm_document_name) - CHARINDEX('.', REVERSE(norm_document_name)))
                ELSE norm_document_name
            END AS base_stem
        FROM scoped s
    ) s
)
INSERT INTO sheet_package_backfill_conflicts (conflict_type, folder_path, sheet_stem, document_role, document_guid, document_name, details)
SELECT 'missing_folder_path', NULL, NULL, NULL, norm_document_guid, norm_document_name, 'sheet_index row has blank folder_path'
FROM stemmed
WHERE NULLIF(norm_folder_path, '') IS NULL
UNION ALL
SELECT 'missing_document_name', norm_folder_path, NULL, NULL, norm_document_guid, NULL, 'sheet_index row has blank document_name'
FROM stemmed
WHERE NULLIF(norm_folder_path, '') IS NOT NULL
  AND NULLIF(norm_document_name, '') IS NULL
UNION ALL
SELECT 'invalid_guid', norm_folder_path, sheet_stem, document_role, norm_document_guid, norm_document_name, 'document_guid is not a valid UNIQUEIDENTIFIER'
FROM stemmed
WHERE NULLIF(norm_folder_path, '') IS NOT NULL
  AND NULLIF(norm_document_name, '') IS NOT NULL
  AND TRY_CAST(norm_document_guid AS UNIQUEIDENTIFIER) IS NULL;
GO

;WITH scoped AS (
    SELECT
        si.*,
        LTRIM(RTRIM(si.folder_path)) AS norm_folder_path,
        LTRIM(RTRIM(si.document_name)) AS norm_document_name,
        LTRIM(RTRIM(si.document_guid)) AS norm_document_guid
    FROM sheet_index si
    WHERE REPLACE(LOWER(LTRIM(RTRIM(si.folder_path))), '\', '/') LIKE '%cadd/sheets%'
),
stemmed AS (
    SELECT
        s.*,
        CASE
            WHEN LOWER(RIGHT(base_stem, 3)) = '-qc' THEN LEFT(base_stem, LEN(base_stem) - 3)
            ELSE base_stem
        END AS sheet_stem,
        CASE
            WHEN LOWER(norm_document_name) LIKE '%-qc.pdf' THEN 'qc_pdf'
            WHEN LOWER(norm_document_name) LIKE '%.dgn' THEN 'dgn'
            WHEN LOWER(norm_document_name) LIKE '%.pdf' THEN 'sheet_pdf'
            ELSE 'other'
        END AS document_role,
        TRY_CAST(norm_document_guid AS UNIQUEIDENTIFIER) AS parsed_guid
    FROM (
        SELECT
            s.*,
            CASE
                WHEN CHARINDEX('.', norm_document_name) > 0
                    THEN LEFT(norm_document_name, LEN(norm_document_name) - CHARINDEX('.', REVERSE(norm_document_name)))
                ELSE norm_document_name
            END AS base_stem
        FROM scoped s
    ) s
)
INSERT INTO #classified (
    source_kind, folder_path, sheet_stem, document_role, document_guid, document_name,
    extension, source_type, pw_state_name, designer_email, reviewer_email, checker_email,
    qc_review_type, qc_assigned_to, qc_cycle_id, qc_cycle_number, file_modified_at,
    production_qc_completed_count, production_qc_last_completed_at,
    peer_review_completed_count, peer_review_last_completed_at,
    independent_check_completed_count, independent_check_last_completed_at
)
SELECT
    'sheet_index_row',
    norm_folder_path,
    sheet_stem,
    document_role,
    parsed_guid,
    norm_document_name,
    extension,
    source_type,
    pw_state_name,
    designer_email,
    reviewer_email,
    checker_email,
    qc_review_type,
    qc_assigned_to,
    qc_cycle_id,
    qc_cycle_number,
    file_modified_at,
    ISNULL(production_qc_completed_count, 0),
    production_qc_last_completed_at,
    ISNULL(peer_review_completed_count, 0),
    peer_review_last_completed_at,
    ISNULL(independent_check_completed_count, 0),
    independent_check_last_completed_at
FROM stemmed
WHERE parsed_guid IS NOT NULL
  AND NULLIF(norm_folder_path, '') IS NOT NULL
  AND NULLIF(norm_document_name, '') IS NOT NULL
  AND NULLIF(sheet_stem, '') IS NOT NULL
  AND document_role IN ('dgn', 'sheet_pdf', 'qc_pdf');
GO

INSERT INTO #classified (
    source_kind, folder_path, sheet_stem, document_role, document_guid, document_name, extension
)
SELECT
    'linked_qc_pdf',
    c.folder_path,
    c.sheet_stem,
    'qc_pdf',
    TRY_CAST(LTRIM(RTRIM(si.qc_pdf_guid)) AS UNIQUEIDENTIFIER),
    LTRIM(RTRIM(si.qc_pdf_name)),
    '.pdf'
FROM sheet_index si
INNER JOIN #classified c
    ON LOWER(LTRIM(RTRIM(si.document_guid))) = LOWER(CONVERT(NVARCHAR(36), c.document_guid))
WHERE NULLIF(LTRIM(RTRIM(si.qc_pdf_guid)), '') IS NOT NULL
  AND NULLIF(LTRIM(RTRIM(si.qc_pdf_name)), '') IS NOT NULL
  AND TRY_CAST(LTRIM(RTRIM(si.qc_pdf_guid)) AS UNIQUEIDENTIFIER) IS NOT NULL
  AND NOT EXISTS (
      SELECT 1
      FROM #classified existing
      WHERE existing.document_guid = TRY_CAST(LTRIM(RTRIM(si.qc_pdf_guid)) AS UNIQUEIDENTIFIER)
  );
GO

INSERT INTO sheet_package_backfill_conflicts (conflict_type, folder_path, sheet_stem, document_role, document_guid, document_name, details)
SELECT
    'duplicate_role',
    c.folder_path,
    c.sheet_stem,
    c.document_role,
    CONVERT(NVARCHAR(36), c.document_guid),
    c.document_name,
    CONCAT(rc.role_count, ' documents compete for role ', c.document_role)
FROM #classified c
INNER JOIN (
    SELECT folder_path, sheet_stem, document_role, COUNT(*) AS role_count
    FROM #classified
    GROUP BY folder_path, sheet_stem, document_role
    HAVING COUNT(*) > 1
) rc
    ON c.folder_path = rc.folder_path
   AND c.sheet_stem = rc.sheet_stem
   AND c.document_role = rc.document_role;
GO

IF OBJECT_ID('tempdb..#package_agg') IS NOT NULL DROP TABLE #package_agg;
SELECT
    c.folder_path,
    c.sheet_stem,
    MAX(CASE WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.document_guid END) AS dgn_guid,
    MAX(CASE WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.document_name END) AS dgn_name,
    MAX(CASE WHEN c.document_role = 'sheet_pdf' AND ISNULL(rc.role_count, 0) <= 1 THEN c.document_guid END) AS sheet_pdf_guid,
    MAX(CASE WHEN c.document_role = 'sheet_pdf' AND ISNULL(rc.role_count, 0) <= 1 THEN c.document_name END) AS sheet_pdf_name,
    MAX(CASE WHEN c.document_role = 'qc_pdf' AND ISNULL(rc.role_count, 0) <= 1 THEN c.document_guid END) AS qc_pdf_guid,
    MAX(CASE WHEN c.document_role = 'qc_pdf' AND ISNULL(rc.role_count, 0) <= 1 THEN c.document_name END) AS qc_pdf_name,
    COALESCE(
        MAX(CASE WHEN c.document_role = 'qc_pdf' AND ISNULL(rc.role_count, 0) <= 1 THEN c.pw_state_name END),
        MAX(CASE WHEN c.document_role = 'sheet_pdf' AND ISNULL(rc.role_count, 0) <= 1 THEN c.pw_state_name END),
        MAX(CASE WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.pw_state_name END)
    ) AS pw_state_name,
    MAX(CASE WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.designer_email
             WHEN c.document_role = 'sheet_pdf' AND ISNULL(rc.role_count, 0) <= 1 THEN c.designer_email END) AS designer_email,
    MAX(CASE WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.reviewer_email
             WHEN c.document_role = 'sheet_pdf' AND ISNULL(rc.role_count, 0) <= 1 THEN c.reviewer_email END) AS reviewer_email,
    MAX(CASE WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.checker_email
             WHEN c.document_role = 'sheet_pdf' AND ISNULL(rc.role_count, 0) <= 1 THEN c.checker_email END) AS checker_email,
    MAX(CASE WHEN c.document_role = 'sheet_pdf' AND ISNULL(rc.role_count, 0) <= 1 THEN c.qc_review_type
             WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.qc_review_type END) AS qc_review_type,
    MAX(CASE WHEN c.document_role = 'sheet_pdf' AND ISNULL(rc.role_count, 0) <= 1 THEN c.qc_assigned_to
             WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.qc_assigned_to END) AS qc_assigned_to,
    MAX(CASE WHEN c.document_role = 'sheet_pdf' AND ISNULL(rc.role_count, 0) <= 1 THEN c.qc_cycle_id END) AS qc_cycle_id,
    MAX(CASE WHEN c.document_role = 'sheet_pdf' AND ISNULL(rc.role_count, 0) <= 1 THEN c.qc_cycle_number END) AS qc_cycle_number,
    MAX(CASE WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.production_qc_completed_count END) AS production_qc_completed_count,
    MAX(CASE WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.production_qc_last_completed_at END) AS production_qc_last_completed_at,
    MAX(CASE WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.peer_review_completed_count END) AS peer_review_completed_count,
    MAX(CASE WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.peer_review_last_completed_at END) AS peer_review_last_completed_at,
    MAX(CASE WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.independent_check_completed_count END) AS independent_check_completed_count,
    MAX(CASE WHEN c.document_role = 'dgn' AND ISNULL(rc.role_count, 0) <= 1 THEN c.independent_check_last_completed_at END) AS independent_check_last_completed_at
INTO #package_agg
FROM #classified c
LEFT JOIN (
    SELECT folder_path, sheet_stem, document_role, COUNT(*) AS role_count
    FROM #classified
    GROUP BY folder_path, sheet_stem, document_role
) rc
    ON c.folder_path = rc.folder_path
   AND c.sheet_stem = rc.sheet_stem
   AND c.document_role = rc.document_role
GROUP BY c.folder_path, c.sheet_stem;
GO

MERGE sheet_packages AS tgt
USING #package_agg AS src
    ON tgt.folder_path = src.folder_path AND tgt.sheet_stem = src.sheet_stem
WHEN MATCHED THEN UPDATE SET
    dgn_guid = src.dgn_guid,
    dgn_name = src.dgn_name,
    sheet_pdf_guid = src.sheet_pdf_guid,
    sheet_pdf_name = src.sheet_pdf_name,
    qc_pdf_guid = src.qc_pdf_guid,
    qc_pdf_name = src.qc_pdf_name,
    pw_state_name = src.pw_state_name,
    designer_email = src.designer_email,
    reviewer_email = src.reviewer_email,
    checker_email = src.checker_email,
    qc_review_type = src.qc_review_type,
    qc_assigned_to = src.qc_assigned_to,
    qc_cycle_id = src.qc_cycle_id,
    qc_cycle_number = src.qc_cycle_number,
    production_qc_completed_count = ISNULL(src.production_qc_completed_count, 0),
    production_qc_last_completed_at = src.production_qc_last_completed_at,
    peer_review_completed_count = ISNULL(src.peer_review_completed_count, 0),
    peer_review_last_completed_at = src.peer_review_last_completed_at,
    independent_check_completed_count = ISNULL(src.independent_check_completed_count, 0),
    independent_check_last_completed_at = src.independent_check_last_completed_at,
    last_updated_at = SYSDATETIMEOFFSET()
WHEN NOT MATCHED THEN INSERT (
    sheet_package_id, sheet_stem, folder_path,
    dgn_guid, dgn_name, sheet_pdf_guid, sheet_pdf_name, qc_pdf_guid, qc_pdf_name,
    pw_state_name, designer_email, reviewer_email, checker_email, qc_review_type, qc_assigned_to,
    qc_cycle_id, qc_cycle_number,
    production_qc_completed_count, production_qc_last_completed_at,
    peer_review_completed_count, peer_review_last_completed_at,
    independent_check_completed_count, independent_check_last_completed_at
) VALUES (
    NEWID(), src.sheet_stem, src.folder_path,
    src.dgn_guid, src.dgn_name, src.sheet_pdf_guid, src.sheet_pdf_name, src.qc_pdf_guid, src.qc_pdf_name,
    src.pw_state_name, src.designer_email, src.reviewer_email, src.checker_email, src.qc_review_type, src.qc_assigned_to,
    src.qc_cycle_id, src.qc_cycle_number,
    ISNULL(src.production_qc_completed_count, 0), src.production_qc_last_completed_at,
    ISNULL(src.peer_review_completed_count, 0), src.peer_review_last_completed_at,
    ISNULL(src.independent_check_completed_count, 0), src.independent_check_last_completed_at
);
GO

IF OBJECT_ID('tempdb..#documents_eligible') IS NOT NULL DROP TABLE #documents_eligible;
SELECT
    c.*,
    sp.sheet_package_id
INTO #documents_eligible
FROM #classified c
INNER JOIN sheet_packages sp
    ON sp.folder_path = c.folder_path AND sp.sheet_stem = c.sheet_stem
INNER JOIN (
    SELECT folder_path, sheet_stem, document_role
    FROM #classified
    GROUP BY folder_path, sheet_stem, document_role
    HAVING COUNT(*) = 1
) single_role
    ON c.folder_path = single_role.folder_path
   AND c.sheet_stem = single_role.sheet_stem
   AND c.document_role = single_role.document_role;
GO

MERGE sheet_documents AS tgt
USING #documents_eligible AS src
    ON tgt.document_guid = src.document_guid
WHEN MATCHED THEN UPDATE SET
    sheet_package_id = src.sheet_package_id,
    document_name = src.document_name,
    document_role = src.document_role,
    pw_state_name = src.pw_state_name,
    extension = src.extension,
    source_type = src.source_type,
    last_seen_at = SYSDATETIMEOFFSET(),
    file_modified_at = src.file_modified_at
WHEN NOT MATCHED THEN INSERT (
    document_guid, sheet_package_id, document_name, document_role,
    pw_state_name, extension, source_type, last_seen_at, file_modified_at
) VALUES (
    src.document_guid, src.sheet_package_id, src.document_name, src.document_role,
    src.pw_state_name, src.extension, src.source_type, SYSDATETIMEOFFSET(), src.file_modified_at
);
GO

UPDATE si
SET si.sheet_package_id = sd.sheet_package_id,
    si.last_updated_at = SYSDATETIMEOFFSET()
FROM sheet_index si
INNER JOIN sheet_documents sd
    ON TRY_CAST(LTRIM(RTRIM(si.document_guid)) AS UNIQUEIDENTIFIER) = sd.document_guid
WHERE si.sheet_package_id IS NULL
   OR si.sheet_package_id <> sd.sheet_package_id;
GO

UPDATE dsh
SET dsh.sheet_package_id = sd.sheet_package_id
FROM document_state_history dsh
INNER JOIN sheet_documents sd
    ON TRY_CAST(LTRIM(RTRIM(dsh.document_guid)) AS UNIQUEIDENTIFIER) = sd.document_guid
WHERE dsh.sheet_package_id IS NULL;
GO

UPDATE te
SET te.sheet_package_id = sd.sheet_package_id
FROM transition_events te
INNER JOIN sheet_documents sd
    ON TRY_CAST(LTRIM(RTRIM(te.document_guid)) AS UNIQUEIDENTIFIER) = sd.document_guid
WHERE te.sheet_package_id IS NULL;
GO

UPDATE qwe
SET qwe.sheet_package_id = sd.sheet_package_id
FROM qc_workflow_events qwe
INNER JOIN sheet_documents sd
    ON TRY_CAST(LTRIM(RTRIM(qwe.document_id)) AS UNIQUEIDENTIFIER) = sd.document_guid
WHERE qwe.sheet_package_id IS NULL;
GO

UPDATE nl
SET nl.sheet_package_id = sd.sheet_package_id
FROM notification_log nl
INNER JOIN sheet_documents sd
    ON TRY_CAST(LTRIM(RTRIM(nl.document_guid)) AS UNIQUEIDENTIFIER) = sd.document_guid
WHERE nl.sheet_package_id IS NULL;
GO

SELECT 'sheet_packages' AS entity, COUNT(*) AS row_count FROM sheet_packages
UNION ALL
SELECT 'sheet_documents', COUNT(*) FROM sheet_documents
UNION ALL
SELECT 'sheet_index_linked', COUNT(*) FROM sheet_index WHERE sheet_package_id IS NOT NULL
UNION ALL
SELECT 'backfill_conflicts', COUNT(*) FROM sheet_package_backfill_conflicts;
GO
