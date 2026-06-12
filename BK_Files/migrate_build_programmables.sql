/* migrate_build_programmables.sql — run against BrokerKnow_Clean.
   Pulls every view/proc/function definition from the SOURCE prod DB (BrokerKnow)
   and creates it in BrokerKnow_Clean. Multi-pass to resolve view-on-view /
   function dependency ordering. READ-ONLY on prod (only reads sys + creates in clean).
   Reports created/failed by type and lists failures (expected: objects that
   reference quarantined [trash] tables, which don't exist in the clean target). */
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
USE BrokerKnow_Clean;
GO

IF OBJECT_ID('tempdb..#prog') IS NOT NULL DROP TABLE #prog;
CREATE TABLE #prog (
    id          int IDENTITY(1,1) PRIMARY KEY,
    obj_name    sysname,
    obj_type    char(2),
    definition  nvarchar(max),
    qi          bit          NOT NULL DEFAULT 1,  -- module's stored QUOTED_IDENTIFIER
    an          bit          NOT NULL DEFAULT 1,  -- module's stored ANSI_NULLS
    done        bit          NOT NULL DEFAULT 0,
    attempt_pass int         NULL,
    last_err    nvarchar(2000) NULL
);

-- definitions come from prod (the migration source); create order V/FN/IF/TF before P.
-- Capture each module's stored SET options: legacy views created with QUOTED_IDENTIFIER OFF
-- use "" as string literals, so they MUST be re-created with QI OFF or they fail to bind.
-- NOTE: join BrokerKnow.sys.schemas (NOT SCHEMA_NAME(o.schema_id)) because SCHEMA_NAME
-- resolves ids in the CURRENT db, and the app schema_id differs between prod and clean
-- (prod app=7, clean app=5) -> SCHEMA_NAME would silently drop every app.* view.
INSERT INTO #prog (obj_name, obj_type, definition, qi, an)
SELECT o.name, o.type, m.definition, m.uses_quoted_identifier, m.uses_ansi_nulls
FROM BrokerKnow.sys.sql_modules m
JOIN BrokerKnow.sys.objects o ON o.object_id = m.object_id
JOIN BrokerKnow.sys.schemas s ON s.schema_id = o.schema_id
WHERE o.type IN ('V','P','FN','IF','TF')
  AND s.name IN ('dbo','app')
  AND m.definition IS NOT NULL
  -- skip any that already exist in clean (idempotent re-run); use prod's schema name
  AND OBJECT_ID(QUOTENAME(s.name) + '.' + QUOTENAME(o.name)) IS NULL;

GO
/* ---- create each object under its OWN stored QUOTED_IDENTIFIER / ANSI_NULLS --------
   One GO-batch per (QI, ANSI_NULLS) combo. The SET MUST be a top-level batch statement:
   QUOTED_IDENTIFIER is a parse-time setting, so toggling it inside IF/WHILE does NOT
   propagate into EXEC() (that silently dropped the QI-OFF legacy views). EXEC() (not
   sp_executesql, which forces QI ON) compiles @def under the batch's session settings,
   so the object is stored with the same options it had in prod. Inner multi-pass
   resolves view-on-view ordering within the combo; #prog persists across these GO
   batches (a temp table lives for the whole connection). Combos run qi-on first so the
   bulk (which has no cross-combo deps) is in place before the qi-off leaf views. */

----- QUOTED_IDENTIFIER ON / ANSI_NULLS ON -----
SET QUOTED_IDENTIFIER ON;  SET ANSI_NULLS ON;
DECLARE @qi bit = 1, @an bit = 1, @p int = 0, @prog int = 1, @id int, @def nvarchar(max);
WHILE @prog > 0 AND @p < 15
BEGIN
    SET @prog = 0; SET @p += 1;
    WHILE EXISTS (SELECT 1 FROM #prog WHERE done = 0 AND qi = @qi AND an = @an AND ISNULL(attempt_pass, 0) < @p)
    BEGIN
        SELECT TOP 1 @id = id, @def = definition FROM #prog
          WHERE done = 0 AND qi = @qi AND an = @an AND ISNULL(attempt_pass, 0) < @p ORDER BY id;
        UPDATE #prog SET attempt_pass = @p WHERE id = @id;
        BEGIN TRY EXEC (@def); UPDATE #prog SET done = 1, last_err = NULL WHERE id = @id; SET @prog += 1; END TRY
        BEGIN CATCH UPDATE #prog SET last_err = ERROR_MESSAGE() WHERE id = @id; END CATCH
    END
END
GO

----- QUOTED_IDENTIFIER ON / ANSI_NULLS OFF -----
SET QUOTED_IDENTIFIER ON;  SET ANSI_NULLS OFF;
DECLARE @qi bit = 1, @an bit = 0, @p int = 0, @prog int = 1, @id int, @def nvarchar(max);
WHILE @prog > 0 AND @p < 15
BEGIN
    SET @prog = 0; SET @p += 1;
    WHILE EXISTS (SELECT 1 FROM #prog WHERE done = 0 AND qi = @qi AND an = @an AND ISNULL(attempt_pass, 0) < @p)
    BEGIN
        SELECT TOP 1 @id = id, @def = definition FROM #prog
          WHERE done = 0 AND qi = @qi AND an = @an AND ISNULL(attempt_pass, 0) < @p ORDER BY id;
        UPDATE #prog SET attempt_pass = @p WHERE id = @id;
        BEGIN TRY EXEC (@def); UPDATE #prog SET done = 1, last_err = NULL WHERE id = @id; SET @prog += 1; END TRY
        BEGIN CATCH UPDATE #prog SET last_err = ERROR_MESSAGE() WHERE id = @id; END CATCH
    END
END
GO

----- QUOTED_IDENTIFIER OFF / ANSI_NULLS ON -----
SET QUOTED_IDENTIFIER OFF; SET ANSI_NULLS ON;
DECLARE @qi bit = 0, @an bit = 1, @p int = 0, @prog int = 1, @id int, @def nvarchar(max);
WHILE @prog > 0 AND @p < 15
BEGIN
    SET @prog = 0; SET @p += 1;
    WHILE EXISTS (SELECT 1 FROM #prog WHERE done = 0 AND qi = @qi AND an = @an AND ISNULL(attempt_pass, 0) < @p)
    BEGIN
        SELECT TOP 1 @id = id, @def = definition FROM #prog
          WHERE done = 0 AND qi = @qi AND an = @an AND ISNULL(attempt_pass, 0) < @p ORDER BY id;
        UPDATE #prog SET attempt_pass = @p WHERE id = @id;
        BEGIN TRY EXEC (@def); UPDATE #prog SET done = 1, last_err = NULL WHERE id = @id; SET @prog += 1; END TRY
        BEGIN CATCH UPDATE #prog SET last_err = ERROR_MESSAGE() WHERE id = @id; END CATCH
    END
END
GO

----- QUOTED_IDENTIFIER OFF / ANSI_NULLS OFF (legacy list views) -----
SET QUOTED_IDENTIFIER OFF; SET ANSI_NULLS OFF;
DECLARE @qi bit = 0, @an bit = 0, @p int = 0, @prog int = 1, @id int, @def nvarchar(max);
WHILE @prog > 0 AND @p < 15
BEGIN
    SET @prog = 0; SET @p += 1;
    WHILE EXISTS (SELECT 1 FROM #prog WHERE done = 0 AND qi = @qi AND an = @an AND ISNULL(attempt_pass, 0) < @p)
    BEGIN
        SELECT TOP 1 @id = id, @def = definition FROM #prog
          WHERE done = 0 AND qi = @qi AND an = @an AND ISNULL(attempt_pass, 0) < @p ORDER BY id;
        UPDATE #prog SET attempt_pass = @p WHERE id = @id;
        BEGIN TRY EXEC (@def); UPDATE #prog SET done = 1, last_err = NULL WHERE id = @id; SET @prog += 1; END TRY
        BEGIN CATCH UPDATE #prog SET last_err = ERROR_MESSAGE() WHERE id = @id; END CATCH
    END
END
GO

SET QUOTED_IDENTIFIER ON; SET ANSI_NULLS ON;
PRINT '===== programmables created vs failed (by type) =====';
SELECT obj_type,
       COUNT(*)                                   AS total,
       SUM(CAST(done AS int))                      AS created,
       SUM(CASE WHEN done = 0 THEN 1 ELSE 0 END)   AS failed
FROM #prog
GROUP BY obj_type
ORDER BY obj_type;

PRINT '===== totals =====';
SELECT COUNT(*) AS total, SUM(CAST(done AS int)) AS created,
       SUM(CASE WHEN done = 0 THEN 1 ELSE 0 END) AS failed
FROM #prog;

PRINT '===== FAILURES (object -> reason) =====';
SELECT obj_type, obj_name, last_err
FROM #prog
WHERE done = 0
ORDER BY obj_type, obj_name;

DROP TABLE #prog;
GO
