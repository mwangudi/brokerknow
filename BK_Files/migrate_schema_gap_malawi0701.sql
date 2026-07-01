SET NOCOUNT ON;
-- (A) Tables the current prod (Malawi0701) has that the clean schema is MISSING.
PRINT '===== TABLES in Malawi0701 not in BrokerKnow_Clean =====';
SELECT s.name AS missing_table,
       (SELECT SUM(p.rows) FROM BrokerKnow_Malawi0701.sys.partitions p WHERE p.object_id=s.object_id AND p.index_id IN(0,1)) AS src_rows
FROM BrokerKnow_Malawi0701.sys.tables s
WHERE s.schema_id = SCHEMA_ID('dbo')
  AND NOT EXISTS (SELECT 1 FROM BrokerKnow_Clean.sys.tables c WHERE c.name = s.name AND c.schema_id=SCHEMA_ID('dbo'))
ORDER BY s.name;

-- (B) String columns NARROWER in clean than in Malawi0701 (width regressions) on shared tables.
PRINT '===== COLUMN width regressions (clean < Malawi0701) =====';
SELECT st.name AS tbl, sc.name AS col, sty.name AS typ,
       sc.max_length AS src_len, cc.max_length AS clean_len
FROM BrokerKnow_Malawi0701.sys.columns sc
JOIN BrokerKnow_Malawi0701.sys.tables st ON st.object_id=sc.object_id AND st.schema_id=SCHEMA_ID('dbo')
JOIN BrokerKnow_Malawi0701.sys.types sty ON sty.user_type_id=sc.user_type_id
JOIN BrokerKnow_Clean.sys.tables ct ON ct.name=st.name AND ct.schema_id=SCHEMA_ID('dbo')
JOIN BrokerKnow_Clean.sys.columns cc ON cc.object_id=ct.object_id AND cc.name=sc.name
WHERE sty.name IN ('varchar','nvarchar','char','nchar')
  AND cc.max_length < sc.max_length AND sc.max_length <> -1
ORDER BY st.name, sc.name;

-- (C) Columns on shared tables that Malawi0701 has but clean LACKS (missing cols).
PRINT '===== COLUMNS in Malawi0701 not in BrokerKnow_Clean (shared tables) =====';
SELECT st.name AS tbl, sc.name AS col
FROM BrokerKnow_Malawi0701.sys.columns sc
JOIN BrokerKnow_Malawi0701.sys.tables st ON st.object_id=sc.object_id AND st.schema_id=SCHEMA_ID('dbo')
JOIN BrokerKnow_Clean.sys.tables ct ON ct.name=st.name AND ct.schema_id=SCHEMA_ID('dbo')
WHERE NOT EXISTS (SELECT 1 FROM BrokerKnow_Clean.sys.columns cc WHERE cc.object_id=ct.object_id AND cc.name=sc.name)
ORDER BY st.name, sc.name;
