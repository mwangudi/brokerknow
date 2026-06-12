SET NOCOUNT ON;
USE BrokerKnow;  -- prod, source of truth for lookup contents
GO
PRINT '===== OrderHoldOption columns (prod) =====';
SELECT c.column_id, c.name, ty.name AS typ,
       CASE WHEN ty.name IN ('varchar','nvarchar','char') THEN '('+CAST(c.max_length AS varchar)+')' ELSE '' END AS len,
       CASE WHEN c.is_nullable=0 THEN 'NOT NULL' ELSE '' END AS nn,
       CASE WHEN c.is_identity=1 THEN 'IDENTITY' ELSE '' END AS ident
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id=c.user_type_id
WHERE c.object_id=OBJECT_ID('dbo.OrderHoldOption') ORDER BY c.column_id;

PRINT '===== OrderHoldOption rows (prod) =====';
SELECT * FROM dbo.OrderHoldOption ORDER BY 1;

PRINT '===== OrderHoldType columns (prod) =====';
SELECT c.column_id, c.name, ty.name AS typ,
       CASE WHEN ty.name IN ('varchar','nvarchar','char') THEN '('+CAST(c.max_length AS varchar)+')' ELSE '' END AS len,
       CASE WHEN c.is_nullable=0 THEN 'NOT NULL' ELSE '' END AS nn,
       CASE WHEN c.is_identity=1 THEN 'IDENTITY' ELSE '' END AS ident
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id=c.user_type_id
WHERE c.object_id=OBJECT_ID('dbo.OrderHoldType') ORDER BY c.column_id;

PRINT '===== OrderHoldType rows (prod) =====';
SELECT * FROM dbo.OrderHoldType ORDER BY 1;
GO
