SET NOCOUNT ON;
USE BrokerKnow_Clean;
GO
PRINT '===== FKs that reference the demo parents (child.col -> parent.col) =====';
SELECT fk.name,
       OBJECT_NAME(fk.parent_object_id) + '.' + cpar.name AS child_col,
       OBJECT_NAME(fk.referenced_object_id) + '.' + cref.name AS parent_col
FROM sys.foreign_keys fk
JOIN sys.foreign_key_columns fkc ON fkc.constraint_object_id = fk.object_id
JOIN sys.columns cpar ON cpar.object_id = fkc.parent_object_id AND cpar.column_id = fkc.parent_column_id
JOIN sys.columns cref ON cref.object_id = fkc.referenced_object_id AND cref.column_id = fkc.referenced_column_id
WHERE OBJECT_NAME(fk.parent_object_id) IN ('Client','Payment','tbOrder','OrdDetail')
ORDER BY child_col;

PRINT '===== NOT-NULL columns of parent lookup + Client/Payment (must populate) =====';
SELECT t.name AS tbl, c.column_id, c.name AS col,
       ty.name + CASE WHEN ty.name IN ('varchar','nvarchar','char','nchar')
                      THEN '(' + CAST(c.max_length AS varchar) + ')' ELSE '' END AS typ,
       CASE WHEN c.is_identity = 1 THEN 'IDENTITY' ELSE '' END AS ident
FROM sys.columns c
JOIN sys.tables t ON t.object_id = c.object_id
JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE t.name IN ('Branch','Class','Commission','Residency','EntityType','Gender','Client','Payment','Status')
  AND c.is_nullable = 0
ORDER BY t.name, c.column_id;
GO
