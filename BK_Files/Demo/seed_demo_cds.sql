-- Pre-seed dummy CDSC numbers on six demo clients so the demo import files
-- reconcile cleanly. Run this ONCE on the droplet before the demo:
--
--   ssh root@46.101.6.131
--   /opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P 'SpringfielD##88' \
--       -C -d BrokerKnow -i /tmp/seed_demo_cds.sql
--
-- Idempotent: re-running just overwrites the same six rows.

SET NOCOUNT ON;

UPDATE Client SET ClientCDSNo = 'DEMO001' WHERE Client_DPA_ = 1;
UPDATE Client SET ClientCDSNo = 'DEMO002' WHERE Client_DPA_ = 2;
UPDATE Client SET ClientCDSNo = 'DEMO003' WHERE Client_DPA_ = 3;
UPDATE Client SET ClientCDSNo = 'DEMO004' WHERE Client_DPA_ = 4;
UPDATE Client SET ClientCDSNo = 'DEMO005' WHERE Client_DPA_ = 5;
UPDATE Client SET ClientCDSNo = 'DEMO006' WHERE Client_DPA_ = 6;

SELECT Client_DPA_ AS Id, ClientName, ClientCDSNo
FROM   Client
WHERE  Client_DPA_ BETWEEN 1 AND 6
ORDER BY Client_DPA_;
