#!/usr/bin/env bash
set +e
echo '===NGINX==='
cat /etc/nginx/sites-enabled/brokerknow
echo
echo '===SYSTEMD==='
cat /etc/systemd/system/brokerknow-api.service
echo
echo '===APPSETTINGS-KEYS==='
grep -E '"(Urls|ConnectionStrings|Database|Uploads|Jwt|AllowedHosts|Cors)"' /opt/brokerknow/api/appsettings.json | sed 's/Password=[^;"]*/Password=***/g'
echo
echo '===DBS==='
SAPW=$(cat /root/.mssql-sa 2>/dev/null)
/opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "$SAPW" -C -h -1 -W -Q "SET NOCOUNT ON; SELECT name + '|' + recovery_model_desc FROM sys.databases;" 2>&1 | head -20
echo
echo '===CURRENT-CRON==='
crontab -l 2>/dev/null
ls /etc/cron.d/ 2>/dev/null
echo
echo '===SPA-DIRS==='
ls -la /var/www/ 2>/dev/null
