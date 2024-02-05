https://developers.google.com/apps-script/api/reference/rest

https://developers.google.com/apps-script/add-ons/how-tos/building-workspace-addons

## create a  cron job  on Ubuntu( every 10 minutes) output the log in output.log file
0 3 * * * TZ='America/New_York' /usr/bin/python3 /root/cron/_main_.py >> /root/cron/output.log 2>&1

## check cron job logs
grep CRON /var/log/syslog