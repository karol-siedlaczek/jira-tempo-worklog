# jira-tempo-worklog

### Description
### `tempo-worklog.py`<br>
Creates a csv report with billed hours from Tempo<br>
Available arguments:
- `-b/--beginDate`	starting range of date in yyyy-mm-dd format, the default is the first day of the previous month
- `-e/--endDate`	ending date range in yyyy-mm-dd format, default is the first day of the current month
- `-k/--projectKey`	billed hours can be took from a specific project, just enter its key, in other case script will download billed hours from all projects
- `-f/--ftpHost`	host to which the report is to be sent
- `-d/--ftpDir`	path to which the report is to be sent (if it does not exist and the user has the appropriate permissions, it will be created)
- `-u/--ftpUser`	user that the script should use to send the report to FTP
- `-p/--ftpPassword`	Password needed to log in to FTP (if needed)
- `-m/--recipientMail`	Report recipient, you can enter several recipients in the format -m mail1 mail2 mail
- `-l/--labels`   Include labels from issue

### API documentation
https://www.tempo.io/server-api-documentation/servlet
