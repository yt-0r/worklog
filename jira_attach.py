# Стандартные библиотеки
import os
# Сторонние библиотеки
from jira import JIRA

def JiraAttach(config, issuekey, file_path, Login, Pasword, flag):
    jira_options = {'server': config["jira"]["jira_server"]}
    jira = JIRA(options=jira_options, basic_auth=(Login, Pasword))
    jql = "issuekey = \""+issuekey+"\" AND attachments is not EMPTY"
    query = jira.search_issues(jql_str=jql, json_result=True, fields="key, attachment")
    for i in query['issues']:
        for a in i['fields']['attachment']:
            if a['filename'].find('.xlsx') > 0:
                try:
                    jira.delete_attachment(a['id'])
                except Exception:
                    pass
    issue = jira.issue(issuekey)
    jira.add_attachment(issue=issue, attachment=file_path)
    if flag == False:
        os.remove(file_path)
