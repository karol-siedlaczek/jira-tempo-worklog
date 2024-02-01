#!/usr/bin/env python3

# Author: K. Siedlaczek 2021

import requests
from requests.exceptions import RequestException
import ftplib
import json
import csv
import os
import re
import argparse
import smtplib
import xml.etree.ElementTree as ET
from datetime import datetime
from dateutil.relativedelta import relativedelta
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


JIRA_URL = '<JIRA_URL>'
FORMAT = 'xml'  # excel/xml
EXTERNAL_HOURS = 'false'  # default = true
ISSUE_DETAILS = 'true'
USER_DETAILS = 'true'
ISSUE_SUMMARY = 'true'
BILLING_INFO = 'false'
TEMPO_API_TOKEN = '<TEMPO_API_TOKEN>'
JIRA_TOKEN = '<JIRA_TOKEN>'
SENDER_EMAIL = '<SENDER_EMAIL>'
SENDER_PASSWORD = '<SENDER_PASSWORD>'
SMTP_SERVER = '<SMTP_SERVER>'
SMTP_PORT = 465


class TempoData:
    __slots__ = ['issue_key', 'issue_summary', 'hours', 'work_date',
                 'full_name', 'period', 'month', 'issue_type', 'issue_status',
                 'project_key', 'project_name', 'issue_labels']
    
  def __init__(self, issue_key, issue_summary, hours, work_date, full_name, period, month, issue_type, issue_status, project_key, project_name, issue_labels):
        self.issue_key = issue_key
        self.issue_summary = issue_summary
        self.hours = hours
        self.work_date = work_date
        self.full_name = full_name
        self.period = period
        self.month = month
        self.issue_type = issue_type
        self.issue_status = issue_status
        self.project_key = project_key
        self.project_name = project_name
        self.issue_labels = issue_labels


def create_tempo_worklog(date_from, date_to, project_key, ftp_host, ftp_dir, ftp_user, ftp_pass, recipients, include_labels=False):
    file = f'tempo-worklog_{project_key}_{date_from}_{date_to}.xml'
    url = f'{JIRA_URL}/plugins/servlet/tempo-getWorklog/' \
          f'?dateFrom={date_from}' \
          f'&dateTo={date_to}' \
          f'&format={FORMAT}' \
          f'&useExternalHours={EXTERNAL_HOURS}' \
          f'&addIssueDetails={ISSUE_DETAILS}' \
          f'&addUserDetails={USER_DETAILS}' \
          f'&addIssueSummary={ISSUE_SUMMARY}' \
          f'&addBillingInfo={BILLING_INFO}' \
          f'&tempoApiToken={TEMPO_API_TOKEN}' \
          f'&projectKey={project_key}'
    # print(url)
    print('requesting for tempo worklogs...')
    response = requests.get(url)
    
    if not response.ok:
        raise RequestException(f'request to get tempo worklogs returned unexpected http code: {response.status_code}')
    if response.ok:
        print(f'request returned {response.status_code}')
        with open(file, 'wb') as f:
            f.write(response.content)
        tempo_data_list = get_tempo_data(file, include_labels)
        file = file.replace('.xml', '.csv')
        save_to_csv(file, tempo_data_list)
        if recipients:
            for recipient in recipients:
                send_email(file, recipient, date_from, date_to)
            os.remove(file)
        elif ftp_host:
            save_to_ftp(file, ftp_host, ftp_dir, ftp_user, ftp_pass)
            os.remove(file)


def get_tempo_data(file, include_labels):  # find elements in tree and deletes .xml file
    tempo_data_list = []
    tree = ET.parse(file)
    os.remove(file)  # delete .xml file
    root = tree.getroot()
    
    for worklog in root:
        issue_key = worklog.find('issue_key').text
        print(f'fetching worklog for {issue_key}')
        issue_summary = worklog.find('issue_summary').text
        issue_labels = None
        
        if include_labels:
            try:
                issue_labels = get_labels(issue_key)
            except RequestException as e:
                print(e)
        hours = worklog.find('hours').text
        work_date = worklog.find('work_date').text
        
        for user_details in worklog.iter('user_details'):
            full_name = user_details.find('full_name').text
       
        for issue_details in worklog.iter('issue_details'):
            issue_type = issue_details.find('type_name').text
            issue_status = issue_details.find('status_name').text
            project_key = issue_details.find('project_key').text
            project_name = issue_details.find('project_name').text
        period = datetime.strptime(work_date, '%Y-%m-%d').strftime('%m') + datetime.strptime(work_date, '%Y-%m-%d').strftime('%y')
        month = datetime.strptime(work_date, '%Y-%m-%d').strftime('%B')
        
        if re.search('.', hours):
            hours = hours.replace('.', ',')
        if re.search(';', issue_summary):
            issue_summary = issue_summary.replace(';', ',')
        if re.search('\t', issue_summary):
            issue_summary = issue_summary.replace('\t', ' ')
        tempo_data_list.append(TempoData(issue_key, issue_summary, hours, work_date, full_name, period, '', issue_type, issue_status, project_key, project_name, issue_labels))
    return tempo_data_list


def save_to_csv(file, tempo_data_list):
    print(f'saving "{file}"...')
    with open(file, 'w', newline='', encoding='utf-16') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(['Issue Key', 'Issue Summary', 'Hours', 'Work date', 'Full Name', 'Period', 'Month', 'Issue Type', 'Issue Status', 'Project Key', 'Project Name', 'Issue labels'])
        for tempo_data in tempo_data_list:
            writer.writerow([tempo_data.issue_key, tempo_data.issue_summary, tempo_data.hours, tempo_data.work_date, tempo_data.full_name, tempo_data.period, tempo_data.month, tempo_data.issue_type, tempo_data.issue_status, tempo_data.project_key, tempo_data.project_name, tempo_data.issue_labels])


def save_to_ftp(input_file, ftp_host, ftp_dir, ftp_user, ftp_pass):
    directories = ftp_dir.split('/')
    ftp_session = ftplib.FTP(ftp_host)
    ftp_session.login(ftp_user, ftp_pass)
    ftp_session.cwd('/')
    
    for directory in directories:  # checks if ftp_dest_dir exists, if not mkdir
        if directory in ftp_session.nlst():
            ftp_session.cwd(directory)
        else:
            ftp_session.mkd(directory)
            ftp_session.cwd(directory)
    
    with open(input_file, 'rb') as output_file:
        ftp_session.storbinary(f'STOR /{ftp_dir}/{input_file}', output_file)
    ftp_session.quit()


def send_email(input_file, recipient, date_from, date_to):
    message = MIMEMultipart()
    message['Subject'] = f'Tempo worklog report from {date_from} to {date_to}'
    message['From'] = SENDER_EMAIL
    message['To'] = recipient
    
    with open(input_file, 'rb') as file:
        attachment = MIMEApplication(file.read(), filename=os.path.basename(input_file))
    attachment['Content-Disposition'] = 'attachment; filename=%s' % os.path.basename(input_file)
    message.attach(attachment)
    
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, recipient, message.as_string())


def get_labels(issue_key):  # Be careful with large data ranges, it consumes a lot of resources
    response = requests.get(
        f'{JIRA_URL}/rest/api/latest/issue/{issue_key}?fields=labels',
        headers={'Authorization': 'Bearer ' + JIRA_TOKEN})
    
    if response.ok:
        json_data = json.loads(response.text)
        try:
            return ', '.join(json_data['fields']['labels'])
        except KeyError:  # field labels does not exist in issue
            return ''
    else:
        raise RequestException(f'request to Jira to get labels returned unexpected http code: {response.status_code}')


def _get_first_day_of_prev_month():
    return (datetime.now() + relativedelta(day=1, months=-1)).strftime('%Y-%m-%d')


def _get_first_day_of_curr_month():
    return (datetime.now() + relativedelta(day=1)).strftime('%Y-%m-%d')


def parse_args():
    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument('-b', '--beginDate', help='Date in format {-b yyyy-mm-dd}. It is not required, default value is first day of previous month', type=str, default=_get_first_day_of_prev_month())
    arg_parser.add_argument('-e', '--endDate', help='Date in format {-e yyyy-mm-dd}. It is not required, default value is last day of previous month', type=str, default=_get_first_day_of_curr_month())
    arg_parser.add_argument('-k', '--projectKey', help='Filter by project key in format {-pk projectKey}. It is not required, in this case script will generate billed hours for all project keys', type=str, default='')
    arg_parser.add_argument('-f', '--ftpHost', help='Dest FTP host in format {-fh host}. It is not required', type=str, default=False)
    arg_parser.add_argument('-d', '--ftpDir', help='Dest FTP dir in format {-fd path/to/dir}. It is not required', type=str)
    arg_parser.add_argument('-u', '--ftpUser', help='FTP username in format {-fu username}. It is not required', type=str)
    arg_parser.add_argument('-p', '--ftpPassword', help='FTP password in format {-fp password}. It is not required', type=str)
    arg_parser.add_argument('-r', '--recipient', nargs='+', help='Recipient in format {-r email} or {-r email1, email2}. It is not required', type=str, default=False)
    arg_parser.add_argument('-l', '--labels', action='store_true', help='include issue labels in reports', default=False)
    return arg_parser.parse_args()


if __name__ == '__main__':
    args = parse_args()
    try:
        create_tempo_worklog(str(args.beginDate), str(args.endDate), args.projectKey, args.ftpHost, args.ftpDir, args.ftpUser, args.ftpPassword, args.recipient, args.labels)
    except RequestException as e:
        print(e)
