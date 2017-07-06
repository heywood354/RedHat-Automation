#!/usr/bin/env python
from __future__ import print_function
import sys
import requests
from datetime import datetime, timedelta
import pandas
import json
import urllib2
import time

start_time = time.time()

#api url provided by Red Hat
API_HOST = 'https://access.redhat.com/labs/securitydataapi'

# print to see that program started
print('Gathering Data...\n')

#function to send query to Red Hat api
def get_data(query):

    full_query = API_HOST + query
    r = requests.get(full_query)

    if r.status_code != 200:
        print('ERROR: Invalid request; returned {} for the following '
              'query:\n{}'.format(r.status_code, full_query))
        sys.exit(1)

    if not r.json():
        print('No data returned with the following query:')
        print(full_query)
        sys.exit(0)

    return r.json()

#sets endpoint to pull cvrf data in json format from Red Hat api. Used later 
endpoint = '/cvrf.json'

#todays date minus 40 days, can change this to user input at a later time
date = datetime.now() - timedelta(days=3)

#setting parameters, this one only searching for data after 40 days ago.
params = 'after=' + str(date.date())

#sending query  with endpoint and parameters
data = get_data(endpoint + '?' + params)


daddy_list = []
baby_list = []

for url in range(0, len(data)): #change 3 back to len(data)

    #opening the url that is in "data" in order to access all JSON information we need - storing it all to url_data
    response = urllib2.urlopen(data[url]['resource_url'])
    url_data = json.load(response) #type is dictionary so we access it like a dictionary

    #grabbing the information we need from the JSON file
    rhsa = url_data['cvrfdoc']['document_tracking']['identification']['id']
    version = url_data['cvrfdoc']['document_tracking']['version']
    advisory_title = url_data['cvrfdoc']['document_title']
    advisory_title_trimmed = advisory_title[27:]
    severity = url_data['cvrfdoc']['aggregate_severity']
    impact = url_data['cvrfdoc']['document_notes']['note']
    try:
        restart = url_data['cvrfdoc']['discovery_date']
    except:
        restart = "May Require Restart"
        pass

    # affected_software = advisory_title_trimmed.split(' ', 1)[0]
    # affected_software = url_data['cvrfdoc']['document_notes']['note']
    JSONdict = {'RHSA': rhsa, 'ADVISORY_TITLE_TRIMMED': advisory_title_trimmed + " " + rhsa + "-" + version, 'SEVERITY': severity,
                    'VULNERABILITY_IMPACT': impact, 'RESTART': restart, 'AFFECTED': advisory_title}
    daddy_list.append(JSONdict)

print(" ")
print('Building Excel Document...')
print(" ")

#get the month for spreadsheet naming convention
today = datetime.today()
num_month = str(today.month)
month_dict = {'1': 'Jan', '2': 'Feb','3': 'March', '4': 'April','5': 'May', '6': 'June','7': 'July', '8': 'Aug','9': 'Sept', '10': 'Oct','11': 'Nov', '12': 'Dec'}
month = month_dict[num_month]

#create workbook
workbook_name = "MS_RHEL_Annoucement_Worksheet_" + str(today.day) + "_" + month + "_" + str(today.year) +".xlsx"


#reorder the columns, if a comlumn is not listed, it will not be displayed.
# myColumns = ['RHSA', 'released_packages', 'severity', 'released_on', 'resource_url', 'package']
# myColumns = ['version']
myColumns = ['RHSA', 'ADVISORY_TITLE_TRIMMED', 'SEVERITY', 'VULNERABILITY_IMPACT', 'RESTART', 'AFFECTED']

#creates a panda DataFrame from the data pulled from the API.
advisoryDF = pandas.DataFrame(daddy_list, columns = myColumns) #used to be "data", not daddy_list

#created excel speadsheet
writer = pandas.ExcelWriter(workbook_name, engine='xlsxwriter')

#adds data frame to excel workbook on sheet named sheet_test
advisoryDF.to_excel(writer, index=False, sheet_name = month + " RHEL_Analysis")

workbook  = writer.book
rhel_analysis_worksheet = writer.sheets[month + " RHEL_Analysis"]

bold_format = workbook.add_format({'bold': True, 'bg_color': '#A6A6A6', 'font_name': 'Verdana', 'font_size': '10', 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'border': True})
bold2_format = workbook.add_format({'bold': True, 'bg_color': '#CCCCCC', 'font_name': 'Verdana', 'font_size': '8.5', 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'border': True})
bold3_format = workbook.add_format({'bold': True, 'bg_color': '#FFEB9C', 'font_name': 'Calibri', 'font_size': '11', 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'border': True})

#add column names to rhel_analysis_worksheet
rhel_analysis_worksheet.write("A1", "Advisory ID", bold_format)
rhel_analysis_worksheet.write("B1", "Advisory Title and RHSA Number", bold_format)
rhel_analysis_worksheet.write("C1", "Maximum Severity Rating ", bold_format)
rhel_analysis_worksheet.write("D1", "Vulnerability Impact", bold_format)
rhel_analysis_worksheet.write("E1", "Restart Requirement", bold_format)
rhel_analysis_worksheet.write("F1", "Affected Software", bold_format)
rhel_analysis_worksheet.write("G1", "Applicable Satellite 5", bold2_format)
rhel_analysis_worksheet.write("H1", "Applicable Satellite 6", bold2_format)
rhel_analysis_worksheet.write("I1", "Total Applicable Satellite 5 & 6", bold2_format)
rhel_analysis_worksheet.write("J1", "Deploy", bold2_format)
rhel_analysis_worksheet.write("K1", "Announce", bold2_format)
rhel_analysis_worksheet.write("L1", "Notes", bold3_format)

#Format Columns to be centered
rhel_analysis_worksheet.set_column('A:A', 15)
rhel_analysis_worksheet.set_column('B:B', 15)
rhel_analysis_worksheet.set_column('C:C', 15)
rhel_analysis_worksheet.set_column('D:D', 15)
rhel_analysis_worksheet.set_column('E:E', 15)
rhel_analysis_worksheet.set_column('F:F', 15)
rhel_analysis_worksheet.set_column('G:G', 15)
rhel_analysis_worksheet.set_column('H:H', 15)
rhel_analysis_worksheet.set_column('I:I', 15)
rhel_analysis_worksheet.set_column('J:J', 15)
rhel_analysis_worksheet.set_column('K:K', 15)
rhel_analysis_worksheet.set_column('L:L', 60)

server_analysis_worksheet = workbook.add_worksheet(name = month + " Server_Analysis")
workstation_analysis_worksheet = workbook.add_worksheet(name = month + " Workstation_Analysis")

#add column names to server_analysis_worksheet
server_analysis_worksheet.write("A1", "Bulletin ID", bold_format)
server_analysis_worksheet.write("B1", "Bulletin Title and KB Number", bold_format)
server_analysis_worksheet.write("C1", "Maximum Severity Rating ", bold_format)
server_analysis_worksheet.write("D1", "Vulnerability Impact", bold_format)
server_analysis_worksheet.write("E1", "Restart Requirement", bold_format)
server_analysis_worksheet.write("F1", "Affected Software", bold_format)
server_analysis_worksheet.write("G1", "Applicable", bold2_format)
server_analysis_worksheet.write("H1", "Available In LANDesk Y/N", bold2_format)
server_analysis_worksheet.write("I1", "Deploy DBHDS Y/N", bold2_format)
server_analysis_worksheet.write("J1", "Deploy Y/N", bold2_format)
server_analysis_worksheet.write("K1", "Announce Y/N", bold2_format)
server_analysis_worksheet.write("L1", "Notes", bold3_format)

#Format Columns to be centered
server_analysis_worksheet.set_column('A:A', 15)
server_analysis_worksheet.set_column('B:B', 15)
server_analysis_worksheet.set_column('C:C', 15)
server_analysis_worksheet.set_column('D:D', 15)
server_analysis_worksheet.set_column('E:E', 15)
server_analysis_worksheet.set_column('F:F', 15)
server_analysis_worksheet.set_column('G:G', 15)
server_analysis_worksheet.set_column('H:H', 15)
server_analysis_worksheet.set_column('I:I', 15)
server_analysis_worksheet.set_column('J:J', 15)
server_analysis_worksheet.set_column('K:K', 15)
server_analysis_worksheet.set_column('L:L', 60)

#add column names to workstation_analysis_worksheet
workstation_analysis_worksheet.write("A1", "Bulletin ID", bold_format)
workstation_analysis_worksheet.write("B1", "Bulletin Title and KB Number", bold_format)
workstation_analysis_worksheet.write("C1", "Maximum Severity Rating ", bold_format)
workstation_analysis_worksheet.write("D1", "Vulnerability Impact", bold_format)
workstation_analysis_worksheet.write("E1", "Restart Requirement", bold_format)
workstation_analysis_worksheet.write("F1", "Affected Software", bold_format)
workstation_analysis_worksheet.write("G1", "Applicable", bold2_format)
workstation_analysis_worksheet.write("H1", "Available In LANDesk Y/N", bold2_format)
workstation_analysis_worksheet.write("I1", "Deploy DBHDS Y/N", bold2_format)
workstation_analysis_worksheet.write("J1", "Deploy Y/N", bold2_format)
workstation_analysis_worksheet.write("K1", "Announce Y/N", bold2_format)
workstation_analysis_worksheet.write("L1", "Notes", bold3_format)

#Format Columns to be centered
workstation_analysis_worksheet.set_column('A:A', 15)
workstation_analysis_worksheet.set_column('B:B', 15)
workstation_analysis_worksheet.set_column('C:C', 15)
workstation_analysis_worksheet.set_column('D:D', 15)
workstation_analysis_worksheet.set_column('E:E', 15)
workstation_analysis_worksheet.set_column('F:F', 15)
workstation_analysis_worksheet.set_column('G:G', 15)
workstation_analysis_worksheet.set_column('H:H', 15)
workstation_analysis_worksheet.set_column('I:I', 15)
workstation_analysis_worksheet.set_column('J:J', 15)
workstation_analysis_worksheet.set_column('K:K', 15)
workstation_analysis_worksheet.set_column('L:L', 60)




#saves .xlsx file
writer.save()

#test print to ensure program is finshed.
print('Task Complete!')
print(" ")
print("This task took %s seconds to complete" % (time.time() - start_time))
