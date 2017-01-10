#!/usr/bin/env python

import sys
import urllib2
import json
import re
import HTMLParser
import gspread
import gdata.docs.client
from oauth2client.service_account import ServiceAccountCredentials
from natsort import natsorted

writers = [ 'jose.m.moya@gmail.com', 'ana.lopezeiranova@gmail.com' ]
# name definitions
APP_NAME = 'PreK12Plaza'
name_spr = 'PreK12Plaza Learning Objects - Test'

# resources for credential
json_key = 'client_secrets.json'
scope = ['https://spreadsheets.google.com/feeds',
         'https://docs.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

filter = {
    'type': 'all',
    'domain': 'all',
    'standard': 'all',
    'language': 'all',
    'subject': 'all',
    }

fields = [ 'Search Grade/Subjects', 'Search Domain',
           'Grade',
           'Standard index',
           'Language',
           'Resource type', 'Resource name', 'Resource URL' ]

search = { 'PreK Math':
           [ '1',
             { 'All Domains': 'all'}
           ],
           'PreK ELA':
           [ '2',
             { 'L.PK - Language': '85' } 
           ],
           'PreK Social Studies':
           [ '122',
             { 'All Domains': 'all' }
           ],
           'PreK Spanish':
           [ '39',
             { 'All Domains': 'all' }
           ],
           'K Math':
           [ '6',
             { 'All Domains': 'all' }
           ],
           'K ELA':
           [ '7',
             { 'L.K - Language': '33',
               'RL.K - Reading Literature': '92',
               'RI.K - Reading Informational Text': '93',
               'RF.K - Reading Foundational Skills': '178' }
           ],
           '1st Grade Math':
           [ '12',
             { '1.OA - Operations and algebraic thinking': '10',
               '1.NBT - Number and operations in base ten': '11',
               '1.MD - Measurements and data': '12',
               '1.G - Geometry': '13' }
           ],
           '1st Grade ELA':
           [ '13',
             { 'L.1 - Language': '41',
               'RL.1 - Reading Literature': '91',
               'RI.1 - Reading Informational Text': '90',
               'RF.1 - Reading Foundational Skills': '179' }
           ],
           '2nd Grade Math':
           [ '18',
             { '2.OA - Operations and algebraic thinking': '14',
               '2.NBT - Number and operations in base ten': '15',
               '2.MD - Measurements and data': '16',
               '2.G - Geometry': '17' }
           ],
           '2nd Grade ELA':
           [ '19',
             { 'L.2 - Language': '42',
               'RL.2 - Reading Literature': '88',
               'RI.2 - Reading Informational Text': '89',
               'RF.2 - Reading Foundational Skills': '180' }
           ],
           '3rd Grade Math':
           [ '24',
             { '3.OA - Operations and algebraic thinking': '18',
               '3.NBT - Number and operations in base ten': '19',
               '3.NF - Number and operations - Fractions': '20',
               '3.MD - Measurements and data': '21',
               '3.G - Geometry': '22',
               'P - Parenting': '125' }
           ],
           '3rd Grade ELA':
           [ '25',
             { 'L.3 - Language': '43',
               'RL.3 - Reading Literature': '53',
               'RI.3 - Reading Informational Text': '62' }
           ],
           '4th Grade Math':
           [ '30',
             { '4.OA - Operations and algebraic thinking': '23',
               '4.NBT - Number and operations in base ten': '24',
               '4.NF - Number and operations - Fractions': '25',
               '4.MD - Measurements and data': '26',
               '4.G - Geometry': '27' }
           ],
           '4th Grade ELA':
           [ '31',
             { 'L.4 - Language': '44',
               'RL.4 - Reading Literature': '86',
               'RI.4 - Reading Informational Text': '87' }
           ],
           '4th Grade Science':
           [ '73',
             { 'All Domains': 'all' }
           ],
           '5th Grade Math':
           [ '36',
             { '5.OA - Operations and algebraic thinking': '28',
               '5.NBT - Number and operations in base ten': '29',
               '5.NF - Number and operations - Fractions': '30',
               '5.MD - Measurements and data': '31',
               '5.G - Geometry': '32' }
           ],
           '5th Grade ELA':
           [ '37',
             { 'L.5 - Language': '45',
               'RL.5 - Reading Literature': '61',
               'RI.5 - Reading Informational Text': '63' }
           ],
           '5th Grade Science':
           [ '74',
             { 'All Domains': 'all' }
           ],
           'Primary School Spanish':
           [ '38',
             { 'All Domains': 'all' }
           ],
           '6th Grade Math':
           [ '44',
             { 'All Domains': 'all' }
           ],
           '6th Grade ELA':
           [ '43',
             { 'L.6 - Language': '95',
               'RL.6 - Reading Literature': '118',
               'RI.6 - Reading Informational Text': '119' }
           ],
           '7th Grade Math':
           [ '46',
             { 'All Domains': 'all' }
           ],
           '7th Grade ELA':
           [ '45',
             { 'L.7 - Language': '116',
               'RL.7 - Reading Literature': '120',
               'RI.7 - Reading Informational Text': '121' }
           ],
           '8th Grade Math':
           [ '48',
             { 'All Domains': 'all' }
           ],
           '8th Grade ELA':
           [ '47',
             { 'L.8 - Language': '114',
               'RL.8 - Reading Literature': '122',
               'RI.8 - Reading Informational Text': '123' }
           ],
           'Middle School Spanish':
           [ '40',
             { 'All Domains': 'all' }
           ],
           '9th Grade Math':
           [ '64',
             { 'All Domains': 'all' }
           ],
           '9th Grade ELA':
           [ '49',
             { 'L.9-10 - Language': '115',
               'RL.9-10 - Reading Literature': '136',
               'RI.9-10 - Reading Informational Text': '132' }
           ],
           '10th Grade Math':
           [ '65',
             { 'All Domains': 'all' }
           ],
           '10th Grade ELA':
           [ '51',
             { 'L.9-10 - Language': '129',
               'RL.9-10 - Reading Literature': '137',
               'RI.9-10 - Reading Informational Text': '133' }
           ],
           '11th Grade ELA':
           [ '53',
             { 'L.11-12 - Language': '130',
               'RL.11-12 - Reading Literature': '138',
               'RI.11-12 - Reading Informational Text': '134' }
           ],
           '12th Grade ELA':
           [ '55',
             { 'L.11-12 - Language': '131',
               'RL.11-12 - Reading Literature': '139',
               'RI.11-12 - Reading Informational Text': '135' }
           ],
           'High School Algebra':
           [ '59',
             { 'All Domains': 'all' }
           ],
           'High School Geometry':
           [ '62',
             { 'All Domains': 'all' }
           ],
           'High School Functions':
           [ '60',
             { 'All Domains': 'all' }
           ],
           'High School Statistics and Probability':
           [ '63',
             { 'All Domains': 'all' }
           ],
           'High School Spanish':
           [ '41',
             { 'All Domains': 'all' }
           ],
           'High School Number and Quantity':
           [ '58',
             { 'All Domains': 'all' }
           ],
           'Parent - Computer Skills':
           [ '119',
             { 'All Domains': 'all' }
           ],
           'Parenting':
           [ '116',
             { 'All Domains': 'all' }
           ],
           'Parent - US Citizenship':
           [ '118',
             { 'All Domains': 'all' }
           ],
           'Administrator - Professional Development':
           [ '125',
             { 'All Domains': 'all' }
           ],
           'Teacher - Professional Development':
           [ '126',
             { 'All Domains': 'all' }
           ],
           'Teacher - Resources':
           [ '127',
             { 'All Domains': 'all' }
           ]
}

# create goole data docs client
client = gdata.docs.client.DocsClient(source=APP_NAME)
#client.http_client.debug = True
client.http_client.debug = False

# create credentials
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_key,
                                                               scope)
auth_token = gdata.gauth.OAuth2TokenFromCredentials(credentials)

# authorise
auth_token.authorize(client)

# authorise gspread
gc = gspread.authorize(credentials)

try:
    # try to open the spreadsheet
    wb = gc.open(name_spr)
except:
    # create document as spreadsheet
    document = gdata.docs.data.Resource(type='spreadsheet', title=name_spr)
    document = client.CreateResource(document)
    for email in writers:
        # add ACL to  spreadsheet
        acl_entry = gdata.docs.data.AclEntry(
            scope=gdata.acl.data.AclScope(value=email, type='user'),
            role=gdata.acl.data.AclRole(value='writer'),
        )
        client.AddAclEntry(document, acl_entry, send_notifications=False)
    pass

wb = gc.open(name_spr)


def prek12plaza_get (grade, domains):
    results = []
    filter['n'] = 1
    filter['subject'] = search[grade][0]
    for domain in natsorted(domains):
        filter['domain'] = domains[domain]
        str = urllib2.urlopen("https://www.prek12plaza.com/bank_educators/get_ajax?sEcho=2&iColumns=1&sColumns=&iDisplayStart=0&iDisplayLength={n}&mDataProp_0=0&a=get_page&keyword=undefined&type={type}&domain={domain}&standard={standard}&language={language}&subject={subject}&assign=all".format(**filter)).read()
        data=json.JSONDecoder().decode(str)

        filter['n'] = int(data['iTotalRecords'])
        print "Grade %s, domain %s : %d objects" % (grade, domain, filter['n'])
        str = urllib2.urlopen("https://www.prek12plaza.com/bank_educators/get_ajax?sEcho=2&iColumns=6&sColumns=&iDisplayStart=0&iDisplayLength={n}&mDataProp_0=0&mDataProp_1=1&mDataProp_2=2&mDataProp_3=3&mDataProp_4=4&mDataProp_5=5&a=get_page&keyword=undefined&type={type}&domain={domain}&standard={standard}&language={language}&subject={subject}&assign=all".format(**filter)).read()
        data = json.JSONDecoder().decode(str)

        objs = data['aaData']

        html = HTMLParser.HTMLParser()

        for e in objs:
            for fld in fields:
                e[fld] = ''
            e['Search Grade/Subjects'] = grade
            e['Search Domain'] = domain
            e['Standard index'] = e['3'].split(' ', 1)[0]
            e['Resource name'] = html.unescape(re.sub('<.*?>', '', e['1']))
            e['Grade'] = html.unescape(re.sub('<.*?>', '', e['2']))
            e['Language'] = html.unescape(e['4'])
            e['Resource type'] = html.unescape(e['5'])
            m = re.match('.*<a href="(.*)" target="_blank">.*', e['1'])
            e['Resource URL'] = m.groups()[0]
        results.extend(objs)

    print '  ...processed'
    return results

def gspread_update_grade (ws, row, objs):
    # authorise gspread
    gc = gspread.authorize(credentials)

    for j in range(len(objs)):
        e = objs[j]
        for i in range(len(fields)):
            ws.update_cell(j + 2, i + 1, e[fields[i]])
    print "  ...done"

def main(argv):
    # select worksheet
    ws_name = "All Learning Objects"
    try:
        ws = wb.worksheet(ws_name)
    except:
        ws = wb.add_worksheet(title=ws_name, rows='5000', cols='20')
        pass
    
    for i in range(len(fields)):
        ws.update_cell(1, i + 1, fields[i])

    row = 2
    domain = ''
    if len(argv) > 1:
        search = [ argv[1] ]
        row = int(argv[2])
        print "Grades: ", search

    if len(argv) > 3:
        domains = argv[3:]
        print "Domains: ", domains

    for grade in natsorted(search):
        if len(domains) == 0:
            domains = search[grade][1]
        objs = prek12plaza_get (grade, domains)
        gspread_update_grade (ws, row, objs)
        row += len(objs)

if __name__ == '__main__':
    main(sys.argv)
