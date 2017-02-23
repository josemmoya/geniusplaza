#!/usr/bin/env python

import sys
import urllib2
import json
import re
import HTMLParser
from natsort import natsorted
import xlsxwriter, xlrd
import argparse

name_spr = 'GeniusPlaza Learning Objects'

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


# URLs of previously existing learning objects
previous = []

def geniusplaza_get (grade, domains):
    results = []
    filter['n'] = 1
    filter['subject'] = search[grade][0]
    for domain in natsorted(domains):
        filter['domain'] = domains[domain]
        str = urllib2.urlopen("https://www.geniusplaza.com/bank_educators/get_ajax?sEcho=2&iColumns=1&sColumns=&iDisplayStart=0&iDisplayLength={n}&mDataProp_0=0&a=get_page&keyword=undefined&type={type}&domain={domain}&standard={standard}&language={language}&subject={subject}&assign=all".format(**filter)).read()
        data=json.JSONDecoder().decode(str)

        filter['n'] = int(data['iTotalRecords'])
        print "Grade %s, domain %s : %d objects" % (grade, domain, filter['n'])
        str = urllib2.urlopen("https://www.geniusplaza.com/bank_educators/get_ajax?sEcho=2&iColumns=6&sColumns=&iDisplayStart=0&iDisplayLength={n}&mDataProp_0=0&mDataProp_1=1&mDataProp_2=2&mDataProp_3=3&mDataProp_4=4&mDataProp_5=5&a=get_page&keyword=undefined&type={type}&domain={domain}&standard={standard}&language={language}&subject={subject}&assign=all".format(**filter)).read()
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
    ignored = 0
    for j in range(len(objs)):
        e = objs[j]
        if e['Resource URL'] in previous:
            ignored += 1
            continue
        for i in range(len(fields)):
            ws.write(row, i, e[fields[i]])
        row += 1
    print "  ...done, ignored %d previous objects" % (ignored)

def main(argv):

    parser = argparse.ArgumentParser(description='Extract learning objects from GeniusPlaza.')
    parser.add_argument('--diff', help="difference with previous excel file", action="append")
    
    args = parser.parse_args()

    for f in args.diff:
        try:
            nprev = 0
            prev_wb = xlrd.open_workbook(f)
            for s in prev_wb.sheets():
                for row in range(1, s.nrows):
                    for col in range(s.ncols):
                        if s.cell(0, col).value == 'Resource URL':
                            url = s.cell(row, col).value.replace("www.prek12plaza.com","www.geniusplaza.com")
                            previous.append(url)
                            nprev += 1
            print "Added %d previous learning objects from %s" % (nprev, f)
        except:
            pass

    # Create an new Excel file and add a worksheet.
    wb = xlsxwriter.Workbook(name_spr + '.xlsx')

    for grade in natsorted(search):
        # select worksheet
        ws = wb.add_worksheet(grade[:30])

        for i in range(len(fields)):
            ws.write(0, i, fields[i])

        domains = search[grade][1]
        objs = geniusplaza_get (grade, domains)
        gspread_update_grade (ws, 1, objs)
        #row += len(objs)

    wb.close()
    
        
if __name__ == '__main__':
    main(sys.argv)

