# -*- coding: utf-8 -*-
"""
Created on Thu Sep 22 20:54:00 2016

@author: mjguidry
"""

import requests, tempfile
url='http://sos.nh.gov/WorkArea/DownloadAsset.aspx?id=28331'
resp = requests.get(url)

tempdir=tempfile.tempdir
tmp_file=tempdir+'/temp_nh.xls'
rep_dir='../'


output = open(tmp_file, 'wb')
output.write(resp.content)
output.close()

# Get town to county matchups from President run
import pickle
f=open('county.pkl','rb')
county_dict=pickle.load(f)
f.close()

import xlrd, re
wb=xlrd.open_workbook(tmp_file)
sheets = wb.sheet_names()

results_dict=dict()
cols_dict=dict()
candidates=['Lamontagne','Hassan','Babiarz','Scatter']
candidate_dict={'Lamontagne':{'name':'Ovide Lamontagne' ,
                              'party':'REP',
                              'winner':False},
                'Hassan' :{'name':'Maggie Hassan',
                           'party':'DEM',
                           'winner':True},
                'Babiarz' :{'name':'John J. Babiarz',
                            'party':'LIB',
                            'winner':False},
                'Scatter' :{'name':'Scatter',
                            'party':'',
                            'winner':False}}

for sheet in sheets:
    ws = wb.sheet_by_name(sheet)
    start_flag=0
    stop_flag=0
    for row in range(ws.nrows):
        if(start_flag==1 and stop_flag==0):
            if('TOTALS' in ws.cell(row,0).value):
                stop_flag=1
            else:
                town=ws.cell(row,0).value
                town=re.sub('\s+\Z','',town)
                town=re.sub('\*','',town)
                results_dict[town]=dict()
#                results_dict[town]['county']=county
                for col in range(1,ws.ncols):
                    candidate=cols_dict[col]
                    value=ws.cell(row,col).value
                    if(value=='' or value==' '):
                        results_dict[town][candidate]=0
                    else:
                        results_dict[town][candidate]=int(value)
        try:
            if('County' in ws.cell(row,0).value):
                value=ws.cell(row,0).value
                start_flag=1
                stop_flag=0
                county=re.search('.*(?=\sCounty)',value).group(0)
                for col in range(1,ws.ncols):
                    candidate=[x for x in candidates if x in ws.cell(row,col).value][0]
                    cols_dict[col]=candidate
        except:
            pass
        
# Debug print statements
# print 'Romney results ', sum([results_dict[x]['Romney'] for x in results_dict.keys()])
# print 'Obama results ', sum([results_dict[x]['Obama'] for x in results_dict.keys()])

# Clean up multiple wards into one set of results per town
wards=[x for x in results_dict.keys() if 'Ward' in x]
towns=set([re.search('.*(?=\sWard)',x).group(0) for x in wards])
for town in towns:
    results_dict[town]=dict()
    town_wards=[x for x in wards if town in x]
    for candidate in candidates:
        results_dict[town][candidate]=sum([results_dict[x][candidate] for x in town_wards])
    for ward in town_wards:
        del results_dict[ward]

for town in results_dict.keys():
    results_dict[town]['county']=county_dict[town]

import csv
csvfile=open(rep_dir+'/20121106__nh__general__town__gov.csv','wb')
csvwriter=csv.writer(csvfile)
csvwriter.writerow(['town',
                    'county', 
                    'office', 
                    'district', 
                    'party', 
                    'candidate',
                    'winner',
                    'votes'])
for candidate in candidates:
    for town in sorted(results_dict.keys()):
        csvwriter.writerow([town,
                            results_dict[town]['county'],
                            'Governor',
                            '',
                            candidate_dict[candidate]['party'],
                            candidate_dict[candidate]['name'],
                            candidate_dict[candidate]['winner'],
                            results_dict[town][candidate]
                            ])

csvfile.close()
        