# -*- coding: utf-8 -*-
"""
Created on Thu Sep 22 20:54:00 2016

@author: mjguidry
"""

import requests, tempfile
url='http://sos.nh.gov/WorkArea/DownloadAsset.aspx?id=29959'
resp = requests.get(url)

tempdir=tempfile.tempdir
tmp_file=tempdir+'/temp_nh.xls'
rep_dir='../'


output = open(tmp_file, 'wb')
output.write(resp.content)
output.close()

import xlrd, re
wb=xlrd.open_workbook(tmp_file)
sheets = wb.sheet_names()

import pickle
f=open('county.pkl','rb')
county_dict=pickle.load(f)
f.close()

results_dict=dict()
cols_dict=dict()
candidates=['Guinta','Shea-Porter','Kelly','Scatter']
candidate_dict={'Guinta':{'name':'Frank C. Guinta' ,
                          'party':'R',
                          'winner':False},
                'Shea-Porter' :{'name':'Carol Shea-Porter',
                                'party':'D',
                                'winner':True},
                'Kelly' :{'name':'Brendan Kelly',
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
            if('Totals' in ws.cell(row,0).value):
                stop_flag=1
            else:
                town=ws.cell(row,0).value
                town=re.sub('\*','',town)
                results_dict[town]=dict()
#                results_dict[town]['county']=county_dict[town]
                for col in range(1,ws.ncols):
                    candidate=cols_dict[col]
                    value=ws.cell(row,col).value
                    if(value=='' or value==' '):
                        results_dict[town][candidate]=0
                    else:
                        results_dict[town][candidate]=int(value)
        try:
            if(any(['Scatter' in str(x.value) for x in ws.row(row)])):
                value=ws.cell(row,0).value
                start_flag=1
                stop_flag=0
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
towns=set([re.search('.*(?=\s-\sWard)',x).group(0) for x in wards])
for town in towns:
    results_dict[town]=dict()
    town_wards=[x for x in wards if town in x]
#    county=results_dict[town_wards[0]]['county']
#    results_dict[town]['county']=county
    for candidate in candidates:
        results_dict[town][candidate]=sum([results_dict[x][candidate] for x in town_wards])
    for ward in town_wards:
        del results_dict[ward]

for town in results_dict.keys():
    results_dict[town]['county']=county_dict[town]

import csv
csvfile=open(rep_dir+'/20121106__nh__general__house__1__town.csv','wb')
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
                            'U.S. House',
                            '1',
                            candidate_dict[candidate]['party'],
                            candidate_dict[candidate]['name'],
                            candidate_dict[candidate]['winner'],
                            results_dict[town][candidate]
                            ])

csvfile.close()

