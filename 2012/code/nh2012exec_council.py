# -*- coding: utf-8 -*-
"""
Created on Sun Sep 25 18:19:18 2016

@author: mike
"""

import requests, tempfile
url='http://sos.nh.gov/WorkArea/DownloadAsset.aspx?id=28334'
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

results_dict=dict()

import csv
csvfile=open(rep_dir+'/20121106__nh__general__executive__council__town.csv','wb')
csvwriter=csv.writer(csvfile)
csvwriter.writerow(['town',
                    'county', 
                    'office', 
                    'district', 
                    'party', 
                    'candidate',
                    'winner',
                    'votes'])

import xlrd, re
wb=xlrd.open_workbook(tmp_file,formatting_info=True)
sheets = wb.sheet_names()

for sheet in sheets:
    ws = wb.sheet_by_name(sheet)
    try:
        district=re.search('[0-9]+',sheet).group(0)
    except:
        district=''
    results_dict[district]=dict()
    candidate_dict=dict()
    cols_dict=dict()
    start_flag=0
    stop_flag=0
    for row in range(ws.nrows):
        try:
            if(any(['Scatter' in str(x.value) for x in ws.row(row)])):
                scatter_test=['Scatter' in str(x.value) for x in ws.row(row)]
                scatter_cols=[i for i, x in enumerate(scatter_test) if x]
                col_offset=[x-scatter_cols[0] for x in scatter_cols]
                start_flag=1
                stop_flag=0
                for col in range(1,scatter_cols[0]+1):
                    value=ws.cell(row,col).value
                    if('Scatter' in value):
                        candidate_dict['Scatter']=dict()
                        candidate_dict['Scatter']['Party']=''
                        candidate='Scatter'
                    elif('write-in' in value):
                        candidate=re.search('.*(?=\swrite-in)',value).group(0)
                        candidate_dict[candidate]=dict()
                        candidate_dict[candidate]['Party']='IND'
                    else:
                        candidate=re.search('.*(?=\s*,)',value).group(0)
                        candidate=re.sub('\s+',' ',candidate)
                        party_code=re.search('(?<=,).*',value).group(0)
                        party_code=re.sub('\s+','',party_code)
                        candidate_dict[candidate]=dict()
#                        if(party_code=='r'):
#                            candidate_dict[candidate]['Party']='REP'
#                        elif(party_code=='d'):
#                            candidate_dict[candidate]['Party']='DEM'
#                        elif(party_code=='lib'):
#                            candidate_dict[candidate]['Party']='LIB'
#                        elif(party_code=='con'):
#                            candidate_dict[candidate]['Party']='CON'
                        candidate_dict[candidate]['Party']=party_code.upper()
                    cols_dict[col]=candidate
        except:
            pass    
        if(start_flag==2 and stop_flag==0):
            for offset in col_offset:
#                if('Totals' in str(ws.cell(row,offset).value)):
#                    stop_flag=1
#                else:
                town=ws.cell(row,offset).value
                town=re.sub('\s+\Z','',town)
                town=re.sub('\*','',town)
                if(town=='Atkinson & Gilmanton Academy Gt'):
                    town='At. & Gil. Academy Grant'
                if(town=='Thompson & Meserve\'s Purchase'):
                    town='Thompson & Meserve\'s Pur.'
                if(town=='Martins\' Location'):
                    town='Martin\'s Location'
                if(town.lower()!='totals' and town!='' and 'correct' not in town):
                    results_dict[district][town]=dict()
                    #results_dict[district][town]['county']=county_dict[town]
                    for col in range(offset+1,offset+len(candidate_dict)+1):
                        candidate=cols_dict[col-offset]
                        value=ws.cell(row,col).value
                        if(value=='' or value==' '):
                            results_dict[district][town][candidate]=0
                        else:
                            results_dict[district][town][candidate]=int(value)
                elif(town.lower()=='totals'):
                    print district, row, ws.row(row),cols_dict.keys()
#                    for col in cols_dict:
#                        candidate=cols_dict[col]
#                        if(candidate in candidate_dict):
#                            fmt=wb.xf_list[ws.cell_xf_index(row,col)]
#                            bold=wb.font_list[fmt.font_index].bold
#                            if(bold==1):
#                                candidate_dict[candidate]['Winner']=True
#                            else:
#                                candidate_dict[candidate]['Winner']=False
                    winning_total=max([int(ws.cell(row,col).value) for col in range(offset+1,offset+len(candidate_dict)+1)])
                    winning_col=[col for col in range(offset+1,offset+len(candidate_dict)+1)
                                 if int(ws.cell(row,col).value)==winning_total][0]
                    for col in range(offset+1,offset+len(candidate_dict)+1):
                        candidate=cols_dict[col-offset]
                        if(col==winning_col):
                            candidate_dict[candidate]['Winner']=True
                        else:
                            candidate_dict[candidate]['Winner']=False
        elif(start_flag==1):
            start_flag=2
    # Clean up multiple wards into one set of results per town
    wards=[x for x in results_dict[district].keys() if 'Ward' in x]
    towns=set([re.search('.*(?=\sWard)',x).group(0) for x in wards])
    for town in towns:
        results_dict[district][town]=dict()
        town_wards=[x for x in wards if town in x]
    #    county=results_dict[town_wards[0]]['county']
    #    results_dict[town]['county']=county
        for candidate in candidate_dict:
            results_dict[district][town][candidate]=sum([results_dict[district][x][candidate] for x in town_wards])
        for ward in town_wards:
            del results_dict[district][ward]
    
    for town in results_dict[district].keys():
        results_dict[district][town]['county']=county_dict[town]
    for candidate in sorted(candidate_dict):
        for town in sorted(results_dict[district].keys()):
            csvwriter.writerow([town,
                                results_dict[district][town]['county'],
                                'Executive Council',
                                district,
                                candidate_dict[candidate]['Party'],
                                candidate,
                                candidate_dict[candidate]['Winner'],
                                results_dict[district][town][candidate]
                                ])        
    #print district,candidate_dict.keys()

csvfile.close()