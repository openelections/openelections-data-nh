# -*- coding: utf-8 -*-
"""
Created on Tue Sep 27 21:30:39 2016

@author: mike
"""

import xlrd, re
import requests, tempfile
urls=['http://sos.nh.gov/WorkArea/DownloadAsset.aspx?id=28345',
      'http://sos.nh.gov/WorkArea/DownloadAsset.aspx?id=28346',
      'http://sos.nh.gov/WorkArea/DownloadAsset.aspx?id=28347',
      'http://sos.nh.gov/WorkArea/DownloadAsset.aspx?id=28348',
      'http://sos.nh.gov/WorkArea/DownloadAsset.aspx?id=28349',
      'http://sos.nh.gov/WorkArea/DownloadAsset.aspx?id=28350',
      'http://sos.nh.gov/WorkArea/DownloadAsset.aspx?id=28351',
      'http://sos.nh.gov/WorkArea/DownloadAsset.aspx?id=28352',
      'http://sos.nh.gov/WorkArea/DownloadAsset.aspx?id=28353',
      'http://sos.nh.gov/WorkArea/DownloadAsset.aspx?id=28354'
      ]

tempdir=tempfile.tempdir
tmp_file=tempdir+'/temp_nh.xls'
rep_dir='../'

# Get town to county matchups from President run
import pickle
f=open('county.pkl','rb')
county_dict=pickle.load(f)
f.close()

results_dict=dict()

import csv
csvfile=open(rep_dir+'/20121106__nh__general__town__state__rep.csv','wb')
csvwriter=csv.writer(csvfile)
csvwriter.writerow(['town',
                    'county', 
                    'office', 
                    'district', 
                    'party', 
                    'candidate',
                    'winner',
                    'votes'])


for url in urls:
    resp = requests.get(url)
    output = open(tmp_file, 'wb')
    output.write(resp.content)
    output.close()
    header_row=0
    wb=xlrd.open_workbook(tmp_file,formatting_info=True)
    sheets = wb.sheet_names()
    for sheet in sheets:
        start_flag=0
        stop_flag=1
        ws = wb.sheet_by_name(sheet)
        for row in range(ws.nrows):
            if(any(['District' in str(x.value) for x in ws.row(row)]) or
               any(['Scatter' in str(x.value) for x in ws.row(row)]) or
               any(['Recount' in str(x.value) for x in ws.row(row)[1:]])):
                if(any(['District' in str(x.value) for x in ws.row(row)])):
                    candidate_dict=dict()
                cols_dict=dict()
                start_flag=1  
                stop_flag=0
                for col in range(0,ws.ncols):
                    value=ws.cell(row,col).value
                    if('Dist' in value):
                        district=re.search('[0-9]+(?=\s+\()',value).group(0)
                        print district
                        results_dict[district]=dict()
                    elif('Dover Ward 15' in value):
                        district='15'
                        print district
                        results_dict[district]=dict()
                    elif('Scatter' in value or 'Scattter' in value):
                        candidate_dict['Scatter']=dict()
                        candidate_dict['Scatter']['Party']=''
                        candidate='Scatter'
                        cols_dict[col]=candidate
                    elif('Recount' in value):
                        if(col>0):
                            cols_dict[col]=candidate
                            del cols_dict[col-1]
                    elif('write-in' in value):
                        candidate=re.search('.*(?=\swrite-in)',value).group(0)
                        candidate_dict[candidate]=dict()
                        candidate_dict[candidate]['Party']='IND'
                        cols_dict[col]=candidate
                    elif(value!='' and not value.isspace()):
                        candidate=re.search('.*(?=\s*,)',value).group(0)
                        party_code=value.split(',')[-1]
                        party_code=re.sub('\s+','',party_code)
                        candidate_dict[candidate]=dict()
                        if(party_code=='r'):
                            candidate_dict[candidate]['Party']='REP'
                        elif(party_code=='rtc'):
                            candidate_dict[candidate]['Party']='RTC'
                        elif(party_code=='d'):
                            candidate_dict[candidate]['Party']='DEM'
                        elif(party_code=='lib'):
                            candidate_dict[candidate]['Party']='LIB'
                        elif(party_code=='con'):
                            candidate_dict[candidate]['Party']='CON'
                        elif(party_code=='und' or party_code=='i' or 
                        party_code=='i.m.' or party_code=='d&r' or 
                        party_code=='ind' or party_code=='i&r'  or party_code=='u&r'):
                            candidate_dict[candidate]['Party']='IND'
                        cols_dict[col]=candidate
                #print candidate_dict.keys()
                header_row=0
            if(start_flag==2 and stop_flag==0):
                town=ws.cell(row,0).value
                town=re.sub('\s+\Z','',town)
                town=re.sub('\*','',town)
                if(town=='Atkinson & Gilmanton Academy Gt' or 
                    town=='Atkinson and Gilmanton Ac. Gt.' or 
                    town=='Atkinson & Gilmanton Ac Gt'):
                    town='At. & Gil. Academy Grant'
                if(town=='Thompson & Meserve\'s Purchase' or 
                    town=='Thomp and Mes\'s Pur'):
                    town='Thompson & Meserve\'s Pur.'
                if(town=='Martins\' Location'):
                    town='Martin\'s Location'
                if(town=='Low and Burbank\'s Grant' or 
                    town=='Low and Burbank\'s Gt.'):
                    town='Low & Burbank\'s Grant'
                if(town=='Recount'):
                    town=prev_town
                if(town.lower()!='totals' and town!='' and 'correct' not in town):
                    if(town not in results_dict[district]):
                        results_dict[district][town]=dict()
                    for col in cols_dict:
                        candidate=cols_dict[col]
                        value=ws.cell(row,col).value
#                        if(value=='' or str(value).isspace()):
#                            results_dict[district][town][candidate]=0
                        try:
                            results_dict[district][town][candidate]=int(value)
                        except:
                            results_dict[district][town][candidate]=0
                    prev_town=town
                else:
                    stop_flag=1
            if(start_flag==1 and stop_flag==0):
                start_flag=2
            try:
                row_test=str(ws.cell(row+1,0).value)=='' and start_flag==2                    
            except:
                row_test=False
            if('totals' in str(ws.cell(row,0).value).lower() or row_test):
                    # Clean up multiple wards into one set of results per town
                wards=[x for x in results_dict[district].keys() if 'Ward' in x or 'Wd' in x]
                towns=set([re.search('.*(?=\sWard)|.*(?=\sWd)',x).group(0) for x in wards])
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

                for col in cols_dict:
                    candidate=cols_dict[col]
                    if(candidate in candidate_dict):
                        fmt=wb.xf_list[ws.cell_xf_index(row,col)]
                        bold=wb.font_list[fmt.font_index].bold
                        if(bold==1 and candidate!='Scatter'):
                            candidate_dict[candidate]['Winner']=True
                        else:
                            candidate_dict[candidate]['Winner']=False
                for candidate in sorted(candidate_dict):
                    for town in sorted(results_dict[district].keys()):
                        csvwriter.writerow([town,
                                            results_dict[district][town]['county'],
                                            'State Representative',
                                            district,
                                            candidate_dict[candidate]['Party'],
                                            candidate,
                                            candidate_dict[candidate]['Winner'],
                                            results_dict[district][town][candidate]
                                            ])
                candidate_dict=dict()


csvfile.close()