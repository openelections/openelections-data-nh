import csv
import xlrd
import requests

counties = ['Carroll', 'Cheshire', 'Coos', 'Grafton', 'Hillsborough', 'Merrimack', 'Rockingham', 'Strafford Sullivan', 'Belknap']
senate_counties = ['Belknap', 'Carroll', 'Cheshire', 'Coos', 'Grafton', 'Hillsborough', 'Merrimack', 'Rockingham', 'Strafford', 'Sullivan']
offices = ['President', 'Governor', 'USS']
house_offices = ['Congressional District 1', 'Congressional District 2']#, 'Executive Council', 'State Senate', 'House']
exec_council_offices = ['1', '2', '3', '4', '5']
state_senate_districts = ['1', '2', '3-4', '5-6', '7-8','9-11', '12-15', '16-18','19-21', '22-24']
results = []

for office in offices:
    name_segment = ' 2016-excel.xls'
    if office == 'USS':
        counties = senate_counties
    for county in counties:
        if county == 'Belknap':
            county_name = 'Sum and Belknap'
        else:
            county_name = county
        # handle council & senate files by district
        wb = xlrd.open_workbook(office+' '+county_name+name_segment)
        ws = wb.sheets()[0]
        candidates = ws.row_values(3)[1:]
        for row in range(4, ws.nrows):
            town = ws.row_values(row)[0]
            if town == '*This is not the total votes cast for President. Votes for write-in candidates will be added later.' or town == '*correction submitted by clerk' or town == 'Sullivan County' or town == '':
                continue
            for result in zip(candidates, ws.row_values(row)[1:]):
                if result[0] == ' ':
                    continue
                else:
                    if ',' in result[0]:
                        candidate, party = result[0].split(', ')
                        party = party.upper()
                    else:
                        candidate = result[0]
                        party = None
                    if county == 'Strafford Sullivan' and town == 'Barrington':
                        results_county = 'Strafford'
                    elif county == 'Strafford' and town == 'Acworth':
                        results_county = 'Sullivan'
                    else:
                        results_county = county
                    results.append([results_county, town.strip().replace('*',''), office, party, candidate, result[1]])

# presidential write-ins

county = None
write_in_counties = ['Belknap', 'Carroll', 'Cheshire', 'Coos', 'Grafton', 'Hillsboro', 'Merrimack', 'Rockingham', 'Strafford', 'Sullivan']

for office in ['President']:
    wb = xlrd.open_workbook(office+'-write-ins.xls')
    for county in write_in_counties:
        ws = wb.sheet_by_name(county.upper())
        candidates = ws.row_values(2)[1:]
        for row in range(3, ws.nrows):
            start_col = 0
            if county == 'Sullivan':
                start_col = 1
            town = ws.row_values(row)[start_col]
            for result in zip(candidates, ws.row_values(row)[start_col+1:]):
                if result[1] == '':
                    continue
                else:
                    candidate = result[0]
                    party = None
                    if county == 'Hillsboro':
                        county = 'Hillsborough'
                    results.append([county, town.strip(), office, party, candidate, result[1]])

# Congressional
for office in house_offices:
    wb = xlrd.open_workbook(office+'-excel.xlsx')
    ws = wb.sheets()[0]
    candidates = ws.row_values(2)[1:]
    for row in range(3, ws.nrows):
        town = ws.row_values(row)[0]
        if town == '*This is not the total votes cast for President. Votes for write-in candidates will be added later.' or town == '*correction submitted by clerk' or town == 'Sullivan County' or town == '':
            continue
        for result in zip(candidates, ws.row_values(row)[1:]):
            if result[0] == '':
                continue
            else:
                if ',' in result[0]:
                    candidate, party = result[0].split(', ')
                    party = party.upper()
                else:
                    candidate = result[0]
                    party = None
                results.append([None, town.strip().replace('*',''), office, party, candidate, result[1]])

# Executive Council
for office in exec_council_offices:
    wb = xlrd.open_workbook("Executive Council District "+office+'-excel.xls')
    ws = wb.sheets()[0]
    candidates = ws.row_values(2)[1:]
    for row in range(3, ws.nrows):
        town = ws.row_values(row)[0]
        if town == '*This is not the total votes cast for President. Votes for write-in candidates will be added later.' or town == '*correction submitted by clerk' or town == 'Sullivan County' or town == '':
            continue
        for result in zip(candidates, ws.row_values(row)[1:]):
            if result[0] == '':
                continue
            else:
                if ',' in result[0]:
                    candidate, party = result[0].split(', ')
                    party = party.upper()
                else:
                    candidate = result[0]
                    party = None
                results.append([None, town.strip().replace('*',''), "Executive Council District "+office, party, candidate, result[1]])

# State Senate
for office in state_senate_districts:
    if '-' in office:
        office_name = 'State Senate Districts '
    else:
        office_name = 'State Senate District '
    wb = xlrd.open_workbook(office_name+office+'-excel.xls')
    ws = wb.sheets()[0]
    candidates = ws.row_values(2)[1:]
    district = ws.row_values(1)[1].strip()
    for row in range(3, ws.nrows):
        if 'State Senate' in str(ws.row_values(row)[1]):
            district = ws.row_values(row)[1].replace(' - Republican','').strip()
            candidates = ws.row_values(row+1)[1:]
        town = ws.row_values(row)[0]
        if town == '*This is not the total votes cast for President. Votes for write-in candidates will be added later.' or town == '*correction submitted by clerk' or town == 'Sullivan County' or town == '':
            continue
        for result in zip(candidates, ws.row_values(row)[1:]):
            if ' --' in str(result[1]) or str(result[1]) == '':
                continue
            else:
                if ',' in result[0]:
                    candidate, party = result[0].split(',')
                    party = party.strip().upper()
                else:
                    candidate = result[0]
                    party = None
                results.append([None, town.strip().replace('*',''), district, party, candidate, result[1]])

# State House
for county in senate_counties:
    if county == 'Coos':
        wb = xlrd.open_workbook('House-'+county+'-2016-excel.xlsx')
    else:
        wb = xlrd.open_workbook('House-'+county+'-2016-excel.xls')
    ws = wb.sheets()[0]
    for row in range(3, ws.nrows):
        if 'District' in str(ws.row_values(row)[0]):
            district = ws.row_values(row)[0].split(' (')[0]
            candidates = ws.row_values(row)[1:]
        else:
            town = ws.row_values(row)[0]
        if town == '*This is not the total votes cast for President. Votes for write-in candidates will be added later.' or town == '*correction submitted by clerk' or town == 'Sullivan County' or town == '':
            continue
        for result in zip(candidates, ws.row_values(row)[1:]):
            if ' --' in str(result[1]) or str(result[1]) == '':
                continue
            else:
                if ',' in result[0]:
                    candidate, party = result[0].replace('Moffett, H. d', 'Moffett H, d').replace('Moffett,', 'Moffett').replace('Ober,', 'Ober').replace(", Sr.", "Sr.").replace(", Jr.", "Jr.").replace(", III", "III").split(',', 2)
                    party = party.strip().upper()
                else:
                    candidate = result[0]
                    party = None
                results.append([county, town.strip().replace('*',''), "State House "+ district, party, candidate, result[1]])

csvfile=open('20161108__nh__general__town.csv','wb')
csvwriter=csv.writer(csvfile)
csvwriter.writerow(['county', 'town', 'office', 'district', 'party', 'candidate', 'votes'])
csvwriter.writerows(results)
