#!/usr/bin/python3.3

# 2014 NH general election Nov. 4, 2014
# This takes raw .xls results & parses into tabular format

import csv
import xlrd
import re

counties = ['Belknap', 'Carroll', 'Cheshire', 'Coos', 'Grafton', 'Hillsborough', 
	'Merrimack', 'Rockingham', 'Strafford', 'Sullivan']

offices = ['Gov', 'USS']

house_offices = ['Congressional District 1', 'Congressional District 2']

exec_council_offices = ['1', '2', '3', '4', '5']

state_senate_districts = ['1', '2', '3-4', '5-6', '7-8','9-11', '12-15', '16-18','19-21', '22-24']


results = []

# headers will be (['county', 'town', 'office', 'party', 'candidate', 'votes'])

print ('WELCOME ' +
        '\n' +
        '\n' +
        '\n' +
        '\n' +
        '********************')

####################
# State House

# *LOTS* of inconsistent layout in state House sheets to correct for.
# Candidate first names are not part of the original data, except in races where two have the same lname.
# PS, yes there is a candidate named "True" :)
# parties are: R,D,U,IND,W-IN,D/R,R/D (yes, "u")

for county in counties:
    wb = xlrd.open_workbook('House-'+ county + '.xls')
    ws = wb.sheets()[0]
    town = ''
    for row in range(3, ws.nrows):

        # OK first there's this one uniquely laid-out district/town combo in the whole state
        # we're going to hard-code a solution for.
        
        if (county == 'Rockingham' and
            row == 89):
            town = "Exeter"
            district = "District No. 18"
            candidates = ws.row_values(89)[1:]
            result_row = ws.row_values(90)[1:]
            
            for result in zip(candidates, result_row):
                if (result[0] == '' or
                result[0] == ' '):
                    continue

                if (',' in result[0] or
                ' ' in result[0]):

                    unparsed_info = re.split(' |,', result[0])
                    party = unparsed_info[-1].strip().upper()
                    s = ' '
                    candidate = (s.join(unparsed_info[:-1]))

                else:
                    # This handles candidate "Scatter"
                    candidate = result[0]
                    party = None

                results.append([county, town.strip().replace('*',''), 'State House '+ district, party, candidate, result[1]])
            continue

        if (county == 'Rockingham' and
            row == 90):
            continue
        # /end hard-code item


        # There are a lot of places formatted like this, where the Dist No. is not repeated
        # district 5, aaron, bourne
        # townville, 5,7
        # ,,,
        # ,,clark, davidson
        # townville, 3,5
        #
        # 
        # First thing, if it's a completely blank row, skip it
        # some are '','',''
        # but some are ' ', ' ', ' '  

        elif ((len(ws.row_values(row)[0]) < 2) and (len(ws.row_values(row)[1]) < 2)):
            continue

        # And then check and see if the next row involves a recounting. If it is, skip it this one. 
        # This is step one toward correcting for one kind of recount layout.
         
        if (row < ws.nrows-1):
            if (ws.row_values(row + 1)[0].upper().strip() == 'RECOUNT'):
                continue   

        if (row < ws.nrows-1):
            if (ws.row_values(row + 1)[0].upper().strip() == 'BALLOT LAW COMMISSION'):
                continue 

        # If it has got a district name in column A, great, harvest the candidate names from that row.
        # There's one case where they didn't put the District No. in Column A, they put in the Ward No.
        # Hard-coded solution here

        if ('District' in ws.row_values(row)[0] or
            'Distict' in ws.row_values(row)[0] or
            'Dover Ward 15 (1)' in ws.row_values(row)[0]):
            # this_row = ws.row_values(row)
            # row_without_blanks = list(filter(None, this_row))
            # district = row_without_blanks[0].split(' (')[0]
            # candidates = row_without_blanks[1:]
            if ('Dover Ward 15 (1)' in ws.row_values(row)[0]):
                district = "District No. 15"
            else:    
                district = ws.row_values(row)[0].split(' (')[0]
            candidates = ws.row_values(row)[1:]
            continue

        # But if column A is blank (and we've already gotten rid of the all-blank rows),
        # you know there's another group of candidate names to harvest in columns B through whatever.
        # District will remain unchanged from previous pass.

        elif (len(ws.row_values(row)[0]) < 2):
            candidates = ws.row_values(row)[1:]
            continue

        # Then almost the last thing ur left with is a perfectly normal row with a town name and vote counts

        else:
            town = ws.row_values(row)[0]
        
        # If it's a footnote about corrections or the 'totals' row, skip it
        
        if ('CORRECTION' in town.upper() or
            'TOTALS' in town.upper()):
            continue

        # Now zip candidate names with the row of results pertaining to them.
        # result[0] will be a single unparsed candidate name & party. 

        for result in zip(candidates, ws.row_values(row)[1:]):

            # Skip rows that are made of empty columns like this:
            # district 1,adams,brown,,scatter
            # townville,5,6,,1
            # gotham, 6,5,0,0

            if (result[0] == '' or
                result[0] == ' '):
                continue

            if (',' in result[0] or
                ' ' in result[0].strip()):
                party = "test"

                # names & parties formatted inconsistently, spaces, commas etc

                # Smith, T., r
                # Smith, T. r
                # Smith, Sr.,r
                # Smith, ind
                # Smith ,d
                # Brown
                # Brown,d 
                # Brown, d
                # Brown,  d                   
                # Smith, T. r
                # Smith,M,r
                # Smith,R,r
                # Smith-Jones, d
                # Smith -Jones, d
                # Smith, d (w-in)
                # Smith, Sr. (w-in)
                # Smith (w-in)
                # Smith (write-in)
                # Jane Smith (w-in)

                
                if ('(w-in)' in result[0]):
                    cand = (result[0].replace('(write-in)','')
                                        .replace('(w-in)','')
                                        .replace(', d ', '')
                                        .replace(', r ', ''))

                    cand = cand.strip()                        
                    candidate = (cand + ' ' + '(w-in)')
                    party = ''
               
                else:

                    if (' ' in result[0].strip()):
                        parsed_name = result[0].strip().rsplit(' ',1)
                        candidate = parsed_name[0]
                        party = parsed_name[1].upper().replace(',','')
                        

                        if (candidate.endswith(',')):
                            candidate = candidate[:-1]

                        candidate = ' '.join(candidate.split())
                        candidate = candidate.replace('- ', '-') #one really messy name

                    else: 
                        # cases Smith,d
                        parsed_name = result[0].rsplit(',',1)
                        candidate = parsed_name[0]
                        party = parsed_name[1]


            else:
                # This handles candidate "Scatter"
                candidate = result[0]
                party = None
            
            
            # And correct for recounts.
            # Recounts are laid out in two different ways sometimes on the same sheet.

            # This takes care of this kind: 
            # district 1, smith,recount,jones,recount
            # townville,6,6,7,8
            # gotham city,5,5,7,8

            this_result = [county, town.strip().replace('*',''), "State House "+ district, party, candidate, result[1]]

            if ('RECOUNT' in this_result[4].upper()):
                # so grab the name & party from the previous record
                candidate = results[-1][4]
                party = results[-1][3]
                # and delete the previous record
                del results[-1] 

            # But there's the other layout for recounts.    
            # And this is the second step of taking care of this kind:
            # district 1,smith,jones
            # townville,5,7
            # recount, 6, 7
            # It's already skipping any row where the next row is 'recount',
            # now we just need to grab the right town name.  
            
            if ('RECOUNT' in town.upper()):
                # grab the town name from the previous spreadsheet row
                town = ws.row_values(row -1)[0]

            # if the ballot law commission was involved, a recount will have come first
            # delete the 'recount' results

            if ('BALLOT LAW COMMISSION' in town.upper()):
                town = ws.row_values(row -2)[0].replace(' (tie)', '')

            results.append([county, town.strip().replace('*',''), "State House " + district.replace('Distict', 'District'), party, candidate, result[1]])

################# /End State House

####################
# State Senate
#

for office in state_senate_districts:
    if '-' in office:
        office_name = 'State Senate Districts '
    else:
        office_name = 'State Senate District '
    wb = xlrd.open_workbook( office_name + office+'.xls')
    ws = wb.sheets()[0]
    candidates = ws.row_values(2)[1:]
    district = ws.row_values(1)[1].strip()
    for row in range(3, ws.nrows):
        if 'State Senate' in str(ws.row_values(row)[1]):
            district = ws.row_values(row)[1].replace(' - Republican','').strip()
            candidates = ws.row_values(row + 1)[1:]
        town = ws.row_values(row)[0]

        if (len(town) < 2 or  # if blank or looks blank
            town.upper() == 'TOTALS' or  # if it's the totals row
            'correction' in town): # or a footnote about corrections
            continue

        for result in zip(candidates, ws.row_values(row)[1:]):
            # Skip rows that are made of empty columns like this:
            # district 1,adams,brown,,scatter
            # townville,5,6,,1
            # gotham, 6,5,0,0

            if (result[0] == '' or
                result[0] == ' '):
                continue

            if ',' in result[0]:
                candidate, party = result[0].split(',')
                party = party.strip().upper()
                candidate = ' '.join(candidate.split())
            else:
                candidate = result[0]
                party = None
            this_result = [None, town.strip().replace('*',''), district, party, candidate, result[1]]
            
            # where there are recounts, columns E & F are populated
            # but in the next race where there aren't recounts, ws.row_values(row)[1:] catches blank E & F
            # so don't log those
            if (this_result[4] == ''):
                continue

            # however there are some recounts
            if 'RECOUNT' in this_result[4].upper():
                # so grab the name & party from the previous record
                candidate = results[-1][4]
                party = results[-1][3]
                # and delete the previous record
                del results[-1] 
            results.append([None, town.strip().replace('*',''), district.replace('Dist.', 'District'), party, candidate, result[1]])
# ################# /End State Senate


####################
# Executive Council
for office in exec_council_offices:
    wb = xlrd.open_workbook("Executive Council District "+ office + '.xls')
    ws = wb.sheets()[0]
    candidates = ws.row_values(2)[1:]
    for row in range(3, ws.nrows):
        town = ws.row_values(row)[0]
        if town == '*corrections received' or town == '' or town == "Totals":
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
                results.append([None, town.strip().replace('*',''), 
                    "Executive Council District " + office, party, candidate, result[1]])
################# /End Executive Council



####################
# Congress

for office in house_offices:
    wb = xlrd.open_workbook(office + '.xls')
    ws = wb.sheets()[0]
    candidates = ws.row_values(2)[1:]
    for row in range(3, ws.nrows):
        town = ws.row_values(row)[0]
        if town == 'Totals' or town == '' or town == '*corrections received':
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

################# /End Congress



####################
# Governor & U.S. Senate
# 
# Zeros and blanks are retained as zeroes and blanks.
# In the .xls, Coos County shows '0' as '-' due to Excel formatting.
# This renders them as 0.

for office in offices:
    name_segment = ' 2014.xls'
    for county in counties:
        if county == 'Belknap':
            # for opening the .xls
            county_name = 'Summary - Belknap'
        else:
            county_name = county
        wb = xlrd.open_workbook(office + ' ' + county_name + ' County' + name_segment, formatting_info=True)
        ws = wb.sheets()[0]
        candidates = ws.row_values(3)[1:]

        #this corrects for Belknap county being in the same .xls w summary
        if county_name == 'Summary - Belknap':
            range_start = 18
        else:
            range_start = 4

        for row in range(range_start, ws.nrows):

            town = ws.row_values(row)[0]

            if (town == 'Belknap County' or 
            town == '' or 
            town == ' ' or 
            'correction' in town or 
            town == 'TOTALS'):
                continue

            # One sheet has some extraneous characters in column D    
            for result in zip(candidates, ws.row_values(row)[1:4]):
                

                if result[0] == ' ':
                    continue
                else:
                    if ',' in result[0]:
                        candidate, party = result[0].split(', ')
                        party = party.upper()
                    else:
                        candidate = result[0]
                        party = None
                    results_county = county
                    # clean up stuff
                    
                    if candidate == 'Hassan':
                        candidate = 'Maggie Hassan'
                    if candidate == 'Havenstein':
                        candidate = 'Walt Havenstein'
                    if candidate == 'Brown':
                        candidate = 'Scott P. Brown'
                    if candidate == 'Shaheen':
                        candidate = 'Jeanne Shaheen'
                    if office == 'USS':
                        office_name = 'U.S. Senate'
                    if office == 'Gov':
                        office_name = 'Governor'

                    results.append([results_county, town.strip().replace('*',''), office_name, party, candidate, result[1]])

################# /End Governor & U.S. Senate




csvfile=open('20141104__nh__general__town.csv','wb')
csvwriter=csv.writer(csvfile)
csvwriter.writerow(['county', 'town', 'office', 'party', 'candidate', 'votes'])
csvwriter.writerows(results)









