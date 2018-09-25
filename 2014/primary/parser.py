#!/usr/bin/python3.3

# 2014 NH primary election Sept. 9, 2014
# This takes raw .xls results & parses into tabular format

import csv
import xlrd
import re
import os

counties = ['Belknap', 'Carroll', 'Cheshire', 'Coos', 'Grafton', 'Hillsborough', 
	'Merrimack', 'Rockingham', 'Strafford', 'Sullivan']

# counties = ['Merrimack']

state_senate_districts = ['1', '2', '3-4', '5-6', '7-8','9-11', '12-15', '16-18','19-21', '22-24']

offices_with_a_district = [
    'Congressional District 1',
    'Congressional District 2',
    'Executive Council District 1',
    'Executive Council District 2',
    'Executive Council District 3',
    'Executive Council District 4',
    'Executive Council District 5'
    ]

voter_types = ['Democratic', 'Republican']
# voter_types = ['Republican']

results = []

# headers will be (['county', 'town', 'office', 'party', 'candidate', 'votes'])

print ('Working' +
        '\n' +
        '********************')

#### NOTE
#
# *******'PARTY' in these results refers to the voter/ballot type NOT the candidate *******
#
# In New Hampshire, primary voters declare a party, 
# but no matter your party, you get a ballot with ALL the D & R choices on it.
#
# Republcans can vote for Democrats and Democracts can vote for Republicans
#
# For example, assume a town, a precinct and an office:
# Note that Clinton's total is 101, not 100
#
# Dem_ballots.xls
# Clinton(d),Gore(d),Bush(r), Reagan(r)
# 100,150,1,1
#
# R_ballots.xls
# Bush(r), Reagan(r), Clinton(d),Gore(d)
# 200,250,1,1
#
#
# output.csv
# Clinton,D,100
# Clinton,R,1
# Gore,D,100
# Gore,R,1
# Bush,D,1
# Bush,R,200
# Reagan,D,1
# Reagan,R,250
#
#
#
## .csv output preserves 0 or '', whatever is  in the original.
#
######################

#### STATE HOUSE
#
# There are candidates named "True" and "human"

# Some inconsistent layout in state House sheets to correct for.
# Candidate first names are not part of the original data, 
# except in races where two have the same lname.

for county in counties:
    for voter_type in voter_types:
        if (county == 'Coos' and voter_type == 'Democratic'):
            filename = 'House Democratic - Coos.xls'
        else:
            filename = 'House ' + voter_type +'-' + county + '.xls'
        wb = xlrd.open_workbook(filename)
        ws = wb.sheets()[0]

        
        candidates = ws.row_values(2)[1:]
        district = ws.row_values(1)[1].strip().replace('No.7', 'No. 7')

        for row in range(3, ws.nrows):
            town = ''

            # First thing, if it's a completely blank row, skip it
            # some are '','',''
            # some are ' ', ' ', ' '

            if ( 
                ((ws.row_values(row)[0]) in ['', ' ']) and
                ((ws.row_values(row)[1]) in ['', ' '])
                ):
                continue

            # And then check and see if the next row involves a recounting. If it does, that means the current
            # row is some results that will be superceded by the next row; so skip the current row.
             
            if (row < ws.nrows-1):
                if (ws.row_values(row + 1)[0].upper().strip() == 'RECOUNT'):
                    continue   
            
            
            # If it has got a district name in column A, great, harvest the candidate names from that row.
            # Some have blanks like: Smith,Jones,,Scatter
            # The blanks will get stripped out later
            # There's one case where they didn't put the District No. in Column A, they put in the Ward No.
            # Hard-coded solution here

            if ('DISTRICT' in ws.row_values(row)[0].upper() or
                'DISTICT' in ws.row_values(row)[0].upper() or 
                'Dover Ward 15 (1)' in ws.row_values(row)[0]):
                if ('Dover Ward 15 (1)' in ws.row_values(row)[0]):
                    district = "District No. 15"
                else:
                    district = ws.row_values(row)[0].split(' (')[0]
                    district = district.replace('No ', 'No. ').replace('  ', ' ')
                    district = district.replace('No.7', 'No. 7')
                    district = district.replace('Distict', 'District').replace('  ', ' ')
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

            # tho skip totals row and rando 'PSP' row at foot   
        
            if ('TOTAL' in ws.row_values(row)[0].upper() or 
                'PSP' in ws.row_values(row)[0].upper()):
                continue

            # Next, zip candidate names with the rows of results pertaining to them.
            # result[0] will be a single unparsed candidate name & party.

            # except three that share an odd format (town name on row
            # with candidate names) that are getting a hard-coded solution 

            if (county == 'Rockingham' and 
                voter_type == 'Democratic' and
                row == 113):
                candidates = ws.row_values(row)[1:]
                row = row+1

            if (county == 'Rockingham' and 
                voter_type == 'Democratic' and
                row == 36):
                candidates = ws.row_values(row)[1:]
                row = row+1

            if (county == 'Rockingham' and 
                voter_type == 'Democratic' and
                row == 39):
                candidates = ws.row_values(row)[1:]
                row = row+1

            if (county == 'Rockingham' and 
                voter_type == 'Republican' and
                row == 115):
                candidates = ws.row_values(row)[1:]
                row = row+1


            for result in zip(candidates, ws.row_values(row)[1:]):
                
                
                # Skip rows that are made of empty columns like this:
                # district 1,adams,brown,,scatter
                # townville,5,6,,1
                # gotham, 6,5,0,0

                if (result[0] in ['', ' ', '  ']):
                    continue

                if (',' in result[0] or
                    ' ' in result[0].strip()):

                    # names & parties formatted inconsistently, spaces, commas etc

                    # Smith, T., r
                    # Smith, T. r
                    # Smith, Sr.,r
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

                    
                    if ('(w-in)' in result[0] or
                        '(write-in)' in result[0]):
                        cand = (result[0].replace('(write-in)','')
                                            .replace('(w-in)','')
                                            .replace(', d ', '')
                                            .replace(', r ', ''))

                        cand = cand.strip()                        
                        candidate = (cand + ' ' + '(w-in)')
                   
                    else:

                        if (' ' in result[0].strip()):
                            parsed_name = result[0].strip().rsplit(' ',1)
                            candidate = parsed_name[0]

                            

                            if (candidate.endswith(',')):
                                candidate = candidate[:-1]

                            candidate = ' '.join(candidate.split())
                            candidate = candidate.replace('- ', '-') #one really messy name
                        else: 
                            parsed_name = result[0].rsplit(',',1)
                            candidate = parsed_name[0]

                        

                else:
                    # This handles candidate "Scatter" or ppl with no party listed
                    candidate = result[0]
                    party = None

                # Now the second step of correcting for recounts; they're all laid out like: 
                # district 1,smith,jones
                # townville,5,7
                # recount, 6, 7
                # It's already skipping any row where the next row is 'recount',
                # now we just need to grab the right town name.  
            
                if ('RECOUNT' in town.upper()):
                    # grab the town name from the previous spreadsheet row
                    town = ws.row_values(row -1)[0]


                results.append([county, 
                                town.strip().replace('*',''), 
                                "State House " + district, 
                                voter_type.replace('Democratic','D').replace('Republican','R'),
                                str(candidate), # candidate named True
                                result[1]])

# ################# /End state House


####################
# State Senate
#
# first names are not part of the original data
#

for office in state_senate_districts:
  for voter_type in voter_types:
    if '-' in office:
        office_name = 'State Senate Districts '
    else:
        office_name = 'State Senate District '
    filename = office_name + office + ' ' + voter_type + '.xls'
    wb = xlrd.open_workbook(filename)
    ws = wb.sheets()[0]
    
    candidates = ws.row_values(2)[1:]
    district = ws.row_values(1)[1].strip()
    district = district.replace('Republican','').strip()
    district = district.replace('Democratic','').strip()
    district = district.replace('-','').strip()
    for row in range(3, ws.nrows):
        if 'State Senate' in str(ws.row_values(row)[1]):
            district = ws.row_values(row)[1].replace(' - Republican','').strip()
            district = district.replace(' - Democratic','').strip()
            district.replace('Dist.', 'District')
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
                candidate = (result[0])
                candidate = candidate.split(', ')
                candidate = candidate[0]
                candidate = ' '.join(candidate.split())
                candidate = candidate.replace(' (w-in)', '')
            else:
                candidate = result[0]
            this_result = [None, town.strip().replace('*',''), district, voter_type, candidate, result[1]]
            
            # where there are recounts, columns E & F are populated
            # but in the next race where there aren't recounts, ws.row_values(row)[1:] catches blank E & F
            # so don't log those
            if (this_result[4] == ''):
                continue

            results.append([county, 
                town.strip().replace('*',''), 
                district.replace('Dist.', 'District'), 
                voter_type.replace('Republican', 'R').replace('Democratic', 'D'), 
                candidate, 
                result[1]])


# ################# /End State Senate





######### US Senate
# These are laid out differently from the other races

for county in counties:
   for voter_type in voter_types:
      filename = ('USS ' + county + ' County ' + voter_type + '.xls')
      wb = xlrd.open_workbook(filename)
      ws = wb.sheets()[0]

      for row in range(2, ws.nrows):
          # if it's a row that begins with a date, skip it. 
         if (type(ws.row_values(row)[0]) == float):
            continue
         # if the row starts with a blank, skip it. 
         elif ((len(ws.row_values(row)[0]) < 2)):
            continue
        

         elif ('County' in ws.row_values(row)[0]):
            towns = ws.row_values(row)[1:]
            continue


         for result in zip(towns, ws.row_values(row)[1:]):
            candidate = (ws.row_values(row)[0])
            candidate = candidate.split(', ')
            candidate = candidate[0]
            candidate = ' '.join(candidate.split())
            town = result[0].replace('*','').replace('Ward1', 'Ward 1').replace('Haqmpton', 'Hampton')
            town = ' '.join(town.split())
            votes = result[1]

            if (town == '' or
               town == ' ' or
               'TOTAL' in town.upper()):
               continue
            elif ('CORRECTION' in candidate.upper()):
               continue
            else:
               results.append([county,
                  town,'U.S. Senate', 
                  voter_type.replace('Republican', 'R').replace('Democratic', 'D'),
                  candidate,
                  votes])
######### END U.S. Senate



######### Governor
#

for county in counties:
   if (county == 'Belknap'):
      county = 'Summary-Belknap'
   for voter_type in voter_types:
      filename = 'Governor ' + county + ' County ' + voter_type + '.xls'
      wb = xlrd.open_workbook(filename)
      ws = wb.sheets()[0]

      if (county == 'Summary-Belknap'):
         candidates = ws.row_values(16)[1:]
         range_start = 17
      else:
         candidates = ws.row_values(3)[1:]
         range_start = 4

      for row in range(range_start, ws.nrows):
         town = ws.row_values(row)[0]
         if (town == 'Weoodstock'):
            town = 'Woodstock'

         if (town == 'Belknap County' or 
         town == '' or 
         town == ' ' or 
         'CORRECTION' in town.upper() or 
         town == 'TOTALS'):
             continue
         
         for result in zip(candidates, ws.row_values(row)[1:]):
            if result[0] == ' ':
              continue
            else:
               if ',' in result[0]:
                  candidate = result[0].split(', ')
                  candidate = candidate[0]
                  candidate = ' '.join(candidate.split())
                  party = voter_type.replace('Republican', 'R').replace('Democratic', 'D')
               else:
                  candidate = result[0]
                  party = voter_type.replace('Republican', 'R').replace('Democratic', 'D')
              
               results_county = county
              
               if (county == 'Summary-Belknap'):
                  results_county = 'Belnap'

               results.append([results_county, town.strip().replace('*',''), 'Governor', party, candidate, result[1]])
              
# ################# /End Governor

####################
# Congress & Executive Council

# First some cleanup, in case you are doing this from a fresh download that has a typo

if os.path.isfile('Executive Council District 2 Democrattic.xls'):
   os.rename('Executive Council District 2 Democrattic.xls', 'Executive Council District 2 Democratic.xls')


for office in offices_with_a_district:
   for voter_type in voter_types:
      wb = xlrd.open_workbook(office + ' ' + voter_type + '.xls')
      ws = wb.sheets()[0]
      candidates = ws.row_values(2)[1:]

      for row in range(3, ws.nrows):
         town = ws.row_values(row)[0].replace('- ', '')
         town = ' '.join(town.split())
         if ('CORRECTION' in town.upper() or
            'TOTAL' in town.upper() or
            town == ''):
                continue
    
         for result in zip(candidates, ws.row_values(row)[1:]):
            if result[0] == '':
               continue
            else:
               if ',' in result[0]:
                  candidate = result[0].split(', ')
                  candidate = candidate[0]
                  candidate = ' '.join(candidate.split())
                  party = voter_type.replace('Republican', 'R').replace('Democratic', 'D')
                  # party = party.upper()
               else:
                  candidate = result[0]
                  party = voter_type.replace('Republican', 'R').replace('Democratic', 'D')
               results.append([None, town.strip().replace('*',''), office, party, candidate, result[1]])

######## End Congress & Executive Council




csvfile=open('20140909__nh__primary__town.csv','wb')
csvwriter=csv.writer(csvfile)
csvwriter.writerow(['county', 'town', 'office', 'party', 'candidate', 'votes'])
csvwriter.writerows(results)















