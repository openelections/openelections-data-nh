# -*- coding: utf-8 -*-
"""
Created on Sat Oct  1 17:28:03 2016

@author: mike
"""

rep_dir='../'
csvfiles=['20121106__nh__general__town__pres.csv',
          '20121106__nh__general__town__cd1.csv',
          '20121106__nh__general__town__cd2.csv',
          '20121106__nh__general__town__gov.csv',
          '20121106__nh__general__town__exec_council.csv',
          '20121106__nh__general__town__state__rep.csv',
          '20121106__nh__general__town__state__sen.csv',
          ]

outfile=rep_dir+'20121106__nh__general__town.csv'
o=open(outfile,'wb')

for k,f_name in enumerate(csvfiles):
    f=open(rep_dir+f_name,'rb')
    if(k==0):
        start_line=0
    else:
        start_line=1
    for l,line in enumerate(f):
        if(l>=start_line):
            o.write(line)
    f.close()

o.close()