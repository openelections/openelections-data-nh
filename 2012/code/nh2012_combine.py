# -*- coding: utf-8 -*-
"""
Created on Sat Oct  1 17:28:03 2016

@author: mike
"""


rep_dir='../'
csvfiles=['20121106__nh__general__president__town.csv',
          '20121106__nh__general__house__1__town.csv',
          '20121106__nh__general__house__2__town.csv',
          '20121106__nh__general__governor__town.csv',
          '20121106__nh__general__executive__council__town.csv',
          '20121106__nh__general__state__house__town.csv',
          '20121106__nh__general__state__senate__town.csv',
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