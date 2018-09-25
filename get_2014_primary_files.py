#!/usr/bin/python3.3

# This downloads results from the 2014 NH General election
# Puts each .xls in a same-level directory called "2014/primary/"

import requests
from bs4 import BeautifulSoup
import time
import os

curr_path = os.path.dirname(os.path.abspath(__file__))
curr_path = curr_path + "/2014/primary/"

if not os.path.exists(curr_path):
    os.makedirs(curr_path)

url_stem = "http://sos.nh.gov"
url_list = [
    "http://sos.nh.gov/2014GovDemPrim.aspx",
    "http://sos.nh.gov/2014USSDemPrim.aspx",
    "http://sos.nh.gov/2014ConDemPrim.aspx",
    "http://sos.nh.gov/2014ExeDemPrim.aspx",
    "http://sos.nh.gov/2014SenDemPrim.aspx",
    "http://sos.nh.gov/2014RepDemPrim.aspx",
    "http://sos.nh.gov/2014GovRepPrim.aspx",
    "http://sos.nh.gov/2014USSRepPrim.aspx",
    "http://sos.nh.gov/2014ConRepPrim.aspx",
    "http://sos.nh.gov/2014ExeRepPrim.aspx",
    "http://sos.nh.gov/2014SenRepPrim.aspx",
    "http://sos.nh.gov/2014RepRepPrim.aspx"
    ]

for j in range(len(url_list)):
    time.sleep(2)
    response = requests.get(url_list[j])
    html = response.content
    soup = BeautifulSoup(html)

    excels = soup.find("div", attrs={"id" : "ctl00_cphMain_dzCenterColumn_columnDisplay_ctl00_zone"})

    dl_links = excels.findAll("a")
    for i in range(len(dl_links)):
        time.sleep(5)
        single_excel = url_stem + dl_links[i]['href']
        title =  curr_path + dl_links[i].get_text() + ".xls"
        resp = requests.get(single_excel)
        output = open(title, 'wb')
        output.write(resp.content)
        output.close()
