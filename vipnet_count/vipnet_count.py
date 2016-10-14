#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Created on 13 окт. 2016 г.

@author: v_ilyin
'''
import json
import os
import datetime

from xlwt import Workbook, easyxf

def vipnet_json2xls(data):
    """Write data to Excel spreadsheet."""

    header_style = easyxf(
        'font: name Arial, bold on, height 160, colour black;'
        'borders: left thin, right thin, top thin, bottom thin;'
        'alignment: horizontal centre, vertical centre;'
        'pattern: pattern solid, fore_colour gray25;'
        )
    cell_style = easyxf(
        'font: name Arial, bold off, height 160;'
        #'borders: left medium, right medium, top medium, bottom medium;'
        #'alignment: horizontal centre, vertical centre;'
        )

    book = Workbook()
    sheet = book.add_sheet('Vipnet', cell_overwrite_ok=True)

    columns = ("name","type","product-version","drv-version","monitor-version","ifaces")
    #print(columns)
    for i, row in enumerate(data):
        #print(row)
        for j, col in enumerate(columns):
            if i == 0:
                sheet.write(i, j, col, header_style)
            if col not in ['ifaces']:
                try:
                    sheet.write(i+1, j, row[col], cell_style)
                except KeyError:
                    sheet.write(i+1, j, '', cell_style)
            elif col in ["ifaces"]:
                sheet.write(i+1, j, len(row[col]), cell_style)
    #rowx = 1
    #for device in data:
    #    sheet.write(rowx, colx, value, header_style)

    current_datetime = datetime.datetime.now().strftime("%Y-%m-%d--%H-%M")
    
    if not os.path.exists(os.path.join(os.getcwd(),"xls\\")):
        os.makedirs(os.path.join(os.getcwd(),"xls\\"))

    try:
        book.save(os.path.join(os.getcwd(), "xls\\" + current_datetime + '.xls'))
    except Exception as e:
        print("Error: %s" % str(e))
    return

def main():
    print("load...")
    with open('data-01.json') as data_file:    
        data = json.load(data_file)

    print("process..")
    vipnet = []
    vipnet_count = 0
    vipnet_count_G1 = 0
    vipnet_count_G2 = 0
    repl = 0
    i = 0
    for dev in data["export"]["rt"]:
        i = i+1
        print(i)
        #print(json.dumps(dev, indent=4, sort_keys=True))
        
        #if os.path.exists("text1.json"):
        #    os.remove("text1.json")
        #f = open("text1.json", "w")
        #f.write(json.dumps(dev, indent=4, sort_keys=True))
        #f.close()

        
        #break
        ifaces = []
        if "ifaces" in dev:
            for ff in dev["ifaces"]["iface"]:
                new_ff = {
                          "name": ff["iface-name"]["#text"],
                          "ip": ff["iface-ip"]["#text"],
                          "netmask": ff["iface-netmask"]["#text"]
                          }
                ifaces.append(new_ff)
        version = ''
        if "product-version" in dev:
            version = dev["product-version"]["#text"]
        drv_ver = ''
        if "drv-version" in dev:
            drv_ver = dev["drv-version"]["#text"]
        mon_ver = ''
        if "monitor-version" in dev:
            mon_ver = dev["monitor-version"]["#text"]
        new = {
                "name": dev["node-name"]["#text"],
                "type": dev["node-type"]["#text"],
                "ifaces": ifaces,
                "product-version": version,
                "drv-version": drv_ver,
                "monitor-version": mon_ver
            }

        node_exist = False
        for j,n in enumerate(vipnet):
            if new["name"] in n["name"]:
                if ("COORDINATOR" in n["type"]) and ("COORDINATOR" not in new["type"]):
                    del vipnet[j]
                    repl = repl+1
                    vipnet_count = vipnet_count-1
                    break
                else:
                    node_exist = True
        if not node_exist:
            vipnet_count = vipnet_count+1
            
            if len(ifaces) == 2:
                vipnet_count_G1 = vipnet_count_G1+1
            if len(ifaces) == 4:
                vipnet_count_G2 = vipnet_count_G2+1
            
            vipnet.append(new)


    print(json.dumps(vipnet, indent=4, sort_keys=True))
    print("обработано записей: ", i)
    print("перезаписано записей: ", repl)
    print("vipnet count G1 + G2 + недоступны: ", vipnet_count)
    print("vipnet G1 count: ", vipnet_count_G1)
    print("vipnet G2 count: ", vipnet_count_G2)
    if os.path.exists("vipnet-01.json"):
        os.remove("vipnet-01.json")
    f = open("vipnet-01.json", "w")
    f.write(json.dumps(vipnet, indent=4, sort_keys=True))
    f.close()

    vipnet_json2xls(vipnet)

    return

if __name__ == '__main__':
    main()