#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Created on 13 окт. 2016 г.

@author: v_ilyin
'''
import json
import optparse
import os
import sys
import datetime

from xlwt import Workbook, easyxf, Formula
from xml2json import xml2json

def vipnet_json2xls(data, out_file):
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

    columns = ("name","type","product-version","drv-version","monitor-version","ifaces","ip-list","timestamp")
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

    #current_datetime = datetime.datetime.now().strftime("%Y-%m-%d--%H-%M")
    
    #if not os.path.exists(os.path.join(os.getcwd(),"xls\\")):
    #    os.makedirs(os.path.join(os.getcwd(),"xls\\"))

    sheet.write(0,9,"G1",header_style)
    sheet.write(1,9,"=СЧЁТЕСЛИ($F:$F;2)",cell_style)
    sheet.write(0,10,"G2",header_style)
    sheet.write(1,10,"=СЧЁТЕСЛИ($F:$F;4)",cell_style)
    sheet.write(0,11,"Недоступно",header_style)
    sheet.write(1,11,"=СЧЁТЕСЛИ($F:$F;0)",cell_style)
    sheet.write(0,12,"Всего",header_style)
    sheet.write(1,12,"=СЧЁТ($F:$F)",cell_style)

    if os.path.exists(out_file):
        os.remove(out_file)

    try:
        #book.save(os.path.join(os.getcwd(), "xls\\" + current_datetime + '.xls'))
        book.save(out_file)
    except Exception as e:
        print("Error: %s" % str(e))
    return

def main():
    current_datetime = datetime.datetime.now().strftime("%Y-%m-%d--%H-%M")
    
    p = optparse.OptionParser(
        description='Converts XML data from StateWatcher to XLS and JSON.  Reads from file.',
        prog='vipnet_count',
        usage='%prog -o OUT file.xml'
    )
    p.add_option('--out', '-o', help="Write to OUT (OUT.xls and OUT.json)", default=current_datetime)
    #p.add_option('--in', '-i', help="Input json file with VipNet data")
    p.add_option(
        '--strip_text', action="store_true",
        dest="strip_text", help="Strip text for xml2json")
    p.add_option(
        '--pretty', action="store_true",
        dest="pretty", help="Format JSON output so it is easier to read")
    p.add_option(
        '--strip_namespace', action="store_true",
        dest="strip_ns", help="Strip namespace for xml2json")
    p.add_option(
        '--strip_newlines', action="store_true",
        dest="strip_nl", help="Strip newlines for xml2json")
    options, arguments = p.parse_args()

    print("load...")
    if len(arguments) == 1:
        try:
            xml_file = open(arguments[0])
        except:
            sys.stderr.write("Problem reading '{0}'\n".format(arguments[0]))
            p.print_help()
            sys.exit(-1)
    else:
        p.print_help()
        sys.exit(-1)

    xml_data = xml_file.read()
    xml_file.close()
        
    strip = 0
    strip_ns = 0
    if options.strip_text:
        strip = 1
    if options.strip_ns:
        strip_ns = 1
    if options.strip_nl:
        xml_data = xml_data.replace('\n', '').replace('\r','')
    
    print("convert...")
    data = xml2json(xml_data, options, strip_ns, strip)
    #data = {"export": {"rt":[]}}
    xml_data = ""
    
    #with open('data-01.json') as data_file:    
    #    data = json.load(data_file)

    print("process..")
    #if (options.in):
    #    try:
    #        xml_file = open(arguments[0])
    #    except:
    #        sys.stderr.write("Problem reading '{0}'\n".format(arguments[0]))
    #        p.print_help()
    #        sys.exit(-1)
    #else:
    vipnet = []
    vipnet_count = 0
    vipnet_count_G1 = 0
    vipnet_count_G2 = 0
    repl = 0
    i = 0
    #print("обработано записей: ", i, end="")
    for dev in data["export"]["rt"]:
        i = i+1
        #print("\b"*len(str(i-1)), i, end="")
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
        
        timestamp = ''
        if "poll-timestamp" in dev:
            timestamp = datetime.datetime.strptime( dev["poll-timestamp"]["#text"], "%Y-%m-%d %H:%M:%S")
        ip_list = ''
        if "ip-list" in dev:
            ip_list = dev["ip-list"]["#text"]

        new = {
                "name": dev["node-name"]["#text"],
                "type": dev["node-type"]["#text"],
                "ip-list": ip_list,
                "ifaces": ifaces,
                "product-version": version,
                "drv-version": drv_ver,
                "monitor-version": mon_ver,
                "timestamp": timestamp
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

    datehandler = lambda obj: (
                               obj.isoformat()
                               if isinstance(obj, datetime.datetime)
                               or isinstance(obj, datetime.date)
                               else None
                               )

    #print(json.dumps(vipnet, indent=4, sort_keys=True, default=datehandler))
    print("обработано записей: ", i)
    print("перезаписано записей: ", repl)
    print("vipnet count G1 + G2 + недоступны: ", vipnet_count)
    print("vipnet G1 count: ", vipnet_count_G1)
    print("vipnet G2 count: ", vipnet_count_G2)

    if (options.out):
        out_file_json = options.out + ".json"
        out_file_xls = options.out + ".xls"

    if os.path.exists(out_file_json):
        os.remove(out_file_json)
    f = open(out_file_json, "w")
    f.write(json.dumps(vipnet, indent=4, sort_keys=True, default=datehandler))
    f.close()

    vipnet_json2xls(vipnet, out_file_xls)

    return

if __name__ == '__main__':
    main()