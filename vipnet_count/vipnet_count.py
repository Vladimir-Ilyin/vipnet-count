#!/usr/bin/env python3
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
import xmltodict

from xlwt import Workbook, easyxf

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
    for i, row in enumerate(data):
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
        book.save(out_file)
    except Exception as e:
        print("Error: %s" % str(e))
    return

def datetime_parser(dct):
    for k, v in dct.items():
        if k in "timestamp":
            try:
                dct[k] = datetime.datetime.strptime(v, "%Y-%m-%dT%H:%M:%S")
            except:
                pass
    return dct

def main():
    current_datetime = datetime.datetime.now().strftime("%Y-%m-%d--%H-%M")

    datehandler = lambda obj: (
                               obj.isoformat()
                               if isinstance(obj, datetime.datetime)
                               or isinstance(obj, datetime.date)
                               else None
                               )

    p = optparse.OptionParser(
        description='Converts XML data from StateWatcher to XLS and JSON.  Reads from file.',
        prog='vipnet_count',
        usage='%prog -o OUT file.xml'
    )
    p.add_option('--out', '-o', help="Write to OUT (OUT.xls and OUT.json)", default=current_datetime)
    p.add_option('--input-json', '-i', help="Input json file with VipNet data")
    options, arguments = p.parse_args()

    if len(arguments) == 1:
        try:
            print("load...")
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
        
    print("convert...")
    data = xmltodict.parse(xml_data)
    xml_data = ""

    print("process...")
    vipnet_count = 0
    vipnet_count_G1 = 0
    vipnet_count_G2 = 0
    if (options.input_json):
        try:
            vipnet = json.load(open(options.input_json), object_hook=datetime_parser)
            for n in vipnet:
                vipnet_count = vipnet_count+1
                if "ifaces" in n:
                    if len(n["ifaces"]) == 2:
                        vipnet_count_G1 = vipnet_count_G1+1
                    if len(n["ifaces"]) == 4:
                        vipnet_count_G2 = vipnet_count_G2+1
        except:
            sys.stderr.write("Problem reading '{0}'\n".format(options.input_json))
            p.print_help()
            sys.exit(-1)
    else:
        vipnet = []
    repl = 0
    repl_renew = 0
    i = 0
    print("обработано записей: %d\r"%i, end="")
    for dev in data["export"]["rt"]:
        i = i+1
        print("обработано записей: %d\r"%i, end="")
        
        ifaces = []
        if "ifaces" in dev:
            for ff in dev["ifaces"]["iface"]:
                new_ff = {
                          "name": ff["iface-name"],
                          "ip": ff["iface-ip"],
                          "netmask": ff["iface-netmask"]
                          }
                ifaces.append(new_ff)
        version = ''
        if "product-version" in dev:
            version = dev["product-version"]
        drv_ver = ''
        if "drv-version" in dev:
            drv_ver = dev["drv-version"]
        mon_ver = ''
        if "monitor-version" in dev:
            mon_ver = dev["monitor-version"]
        
        timestamp = ''
        if "poll-timestamp" in dev:
            timestamp = datetime.datetime.strptime( dev["poll-timestamp"], "%Y-%m-%d %H:%M:%S")
        ip_list = ''
        if "ip-list" in dev:
            ip_list = dev["ip-list"]

        new = {
                "name": dev["node-name"],
                "type": dev["node-type"],
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
                elif ("COORDINATOR" not in n["type"]) and ("COORDINATOR" not in new["type"]) and (n["timestamp"] < new["timestamp"]):
                    if len(n["ifaces"]) == 2:
                        vipnet_count_G1 = vipnet_count_G1-1
                    if len(n["ifaces"]) == 4:
                        vipnet_count_G2 = vipnet_count_G2-1
                    del vipnet[j]
                    repl_renew = repl_renew+1
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

    print("обработано записей: ", i)
    print("перезаписано записей с новым типом: ", repl)
    print("перезаписано записей с новой датой: ", repl_renew)
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