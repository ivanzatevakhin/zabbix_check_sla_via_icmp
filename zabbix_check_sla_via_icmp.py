#!/usr/bin/env python3
import json
import sys
import os
import logging
import logging.handlers
import requests
import xmltodict
import xlsxwriter

ZABBIX_API_URL = "http://127.0.0.1/zabbix/api_jsonrpc.php"
UNAME = "APITEST"
PWORD = "testpassword"

r = requests.post(ZABBIX_API_URL,
                  json={
                        "jsonrpc": "2.0",
                        "method": "user.login",
                        "params": {
                        "user": UNAME,
                        "password": PWORD},
                        "id": 1
                  })

#print(json.dumps(r.json(), indent=4, sort_keys=True))

AUTHTOKEN = r.json()["result"]

host_group = input("Ведите имя хостгруппы в которой будем осуществлять поиск: ")
time_from = input("Ведите время начала подсчета: ")
time_till = input("Ведите время окончания подсчета: ")

max_length_A = 0
max_length_B = 0
max_length_C = 0
max_length_D = 0
max_length_E = 0
max_length_F = 0
max_length_G = 0

xlsx_dir = "excel"
excelfile = '%s/items.xlsx' % xlsx_dir
workbook = xlsxwriter.Workbook(excelfile)                           # Создает новый excel файл с именем items.xlsx
bold = workbook.add_format(
        {
            'bold': True
            }
        )
item_header = workbook.add_format(
        {
            'bg_color': '#CCCCCC',
            'bold': True,
            'border': 1
            }
        )
item_text = workbook.add_format(
        {
            'border': 1
            }
        )

index = 3

try:
    worksheet = workbook.add_worksheet("Zabbixgroup")                   # Создает файл
except Exception as ex:
    print("Error while adding worksheet")

worksheet.write('A1', 'Хост id', item_header)
worksheet.write('B1', 'Метрика id', item_header)
worksheet.write('C1', 'Имя хоста', item_header)
worksheet.write('D1', 'Среднее значение', item_header)
worksheet.write('E1', 'Инвентарные данные 1', item_header)
worksheet.write('F1', 'Инвентарные данные 2', item_header)
worksheet.write('G1', 'Инвентарные данные 3', item_header)

max_length_A = max(max_length_A, len('Хост id'))
max_length_B = max(max_length_B, len('Метрика id'))
max_length_C = max(max_length_C, len('Имя хоста'))
max_length_D = max(max_length_D, len('Среднее значение'))
max_length_E = max(max_length_E, len('Инвентарные данные 1'))
max_length_F = max(max_length_E, len('Инвентарные данные 2'))
max_length_G = max(max_length_E, len('Инвентарные данные 3'))

def get_hostname(hostid):
    j = requests.post(ZABBIX_API_URL,

                json={
                    "jsonrpc": "2.0",
                    "method": "host.get",
                    "params": {
                        "output": ["host"],
                        "filter": {
                            "hostid" : hostid
                        }
                    },
                    "id": 2,
                    "auth": AUTHTOKEN
                })
    data_j = json.dumps(j.json(), indent=4, sort_keys=True)
    data_j_1 = json.loads(data_j)
    dat = (data_j_1['result'])
    if dat == []:
        data_j_j_1 = "Это шаблон, либо хост в статусе Disabled"
    else:
        data_j_j_1 = (data_j_1['result'][0]['host'])
    #data_j_1_1 = (data_j_1)['result']['name']
    #(data_j_1['result'][0]['hostid'])
    return data_j_j_1


def get_history(items,time_from,time_till):
    spisok = []
    q = requests.post(ZABBIX_API_URL,

                json={
                    "jsonrpc": "2.0",
                    "method": "trend.get",
                    "params": {
                        "output":"extend",
                        "itemids": items,
                        "time_from": time_from,
                        "time_till": time_till
                    },
                    "id": 2,
                    "auth": AUTHTOKEN
                })
    data_q = json.dumps(q.json(), indent=4, sort_keys=True)
    data_q_1 = json.loads(data_q)
    for i in data_q_1['result']:
        spisok.append(float(i['value_avg']))
    if spisok == []:
        avg_len = 0
    else:
        avg_len = (sum(spisok))/(len(spisok))
    return avg_len

items_spisok = []
hosts_spisok = []
items_spisok_1 = []
hosts_spisok_1 = []

def get_items():

    r = requests.post(ZABBIX_API_URL,
                                    json={
                                        "jsonrpc": "2.0",
                                        "method": "item.get",
                                        "params": {
                                            "output": ["hostid",
                                            "groupids",
                                            "key_",
                                            ],
                                            "group": host_group,
                                            #"hostids": 11232,
                                            "search": {
                                                "name": "ICMP ping"
                                            },
                                            "sortfield": "name",
                                            "limit":1000
                                        },
                                        "id": 2,
                                        "auth": AUTHTOKEN
                                    })

    data = json.dumps(r.json(), indent=4, sort_keys=True)
    data_1 = json.loads(data)

    for items_ids in (data_1)['result']:
        items_spisok.append(int(items_ids['itemid']))
#    print(items_spisok)
#    print(data_1 )
    return(items_spisok)



def get_hosts():

    h = requests.post(ZABBIX_API_URL,
                                    json={
                                        "jsonrpc": "2.0",
                                        "method": "item.get",
                                        "params": {
                                            "output": ["hostid",
                                            "groupids",
                                            "key_",
                                            ],
                                            "group": host_group,
                                            #"hostids": 11232,
                                            "search": {
                                                "name": "ICMP ping"
                                            },
                                            "sortfield": "name",
                                            "limit":1000
                                        },
                                        "id": 2,
                                        "auth": AUTHTOKEN
                                    })

    data_b = json.dumps(h.json(), indent=4, sort_keys=True)
    data_b_1 = json.loads(data_b)

    for hosts_ids_1 in (data_b_1)['result']:
        hosts_spisok_1.append(int(hosts_ids_1['hostid']))
#    print(items_spisok)
#    print(data_1 )
#    print(hosts_spisok_1)
    return(hosts_spisok_1)

#======================== ПОЛУЧЕНИЕ ИНВЕНТАРНЫХ ДАННЫХ. НАЧАЛО =============================

def get_inventory_a(hostid,invent_a):
    j = requests.post(ZABBIX_API_URL,

                json={
                    "jsonrpc": "2.0",
                    "method": "host.get",
                    "params": {
                        "output": ["inventory"],
                        "selectInventory": [invent_a],
                        "filter": {
                            "hostid" : hostid
                        }
                    },
                    "id": 2,
                    "auth": AUTHTOKEN
                })
    data_j = json.dumps(j.json(), indent=4, sort_keys=True)
    data_j_1 = json.loads(data_j)
    print(data_j)
    if data_j_1['result'][0]['inventory'] == []:
        return "none"
    else: 
        dat = (data_j_1['result'][0]['inventory'][invent_a])
        return dat

def get_inventory_b(hostid,invent_b):
    j = requests.post(ZABBIX_API_URL,

                json={
                    "jsonrpc": "2.0",
                    "method": "host.get",
                    "params": {
                        "output": ["inventory"],
                        "selectInventory": [invent_b],
                        "filter": {
                            "hostid" : hostid
                        }
                    },
                    "id": 2,
                    "auth": AUTHTOKEN
                })
    data_j = json.dumps(j.json(), indent=4, sort_keys=True)
    data_j_1 = json.loads(data_j)
    if data_j_1['result'][0]['inventory'] == []:
        return "none"
    else: 
        dat = (data_j_1['result'][0]['inventory'][invent_b])
        return dat

def get_inventory_c(hostid,invent_c):
    j = requests.post(ZABBIX_API_URL,

                json={
                    "jsonrpc": "2.0",
                    "method": "host.get",
                    "params": {
                        "output": ["inventory"],
                        "selectInventory": [invent_c],
                        "filter": {
                            "hostid" : hostid
                        }
                    },
                    "id": 2,
                    "auth": AUTHTOKEN
                })
    data_j = json.dumps(j.json(), indent=4, sort_keys=True)
    data_j_1 = json.loads(data_j)
    if data_j_1['result'][0]['inventory'] == []:
        return "none"
    else: 
        dat = (data_j_1['result'][0]['inventory'][invent_c])
        return dat

#======================== ПОЛУЧЕНИЕ ИНВЕНТАРНЫХ ДАННЫХ. КОНЕЦ =============================

def main():
    n = 0

    invent_a = "software_app_a"
    invent_b = "software_app_b"
    invent_c = "software_app_c"

    max_length_A = 0
    max_length_B = 0
    max_length_C = 0
    max_length_D = 0
    max_length_E = 0
    max_length_F = 0
    max_length_G = 0

    ab = get_items()
    cd = get_hosts()

    index = 3

    for items,hostid in zip(ab,cd):
        print("Хост id: ",hostid)
        print("Метрика id: ",items)
        b = get_hostname(hostid)
        print("Имя хоста: ",b)
        k = "%.3f" % (get_history(items,time_from,time_till)*100)
        print("Среднее Значение: ",k)
        print(" ")
        invent_r_a = get_inventory_a(hostid,invent_a)
        print(invent_r_a)
        invent_r_b = get_inventory_b(hostid,invent_b)
        invent_r_c = get_inventory_c(hostid,invent_c)
        n = n + 1

        worksheet.write('A%s' % index, hostid , bold)
        worksheet.write('B%s' % index, items , bold)
        worksheet.write('C%s' % index, str(b) , bold)
        worksheet.write('D%s' % index, k , bold)
        worksheet.write('E%s' % index, str(invent_r_a), bold)
        worksheet.write('F%s' % index, str(invent_r_b), bold)
        worksheet.write('G%s' % index, str(invent_r_c), bold)

        max_length_A = max(max_length_A, len(str(hostid)))
        max_length_B = max(max_length_B, len(str(items)))
        max_length_C = max(max_length_C, len(b))
        max_length_D = max(max_length_D, len(str(k)))
        max_length_E = max(max_length_E, len(invent_r_a))
        max_length_F = max(max_length_F, len(invent_r_b))
        max_length_G = max(max_length_G, len(invent_r_c))
        

        index = index + 1

        worksheet.set_column('A:A', max_length_A)
        worksheet.set_column('B:B', max_length_B)
        worksheet.set_column('C:C', max_length_C)
        worksheet.set_column('D:D', max_length_D)
        worksheet.set_column('E:E', max_length_E)
        worksheet.set_column('F:F', max_length_F)
        worksheet.set_column('G:G', max_length_G)

    workbook.close()

    print("Количество полученных значений: ",n)

main()
