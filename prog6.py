#!/usr/bin/python

import argparse
import csv
import xlwt
from datetime import datetime

def initParser():
    parser = argparse.ArgumentParser(description='Differences between tables')
    parser.add_argument('address1', type=str, help='first csv file address')
    parser.add_argument('address2', type=str, help='second csv file address')
    parser.add_argument('--pkey', type=str, help='primary key', nargs='+')
    return(parser)

def parseCsvFile(address, primary_keys):
    dict = {}
    with open(address, 'rb') as csvfile:
        reader = csv.DictReader(csvfile, delimiter='\t', quoting=csv.QUOTE_NONE)

        header = reader.fieldnames
        sorted_header = primary_keys + sorted(set(header).difference(primary_keys))

        for row in reader:
            pkeys_value =[]
            if primary_keys is None:
                for i in row:
                    pkeys_value.append(row[i])
            else:
                for i in row:
                    if i in primary_keys:
                        pkeys_value.append(row[i])
            dict [tuple(pkeys_value)] = row
        return(dict)

def createHeaderSecondPart(address, primary_keys):
    with open(address, 'rb') as csvfile:
        reader = csv.reader(csvfile, delimiter='\t', quoting=csv.QUOTE_NONE)
        for row in reader:
            header = row
            second_header = []
            for coloumn in header:
                if coloumn not in primary_keys:
                    second_header.append(coloumn)
            break
        return(second_header)

def parseDate(date):
    try:
        return datetime.strptime(date, "%Y-%m-%d")
    except ValueError:
        try:
            return datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return False



def createDiffTable(a, b, primary_keys, second_header):

    book = xlwt.Workbook()
    sheet1 = book.add_sheet("PySheet1")

    cols = primary_keys + second_header
    wrong_data_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_colour coral; font: bold on, colour white')
    right_data_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour light_green;')
    missing_data_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour light_yellow;')
    primary_key_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour light_turquoise;')


    for index, col in enumerate(cols):
        sheet1.write(0, index, col, primary_key_style)

    x=1
    while x < len(a):

        for i in a:
            for k in b:
                if i == k:
                    data = {}
                    for j in a[i].keys():
                        if j in primary_keys:
                            data.update({j : {'value' : a[i].get(j), 'style' : primary_key_style}})
                        else:
                            for h in b[k].keys():
                                if j == h and j in cols:
                                    if a[i].get(j) == b[k].get(h):
                                        data.update({j: {'value': 0, 'style': right_data_style}})
                                    else:
                                        try:
                                            value = float(a[i].get(j).replace(',', '.')) - float(b[k].get(h).replace(',', '.'))
                                            if value == 0:
                                                style = right_data_style
                                            else:
                                                style = wrong_data_style
                                            data.update({j : {'value' : value, 'style' : style}})
                                        except ValueError:
                                            try:
                                                if a[i].get(j) is '':
                                                    data.update({j: {'value': str(parseDate(b[k].get(h))), 'style': missing_data_style}})
                                                elif b[k].get(h) is '':
                                                    data.update({j: {'value': str(parseDate(a[i].get(j))), 'style': missing_data_style}})
                                                else:
                                                    if parseDate(a[i].get(j)) == parseDate(b[k].get(h)):
                                                        data.update({j: {'value': str(parseDate(a[i].get(j))), 'style': right_data_style}})
                                                    else:
                                                        data.update({j: {'value': str(parseDate(a[i].get(j)) - parseDate(b[k].get(h))), 'style': wrong_data_style}})
                                            except ValueError:
                                                data.update({j: {'value': 'wrong string', 'style': wrong_data_style}})

                    for index in range(len(cols)):
                        col = cols[index]
                        for colomn in data:
                            if colomn == col:
                                in_data = data.get(colomn)
                                value = in_data.get('value')
                                style = in_data.get('style')
                                sheet1.write(x, index, value, style)
                    x=x+1

    book.save("table.xls")

def main():

    args = initParser().parse_args()
    second_header = createHeaderSecondPart(args.address1, args.pkey)

    first_table = parseCsvFile(args.address1, args.pkey)
    second_table = parseCsvFile(args.address2, args.pkey)

    createDiffTable(first_table, second_table, args.pkey, second_header)

if __name__ == '__main__':
    main()




