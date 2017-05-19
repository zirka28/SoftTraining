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
            pkeys_value = []
            if primary_keys is None:
                for key in row:
                    pkeys_value.append(row[key])
            else:
                for key in row:
                    if key in primary_keys:
                        pkeys_value.append(row[key])
            dict [tuple(pkeys_value)] = row
        return(dict, sorted_header)

def parseDate(date):
    try:
        return datetime.strptime(date, "%Y-%m-%d")
    except ValueError:
        try:
            return datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return False


def diffSearch(f_table, s_table, f_ext_key, s_ext_key, primary_keys, header):

    wrong_data_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_colour coral; font: bold on, colour white')
    right_data_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour light_green;')
    missing_data_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour light_yellow;')
    primary_key_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour light_turquoise;')

    data = {}
    for f_key in f_table[f_ext_key].keys():
        if f_key in primary_keys:
            data.update({f_key: {'value': f_table[f_ext_key].get(f_key), 'style': primary_key_style}})
        else:
            for s_key in s_table[s_ext_key].keys():
                if f_key == s_key and f_key in header:
                    if f_table[f_ext_key].get(f_key) == s_table[s_ext_key].get(s_key):
                        data.update({f_key: {'value': 0, 'style': right_data_style}})
                    else:
                        try:
                            value = float(f_table[f_ext_key].get(f_key).replace(',', '.')) - float(s_table[s_ext_key].get(s_key).replace(',', '.'))
                            if value == 0:
                                style = right_data_style
                            else:
                                style = wrong_data_style
                            data.update({f_key: {'value': value, 'style': style}})
                        except ValueError:
                            try:
                                if f_table[f_ext_key].get(f_key) is '':
                                    data.update(
                                        {f_key: {'value': str(parseDate(s_table[s_ext_key].get(s_key))), 'style': missing_data_style}})
                                elif s_table[s_ext_key].get(s_key) is '':
                                    data.update(
                                        {f_key: {'value': str(parseDate(f_table[f_ext_key].get(f_key))), 'style': missing_data_style}})
                                else:
                                    if parseDate(f_table[f_ext_key].get(f_key)) == parseDate(s_table[s_ext_key].get(s_key)):
                                        data.update(
                                            {f_key: {'value': str(parseDate(f_table[f_ext_key].get(f_key))), 'style': right_data_style}})
                                    else:
                                        data.update({f_key: {'value': str(parseDate(f_table[f_ext_key].get(f_key)) - parseDate(s_table[s_ext_key].get(s_key))),
                                                         'style': wrong_data_style}})
                            except ValueError:
                                data.update({f_key: {'value': 'wrong string', 'style': wrong_data_style}})
    return(data)

def createDiffTable(f_table, s_table, primary_keys, header):

    header_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_colour sea_green; font: bold on, colour white')

    book = xlwt.Workbook()
    sheet = book.add_sheet("Table")

    for index, col in enumerate(header):
        sheet.write(0, index, col, header_style)

    row_index = 1
    while row_index  < len(f_table):

        for f_ext_key in f_table:
            for s_ext_key in s_table:
                if f_ext_key == s_ext_key:
                    data = diffSearch(f_table, s_table, f_ext_key, s_ext_key, primary_keys, header)
                    for index, col in enumerate(header):
                        for colomn in data:
                            if colomn == col:
                                in_data = data.get(colomn)
                                value = in_data.get('value')
                                style = in_data.get('style')
                                sheet.write(row_index, index, value, style)
                    row_index = row_index +1

    book.save("table.xls")

def main():

    args = initParser().parse_args()

    first_table, header = parseCsvFile(args.address1, args.pkey)
    second_table, header = parseCsvFile(args.address2, args.pkey)

    createDiffTable(first_table, second_table, args.pkey, header)

if __name__ == '__main__':
    main()




