#!/usr/bin/python

import argparse
import csv
import xlwt
from datetime import datetime


def init_parser():
    parser = argparse.ArgumentParser(description='Differences between tables')
    parser.add_argument('first_address', type=str, help='first csv file address')
    parser.add_argument('second_address', type=str, help='second csv file address')
    parser.add_argument('--index', type=str, help='indexes or primary keys', nargs='+')
    return parser


def parse_csv_file(address, index):
    table = {}
    with open(address, 'rb') as csvfile:
        reader = csv.DictReader(csvfile, delimiter='\t')
        header_array = reader.fieldnames
        header = create_header_with_type(address, header_array)

        indexes = []
        for row in reader:
            new_row = {}
            index_value = []
            if index is None:
                for key in row:
                    for column in header:
                        if key == column:
                            new_value = format_type(row[key], header[column])
                            index_value.append(new_value)
                            new_row[key] = new_value
                table[tuple(index_value)] = [new_row]
            else:
                for key in row:
                    for column in header:
                        if key == column:
                            new_value = format_type(row[key], header[column])
                            if key in index:
                                index_value.append(new_value)
                            new_row[key] = new_value
                if tuple(index_value) not in indexes:
                    indexes.append(tuple(index_value))
                    table[tuple(index_value)] = [new_row]
                else:
                    table[tuple(index_value)] += [new_row]

        if index is not None:
            sorted_header = index + sorted(set(header).difference(index))
        else:
            sorted_header = header
        return table, sorted_header


def create_header_with_type(address, header_array):
    with open(address, 'rb') as csvfile:
        reader = csv.DictReader(csvfile, delimiter='\t')
        header = {}
        for row in reader:
            for key in row:
                for column_name in header_array:
                    if key == column_name:
                        header[column_name] = parse_type(row[key])
            break
    return header


def parse_type(data):
    try:
        datetime.strptime(data, "%Y-%m-%d")
        return 'date'
    except ValueError:
        try:
            datetime.strptime(data, "%Y-%m-%d %H:%M:%S")
            return 'datetime'
        except ValueError:
            try:
                float(data.replace(',', '.'))
                return 'float'
            except ValueError:
                return 'string'


def format_type(data, data_type):
    try:
        if data_type == 'date':
            new_data = datetime.strptime(data, "%Y-%m-%d")
        elif data_type == 'datetime':
            new_data = datetime.strptime(data, "%Y-%m-%d %H:%M:%S")
        elif data_type == 'float':
            new_data = float(data.replace(',', '.'))
        else:
            new_data = data
    except ValueError:
        new_data = 'Error!'
    return new_data


def create_data_with_type(table, header):
    for index in table:
        for row in table[index]:
            for key in row:
                for column_name in header:
                    if key == column_name:
                        dict(row[key])


def diff_search(a, b, primary_keys, header):
    wrong_data_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_colour coral; '
        'font: bold on, colour white')
    right_data_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour light_green;')
    missing_data_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour light_yellow;')
    primary_key_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour light_turquoise;')

    data = {}
    for f_key in a:
        if f_key in primary_keys:
            data[f_key] = {'value': a[f_key], 'style': primary_key_style}
        else:
            for s_key in b:
                if f_key == s_key and f_key in header:
                    if a[f_key] == b[s_key]:
                        data[f_key] = {'value': 0, 'style': right_data_style}
                    else:
                        try:
                            data[f_key] = {'value': a[f_key] - b[s_key], 'style': wrong_data_style}
                        except TypeError:
                            if a[f_key] is '':
                                data[f_key] = {'value': b[s_key], 'style': missing_data_style}
                            elif b[s_key] is '':
                                data[f_key] = {'value': a[f_key], 'style': missing_data_style}
                            else:
                                data[f_key] = {'value': 'wrong string', 'style': wrong_data_style}
    return data


def create_diff_table(f_table, s_table, primary_keys, header):
    book = xlwt.Workbook()
    sheet = book.add_sheet("Table")
    row_index = 1
    header_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_colour sea_green; '
        'font: bold on, colour white')
    missing_data_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour light_yellow;')
    right_data_style = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour light_green;')

    def write_row(row, style, row_index):
        for key in row:
            for i, column in enumerate(header):
                if key == column:
                    value = row[key]
                    sheet.write(row_index, i, value, style)
        new_row_index = row_index + 1
        return new_row_index

    def write_missing_row(table, row_i, used_v):
        for row in table:
            if row not in used_v:
                row_i = write_row(row, missing_data_style, row_i)
                used_v.append(row)
        return row_i, used_v

    for index, col in enumerate(header):
        sheet.write(0, index, col, header_style)

    if primary_keys is None:
        used_keys = []
        for f_ext_key in f_table:
            for s_ext_key in s_table:
                if f_ext_key == s_ext_key:
                    used_keys.append(f_ext_key)
                    for row in f_table[f_ext_key]:
                        row_index = write_row(row, right_data_style, row_index)
        for f_ext_key in f_table:
            if f_ext_key not in used_keys:
                for f_row in f_table[f_ext_key]:
                    row_index = write_row(f_row, missing_data_style, row_index)
        for s_ext_key in s_table:
            if s_ext_key not in used_keys:
                for s_row in s_table[s_ext_key]:
                    row_index = write_row(s_row, missing_data_style, row_index)

    else:
        used_keys = []
        used_values = []
        for f_ext_key in f_table:
            for s_ext_key in s_table:
                if f_ext_key == s_ext_key:
                    if len(f_table[f_ext_key]) == 1 and len(f_table[f_ext_key]) == 1:
                        used_keys.append(f_ext_key)
                        data = diff_search(f_table[f_ext_key][0], s_table[s_ext_key][0], primary_keys, header)
                        for index, col in enumerate(header):
                            for column in data:
                                if column == col:
                                    in_data = data[column]
                                    value = in_data['value']
                                    style = in_data['style']
                                    sheet.write(row_index, index, value, style)
                        row_index = row_index + 1
                    else:
                        used_keys.append(f_ext_key)
                        used_values = []
                        for f_row in f_table[f_ext_key]:
                            for s_row in s_table[s_ext_key]:
                                if f_row == s_row:
                                    row_index = write_row(f_row, right_data_style, row_index)
                                    used_values.append(f_row)
                        row_index, used_values = write_missing_row(f_table[f_ext_key], row_index, used_values)
                        row_index, used_values = write_missing_row(s_table[s_ext_key], row_index, used_values)
        for f_ext_key in f_table:
            if f_ext_key not in used_keys:
                row_index, used_values = write_missing_row(f_table[f_ext_key], row_index, used_values)
        for s_ext_key in s_table:
            if s_ext_key not in used_keys:
                row_index, used_values = write_missing_row(s_table[s_ext_key], row_index, used_values)
    book.save("table.xls")



def main():
    args = init_parser().parse_args()

    first_table, header_array = parse_csv_file(args.first_address, args.index)
    second_table, header_array = parse_csv_file(args.second_address, args.index)
    create_diff_table(first_table, second_table, args.index, header_array)


if __name__ == '__main__':
    main()
