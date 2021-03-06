#!/usr/bin/env python
#-*- encoding:utf-8 -*-

import re
import xlrd
import xlwt
import string
import warnings
import numpy as np
from xlrd import XLRDError

__version__ = '0.1.0'


def _create_alphabet():
    """
        Creates a dictionary with one number for each letter from a to zz
    """
    alpha = list(string.lowercase)
    beta = alpha + [x + y for x in alpha for y in alpha]

    alpha_dict = dict(zip(beta, range(len(beta))))
    return alpha_dict


def _decode_range(xls_range):

    #Create the alphabet dict
    alphabet = _create_alphabet()

    #Divide the provided range
    try:
        start, end = xls_range.lower().split(":")
    except ValueError:
        raise ValueError("Input must have a semicolon ':'")

    #Use regex to separte string from numeric string
    group_start = re.findall('\d+|\D+', start)
    group_end = re.findall('\d+|\D+', end)

   #Treat bad inputs
    if group_start[0].isdigit():
        raise TypeError("The first part of the range must be a letter")
    if group_end[0].isdigit():
        raise TypeError("The second part of the range must be a letter")
    if len(group_start) < 2:
        raise TypeError("The range must contain a numeric string")
    if len(group_end) < 2:
        raise TypeError("The range must contain a numeric string")
    if len(group_start) > 2:
        raise TypeError("Band range Input!")
    if len(group_end) > 2:
        raise TypeError("Bad range Input!")

    #Get the number associated with the string
    try:
        string_start = alphabet[group_start[0]]
        string_end = alphabet[group_end[0]]
    except KeyError:
        raise KeyError("The range providede is too big!")
    if string_start > string_end:
        return []

    #Get the number of the original string
    numeric_start = int(group_start[1])
    numeric_end = int(group_end[1])

    return [string_start, numeric_start, string_end, numeric_end]


def _convert_to_matrix(list_, row):
    return [list_[i: i + row] for i in range(0, len(list_), row)]


def xlsread(filename, xlsrange, sheet=0, read_by_column=True,
            keep_format=False, output_format='list'):

    if not isinstance(output_format, basestring):
        raise TypeError("output_format must be a string!")
    if not (output_format == 'list' or output_format == 'numpy'):
        raise NameError("output_format must be 'list' or 'numpy'!")

    try:
        fobj = xlrd.open_workbook(filename)
    except IOError:
        raise IOError("File does not exist!")

    if isinstance(sheet, int):
        try:
            workbook = fobj.sheet_by_index(sheet)
        except IndexError:
            raise IndexError("There is no such sheet!")
    elif isinstance(sheet, basestring):
        try:
            workbook = fobj.sheet_by_name(sheet)
        except XLRDError:
            raise XLRDError("There is no such sheet!")

    range_values = _decode_range(xlsrange)
    print range_values

    rows = range(range_values[1] - 1, range_values[-1])
    columns = range(range_values[0], range_values[2] + 1)

    data = []
    if read_by_column:
        for col in columns:
            for row in rows:
                try:
                    value = workbook.cell_value(row, col)
                except IndexError:
                    warnings.warn("Range with empties cells!")
                    continue

                if value:
                    data.append(value)
                else:
                    data.append(None)
    else:
        for row in rows:
            for col in columns:
                try:
                    value = workbook.cell_value(row, col)
                    print value
                except IndexError:
                    warnings.warn("Range with empties cells!")
                    continue

                if value:
                    data.append(value)
                else:
                    data.append(None)
    non_digit_flag = False
    if keep_format:
        data = _convert_to_matrix(data, len(rows))

    if output_format == "numpy" and not non_digit_flag:
        data = np.array(data, dtype=np.float)
        if keep_format:
            data = np.reshape(data, (range_values[-1], range_values[-2] + 1),
                              order='F')



    return data
