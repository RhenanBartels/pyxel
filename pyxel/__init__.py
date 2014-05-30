#!/usr/bin/env python
#-*- encoding:utf-8 -*-

import re
import xlrd
import xlwt
import string

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
    start, end = xls_range.lower().split(":")

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
        string_start = min([alphabet[group_start[0]], alphabet[group_end[0]]])
        string_end = max([alphabet[group_start[0]], alphabet[group_end[0]]])
    except KeyError:
        raise KeyError("The range providede is too big!")

    #Get the number of the original string
    numeric_start = int(group_start[1])
    numeric_end = int(group_end[1])

    return [string_start, numeric_start, string_end, numeric_end]
