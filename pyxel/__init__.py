#!/usr/bin/env python
#-*- encoding:utf-8 -*-

import re
import xlrd
import xlwt
import string

__version__ = '0.1.0'


def _create_alfabet():
    """
        Creates a dictionary with one number of each letter from a to zz
    """
    alfa = list(string.lowercase)
    beta = alfa + [x + y for x in alfa for y in alfa]

    alfa_dict = dict(zip(beta, range(len(beta))))
    return alfa_dict

def _decode_range(xls_range):
    """
        Transform the input range (i.e 'A1:B1') into integers
    """

    #Create the alfabet dict
    alfabet = _create_alfabet()

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
        string_start = alfabet[group_start[0]]
        string_end =  alfabet[group_end[0]]
    except KeyError:
        raise KeyError("The range providede is too big!")

    #Get the number the original string
    numeric_start = int(group_start[1])
    numeric_end = int(group_end[1])


    return [string_start, numeric_start, string_end, numeric_end]





