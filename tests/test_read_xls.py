#!/usr/bin/env python
#-*-encoding:utf-8 -*-

import unittest
import pyxel
import numpy as np
from xlrd import XLRDError


class TestReadXls(unittest.TestCase):
    def test_create_alphabet(self):
        reference = [['a', 0], ['b', 1], ['h', 7]]
        alpha_dict = pyxel._create_alphabet()
        self.assertEquals(reference[0][1], alpha_dict[reference[0][0]])
        self.assertEquals(reference[1][1], alpha_dict[reference[1][0]])
        self.assertEquals(reference[2][1], alpha_dict[reference[2][0]])

    def test_decode_range(self):
        xlsrange = "A1:B1"
        reference = [0, 1, 1, 1]
        xlsrange_two = "H10:J11"
        reference_two = [7, 10, 9, 11]
        result = pyxel._decode_range(xlsrange)
        result_two = pyxel._decode_range(xlsrange_two)
        self.assertEquals(reference, result)
        self.assertEquals(reference_two, result_two)


class TesteBadInputsReadXls(unittest.TestCase):
    def test_negative_integers(self):
        neg_xlsrange = "A1:B-1"
        neg_xlsrange_two = "XX-10:BZ-15"
        self.assertRaises(KeyError, pyxel._decode_range, neg_xlsrange)
        self.assertRaises(KeyError, pyxel._decode_range, neg_xlsrange_two)

    def test_two_numeric_inputs(self):
        numeric_input = "11:B3"
        numeric_input_two = "FG30:5689"
        numeric_input_three = "-23:56"
        self.assertRaises(TypeError, pyxel._decode_range, numeric_input)
        self.assertRaises(TypeError, pyxel._decode_range, numeric_input_two)
        self.assertRaises(TypeError, pyxel._decode_range, numeric_input_three)

    def test_two_alpha_inputs(self):
        string_input = "AA:B1"
        string_input_two = "AA:BB"
        string_input_three = "ZZ:B-9"
        self.assertRaises(TypeError, pyxel._decode_range, string_input)
        self.assertRaises(TypeError, pyxel._decode_range, string_input_two)

    def test_out_of_range_inputs(self):
        out_range_input = "A1:ZZZ56"
        out_range_input_two = "ZZZ12:B1"
        self.assertRaises(KeyError, pyxel._decode_range, out_range_input)
        self.assertRaises(KeyError, pyxel._decode_range, out_range_input_two)

    def test_bad_mixture_inputs(self):
        mix_input = "a2d:b1"
        mix_input_two = "af3:b7u"
        self.assertRaises(TypeError, pyxel._decode_range, mix_input)
        self.assertRaises(TypeError, pyxel._decode_range, mix_input_two)

    def test_bad_char(self):
        bad_char_input = "@23:b45"
        self.assertRaises(KeyError, pyxel._decode_range, bad_char_input)

    def test_unorderd_alpha(self):
        unorded = "C2:A2"
        reference = []
        result = pyxel._decode_range(unorded)
        self.assertEquals(result, reference)

    def test_without_semicolon(self):
        input = "A2B3"
        self.assertRaises(ValueError, pyxel._decode_range, input)


class TestFiles(unittest.TestCase):
    def test_existing_file(self):
        filename = "rhenan.xls"
        self.assertRaises(IOError, pyxel.xlsread, filename, None, None)

    def test_sheet_index(self):
        filename = "test1.xls"
        sheet = 1
        self.assertRaises(IndexError, pyxel.xlsread, filename, None, sheet)

    def test_sheet_by_name(self):
        filename = "test1.xls"
        sheet = "test"
        self.assertRaises(XLRDError, pyxel.xlsread, filename, None, sheet)


class TestData(unittest.TestCase):
    def test_values(self):
        filename = "test1.xls"
        xlsrange = "A1:A6"
        reference = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0]
        result = pyxel.xlsread(filename, xlsrange)
        self.assertEquals(reference, result)

    def test_empty_cells(self):
        filename = "test3.xls"
        xlsrange = "A1:B3"
        reference = [1.0, 2.0, 3.0, None, 5.0, 6.0]
        result = pyxel.xlsread(filename, xlsrange)
        self.assertEquals(reference, result)

    def test_read_columns(self):
        filename = "test2.xls"
        xlsrange = "A1:B3"
        reference = [1.0, 4.0, 2.0, 5.0, 3.0, 6.0]
        result = pyxel.xlsread(filename, xlsrange, read_by_column=False)
        self.assertEquals(reference, result)

    def test_output_format(self):
        filename = "test4.xls"
        xlsrange = "A1:B7"
        reference = np.array([1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0,
                              None, None, None], dtype=np.float)
        result = pyxel.xlsread(filename, xlsrange, output_format="numpy")
        self.assertTrue(((reference == result) | (np.isnan(reference)
                                                  & np.isnan(result))).all())

    def test_values_format(self):
        pass


if __name__ == "__main__":

    unittest.main()
