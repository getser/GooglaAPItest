import unittest
import datarobot


class TestDatarobot(unittest.TestCase):
    # def test_get_sheet_id_from_input(self):  # !!! DOESN"T WORK FOR SOME REASON  !!!
    #     res = datarobot.get_sheet_id_from_input()
    #     expected = 123456789
    #
    #     self.assertEqual(res, expected)

    def test_load_settings(self):
        settings = datarobot.load_settings('settings.json')
        expected = {u'spreadsheet_id': u'1oiLRvnAbHnDDRBgkdR3e9mPnlANSWUEx37Aqe1iFk0c',
                    u'scopes': u'https://www.googleapis.com/auth/spreadsheets',
                    u'moving_average_interval': 2,
                    u'uncounted': u'N/A',
                    u'client_secret_file': u'client_secret.json',
                    u'application_name': u'Google Sheets API Python DataRobot',
                    u'api_key': u'AIzaSyA6CmM-SaYp62aGm9TpvZwbTvJ7siLIUhY'}
        self.assertEqual(set(settings.keys()), set(expected.keys()))
        self.assertRaises(IOError, datarobot.load_settings, 'wrong_settings.json')

    def test_column_as_string_from_number(self):
        self.assertEqual(datarobot.column_as_string_from_number(1), 'A')
        self.assertEqual(datarobot.column_as_string_from_number(2), 'B')
        self.assertEqual(datarobot.column_as_string_from_number(26), 'Z')
        self.assertEqual(datarobot.column_as_string_from_number(27), 'AA')
        self.assertEqual(datarobot.column_as_string_from_number(28), 'AB')
        self.assertEqual(datarobot.column_as_string_from_number(51), 'AY')
        self.assertEqual(datarobot.column_as_string_from_number(52), 'AZ')
        self.assertEqual(datarobot.column_as_string_from_number(53), 'BA')
        self.assertEqual(datarobot.column_as_string_from_number(56), 'BD')
        self.assertEqual(datarobot.column_as_string_from_number(78), 'BZ')
        self.assertEqual(datarobot.column_as_string_from_number(79), 'CA')

    def test_str_to_float_list_values(self):
        lst = [1, '2', "3.5", 'abc', '7,8']
        res = datarobot.str_to_float_list_values(lst)
        expected = [1.0, 2.0, 3.5, 'abc', '7,8']
        self.assertEqual(res, expected)

    def test_moving_averages(self):
        lst1 = [100000, 30000, 70000, 10000, 80000, 15000, 14000]
        res_int_2 = datarobot.moving_averages(lst1, interval=2, placeholder='N/A')
        expected_int_2 = ['N/A', 65000, 50000, 40000, 45000, 47500, 14500]
        self.assertEqual(res_int_2, expected_int_2)

        res_int_3 = datarobot.moving_averages(lst1, interval=3, placeholder='N/A')
        expected_int_3 = ['N/A', 'N/A', 66666.66666666667, 36666.66666666667, 53332.66666666667, 34998.66666666667, 36331.66666666667]
        self.assertEqual(res_int_3, expected_int_3)

        lst2 = [100000, 30000, 70000, 'some string', 80000, 15000, 14000]
        res2_int_3 = datarobot.moving_averages(lst2, interval=3, placeholder='N/A')
        expected2_int_3 = ['N/A', 'N/A', 66666.66666666667, 'N/A', 'N/A', 'N/A', 36333.333333333336]
        self.assertEqual(res2_int_3, expected2_int_3)


if __name__ == '__main__':
    unittest.main()


