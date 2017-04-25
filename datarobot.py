from __future__ import print_function
import httplib2
import json

from apiclient import discovery

from creds_manager import get_credentials


def load_settings(settingsfile):
    """
    settingsfile - string() filename

    Returns:
        dict() settings loaded from file.
    """

    try:
        with open(settingsfile) as json_settings_file:
            settings = json.loads(json_settings_file.read())
    except IOError:
        print("The setting are not loaded! Check if file 'setting.json' is present in working directory.")
        raise IOError
    return settings


def get_sheet_id_from_input():
    """
    Returns: int() sheet ID.
    """
    input_message = """Input sheet ID (9 numbers or zero at the end of URL address \n
of desired sheet plased after the "/edit#gid=")\nor just hit Enter to work with the first sheet\n>>>"""

    sheet_id_from_input = raw_input(input_message)

    while sheet_id_from_input:
        if len(sheet_id_from_input) != 9 and sheet_id_from_input != '0':
            print("Wrong value entered!")
            sheet_id_from_input = raw_input(input_message)
            continue

        try:
            sheet_id = int(sheet_id_from_input)
            return sheet_id

        except ValueError:
            print("Wrong value entered!")
            sheet_id_from_input = raw_input(input_message)

    else:
        sheet_id = 0

    return sheet_id


def column_as_string_from_number(num):
    """
    num: int() - number of column in a sheet

    Returns:
        Column name string as 'A' or 'D' or 'CF'.
    """
    return (chr(ord('A') + (num - 27) / 26) if num > 26 else '') + chr(ord('A') + (num - 1) % 26)


def str_to_float_list_values(data):
    """data: list().

    Tries to convert 'data' elements to float.
    If can't convert - leaves element unchanged.

    Returns:
        List (of the same length as 'data') of floats where were possible to convert.
    """
    result = []
    for val in data:
        try:
            res = float(val)
        except ValueError:
            res = val

        result.append(res)
    return result


def moving_averages(data, interval, placeholder):
    """data: list() - data using for calculations,
    interval: int() >0 service value for moving average calculating,
    placeholder: any type to fill places if it is unable to calculate result.

    If possible, calculates moving averages using 'interval' and 'data',
    else fills place with placeholder.

    Returns:
        List of results and placeholders of the same length as 'data'.
    """

    assert interval == int(interval)
    assert interval > 0
    assert interval < len(data)

    uncounted = placeholder
    result = []

    for idx in range(len(data)):
        if idx < interval - 1:
            res = uncounted

        else:
            try:
                previous_res = result[idx - 1]

                if previous_res == uncounted:
                    res = sum(data[idx - (interval - 1): idx + 1]) / float(interval)
                else:
                    res = previous_res + (data[idx] - data[idx - interval]) / interval

            except ValueError:
                res = uncounted
            except TypeError:
                res = uncounted

        result.append(res)

    return result


def main():
    settings = load_settings('settings.json')

    scopes = settings['scopes']
    client_secret_file = settings['client_secret_file']
    application_name = settings['application_name']
    api_key = settings['api_key']

    spreadsheet_id = settings['spreadsheet_id']

    moving_average_interval = settings['moving_average_interval']  # SMA window
    uncounted = settings['uncounted']  # used to substitute result data if it is unable to calculate the result

    credentials = get_credentials(client_secret_file, application_name, scopes)
    http_auth = credentials.authorize(httplib2.Http())
    discovery_url = 'https://sheets.googleapis.com/$discovery/rest?version=v4'

    service = discovery.build('sheets', 'v4', http=http_auth, discoveryServiceUrl=discovery_url, developerKey=api_key)
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    sheet_id = get_sheet_id_from_input()

    for sheet in spreadsheet['sheets']:
        if sheet.get('properties', {}).get('sheetId') == sheet_id:
            sheet_name = sheet['properties'].get('title', '')
            last_column = sheet.get('properties', {}).get('gridProperties', {}).get('columnCount')
            last_row = sheet.get('properties', {}).get('gridProperties', {}).get('rowCount')
            break

    last_column_string = column_as_string_from_number(last_column)

    work_range = {
        'start_col': 'A',
        'start_row': 1,
        'end_col': last_column_string,
        'end_row': last_row
    }

    spreadsheet_data = service.spreadsheets().values()

    source_range = sheet_name + '!' + ''.join([work_range['start_col'], str(work_range['start_row']),
                                                ':',
                                                work_range['end_col'], str(work_range['end_row'])])  # 'A1:J'

    result = spreadsheet_data.get(spreadsheetId=spreadsheet_id, range=source_range).execute()
    values = result.get('values', [])

    header_index = None

    for row_idx in range(len(values)):
        if 'Visitors' in values[row_idx] and 'Date' in values[row_idx]:
            header_index = row_idx
            break

    if header_index is None:
        print('Wrong header row or no data present!')
        raise IndexError
    else:
        header = values[header_index]

    current_data = values[header_index + 1:]
    source_column_index = header.index('Visitors')

    filling_column_relative, filling_start_row_relative, head = (len(header), 0, ['Moving Average']) if \
                                            'Moving Average' not in header else (header.index('Moving Average'), 1, [])

    filling_column_string = chr(ord(work_range['start_col']) + filling_column_relative)
    filling_start_row = work_range['start_row'] + filling_start_row_relative + header_index
    filling_end_row = filling_start_row + len(current_data)
    update_range = sheet_name + '!' + ''.join([filling_column_string, str(filling_start_row),
                                                ':',
                                                filling_column_string, str(filling_end_row)])

    if not current_data:
        print('No data found.')
    else:
        dataset_strings = [row[source_column_index] if row else uncounted for row in current_data]
        dataset_floats = str_to_float_list_values(dataset_strings)
        data_to_fill = head + moving_averages(dataset_floats, moving_average_interval, uncounted)
        value_range_body = {"values": [[x] for x in data_to_fill]}
        spreadsheet_data.update(spreadsheetId=spreadsheet_id, range=update_range,
                                valueInputOption='USER_ENTERED', body=value_range_body).execute()

if __name__ == '__main__':
    main()
