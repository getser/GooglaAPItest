from __future__ import print_function
import httplib2
import json
import sys

from apiclient import discovery

from creds_manager import get_credentials

# SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
# CLIENT_SECRET_FILE = 'client_secret.json'
# APPLICATION_NAME = 'Google Sheets API Python DataRobot'


def sheet_ids_from_commandline():
    """data: list().

    Tries to get spreadsheet id and sheet id.
    If can't get - assigns to ''. If can't convert sheet id into int() - assigns to ''.

    Returns:
        spreadsheet_id, sheet_id.
    """

    command_line_argv = sys.argv
    id_from_argv = command_line_argv[1] if len(command_line_argv) > 1 else ''
    ids_divider = '/edit#gid='

    if ids_divider in id_from_argv:
        divide_index = id_from_argv.index(ids_divider)
        spreadsheet_id, sheet_id = id_from_argv[:divide_index], id_from_argv[divide_index:]
        try:
            sheet_id = int(sheet_id)
        except ValueError:
            print('Wrong sheet id input.')
            sheet_id = ''

    else:
        spreadsheet_id, sheet_id = id_from_argv, ''

    return spreadsheet_id, sheet_id


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


def load_settings(settingsfile):
    """
    settingsfile - string() filename
    Returns:
        dict() settings loaded from file.
    """

    with open(settingsfile) as json_settings_file:
        settings = json.loads(json_settings_file.read())
    return settings


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
    spreadsheet_id_from_argv, sheet_id_from_argv = sheet_ids_from_commandline()

    settings = load_settings('settings.json')

    scopes = settings['scopes']
    client_secret_file = settings['client_secret_file']
    application_name = settings['application_name']
    api_key = settings['api_key']
    spreadsheet_id = settings['spreadsheet_id'] if not spreadsheet_id_from_argv else spreadsheet_id_from_argv
    moving_average_interval = settings['moving_average_interval']
    uncounted = settings['uncounted']  # used to substitute result data if it is unable to calculate the result
    sheet_name_from_settings = settings.get('sheet_name')  # is optional, if needed to work with any sheet except the first one
    # work_range = settings['work_range']
    # header_row = settings['header_row']

    credentials = get_credentials(client_secret_file, application_name, scopes)
    http_auth = credentials.authorize(httplib2.Http())
    discovery_url = ('https://sheets.googleapis.com/$discovery/rest?version=v4')

    service = discovery.build('sheets', 'v4', http=http_auth, discoveryServiceUrl=discovery_url, developerKey=api_key)

    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    if sheet_id_from_argv:
        for sheet in spreadsheet['sheets']:
            if sheet['merges'][0].get('sheetId') == sheet_id_from_argv:
                sheet_name_from_argv = sheet['properties'].get('title', '')
    else:
        sheet_name_from_argv = ''

    sheet_name = sheet_name_from_argv or sheet_name_from_settings or ''




    spreadsheet_data = service.spreadsheets().values()

    source_range = ''.join([work_range['start_col'], str(work_range['start_row']),
                            ':',
                            work_range['end_col'], str(work_range['end_row'])])  # 'B2:J'

    if sheet_name:
        source_range = sheet_name + '!' + source_range

    result = spreadsheet_data.get(spreadsheetId=spreadsheet_id, range=source_range).execute()
    values = result.get('values', [])

    try:
        header_index = header_row - 1
        header = values[header_index]
    except IndexError:
        print('Wrong header row or no data present!')
        raise IndexError

    if values and not ('Visitors' in header and 'Date' in header):
        print('The required columns are not found in the header!')
        raise ValueError
    else:
        current_data = values[(header_row + 1):]
        source_column_index = header.index('Visitors')
        # source_column_string = chr(ord(work_range['start_col']) + source_column_index)

    filling_column_relative, filling_start_row_relative, head = (len(header), 0, ['Moving Average']) if \
                                            'Moving Average' not in header else (header.index('Moving Average'), 1, [])

    filling_column_string = chr(ord(work_range['start_col']) + filling_column_relative)
    filling_start_row = work_range['start_row'] + filling_start_row_relative + header_row
    filling_end_row = filling_start_row + len(current_data)
    update_range = ''.join([filling_column_string, str(filling_start_row),
                            ':',
                            filling_column_string, str(filling_end_row)])
    if sheet_name:
        update_range = sheet_name + '!' + update_range

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
