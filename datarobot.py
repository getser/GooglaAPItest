from __future__ import print_function
import httplib2
import os
import json

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None


SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python DataRobot'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'googe_sheets_api_creds.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


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
    interval: int() service value for moving average calculating,
    placeholder: any type to fill places if it is unable to calculate result.

    If possible, calculates moving averages using 'interval' and 'data',
    else fills place with placeholder.

    Returns:
        List of results and placeholders of the same length as 'data'.
    """
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


def load_settings(settingsfile):
    """
    settingsfile - string() filename
    Returns:
        dict() settings loaded from file.
    """

    with open(settingsfile) as json_settings_file:
        settings = json.loads(json_settings_file.read())
    return settings


def main():
    settings = load_settings('settings.json')
    spreadsheet_id = settings['spreadsheet_id']
    api_key = settings['api_key']
    moving_average_interval = settings['moving_average_interval']
    uncounted = settings['uncounted']
    work_range = settings['work_range']
    header_row = settings['header_row']

    credentials = get_credentials()
    http_auth = credentials.authorize(httplib2.Http())
    discovery_url = ('https://sheets.googleapis.com/$discovery/rest?' 'version=v4')

    service = discovery.build('sheets', 'v4', http=http_auth, discoveryServiceUrl=discovery_url, developerKey=api_key)
    spreadsheet_data = service.spreadsheets().values()

    range_string = ''.join([work_range['start_col'], str(work_range['start_row']),
                            ':',
                            work_range['end_col'], str(work_range['end_row'])])  # 'B2:J'

    result = spreadsheet_data.get(spreadsheetId=spreadsheet_id, range=range_string).execute()
    values = result.get('values', [])

    try:
        header = values[header_row]
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

    filling_column_relative, filling_start_row_relative, head = (len(header), 0, ['Moving average']) if \
                                            'Moving average' not in header else (header.index('Moving average'), 1, [])

    filling_column_string = chr(ord(work_range['start_col']) + filling_column_relative)
    filling_start_row = work_range['start_row'] + filling_start_row_relative + header_row
    filling_end_row = filling_start_row + len(current_data)
    update_range = ''.join([filling_column_string, str(filling_start_row),
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
