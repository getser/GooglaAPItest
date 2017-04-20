from __future__ import print_function
import httplib2
import os

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
    credential_path = os.path.join(credential_dir,
                                   'googe_sheets_api_creds.json')

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
    result = []
    for val in data:
        try:
            res = float(val)
        except ValueError:
            res = val

        result.append(res)
    return result


def moving_averages(dataset, interval, placeholder):
    uncounted = placeholder
    result = []

    for idx in range(len(dataset)):
        if idx < interval - 1:
            res = uncounted

        else:
            try:
                previous_res = result[idx - 1]

                if previous_res == uncounted:
                    res = sum(dataset[idx - (interval - 1): idx + 1]) / float(interval)
                else:
                    res = previous_res + (dataset[idx] - dataset[idx - interval]) / interval

            except ValueError:
                res = uncounted
            except TypeError:
                res = uncounted

        result.append(res)

    return result


def main():
    credentials = get_credentials()
    http_auth = credentials.authorize(httplib2.Http())
    discovery_url = ('https://sheets.googleapis.com/$discovery/rest?' 'version=v4')

    spreadsheet_id = '1oiLRvnAbHnDDRBgkdR3e9mPnlANSWUEx37Aqe1iFk0c'
    api_key = 'AIzaSyA6CmM-SaYp62aGm9TpvZwbTvJ7siLIUhY'
    moving_average_interval = 3
    uncounted = 'N/A'
    work_range = {
        'start_col': 'B', 'start_row': '2',
        'end_col': 'I', 'end_row': ''
    }
    header_row = 0
    update_range = 'J2:J14'

    service = discovery.build('sheets', 'v4', http=http_auth, discoveryServiceUrl=discovery_url, developerKey=api_key)
    spreadsheet_data = service.spreadsheets().values()

    range_string = ':'.join([''.join([work_range['start_col'], work_range['start_row']]),
                            ''.join([work_range['end_col'], work_range['end_row']])])  # 'B2:J'

    result = spreadsheet_data.get(spreadsheetId=spreadsheet_id, range=range_string).execute()
    values = result.get('values', [])
    header = values[header_row]

    if values and not ('Visitors' in header and 'Date' in header):
        raise ValueError
    else:
        current_data = values[(header_row + 1):]
        source_column_index = header.index('Visitors')
        # source_column_string = chr(ord(work_range['start_col']) + source_column_index)

    filling_column_relative, filling_start_row, head = (len(header), 0, ['Moving Average']) if \
                                                                            'Moving Average' not in header else \
                                                                            (header.index('Moving Average'), 1, [])
    # filling_column_string = chr(ord(work_range['start_col']) + filling_column_relative)

    if not current_data:
        print('No data found.')
    else:
        print(current_data)

        dataset_strings = [row[source_column_index] for row in current_data]
        dataset_floats = str_to_float_list_values(dataset_strings)
        data_to_fill = head + moving_averages(dataset_floats, moving_average_interval, uncounted)
        value_range_body = {"values": [[x] for x in data_to_fill]}
        # print(data_to_fill)
        spreadsheet_data.update(spreadsheetId=spreadsheet_id, range=update_range,
                                valueInputOption='USER_ENTERED', body=value_range_body).execute()
    # request = spreadsheet_data.update(spreadsheetId=spreadsheet_id, range=update_range,
    #                                   valueInputOption='USER_ENTERED', body=data_to_fill)
    # response = request.execute()
    # print(response)

if __name__ == '__main__':
    main()
