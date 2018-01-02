from __future__ import print_function
import httplib2
import os
import sys
import yaml
import coinmarketcap

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from datetime import datetime

if 'release' not in sys.argv:
    DEBUG_MODE = True
else:
    DEBUG_MODE = False

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets Account Update'
SPREADSHEET_INFO = 'spreadsheet_information.yaml'

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
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def get_script_path():
    """Returns:
        The filepath wrt source python execution.
    """
    return os.path.dirname(os.path.realpath(sys.argv[0]))

def get_spreadsheet_information(file_name):
    """Gets relevant spreadsheet information from storage. 

    If nothing has been stored, this call will throw and log an error. 

    Returns:
        Dictionary, with spreadsheetId and rangeName.
    """
    full_file_name = get_script_path() + '/' + file_name
    if os.path.isfile(full_file_name):
        print('Spreadsheet information file exists. Obtaining information...') if DEBUG_MODE == True else None
        with open(full_file_name, 'r') as stream:
            try:
                spreadsheet_info = yaml.load(stream)
            except:
                print('There was an error with loading the yaml file. Exiting gracefully.') if DEBUG_MODE == True else None
                sys.exit(0)
    else:
        print('Spreadsheet information file does not exist.') if DEBUG_MODE == True else None
    print('Done.') if DEBUG_MODE == True else None
    return spreadsheet_info

def get_current_time_for_update(update_time_range_name):
    """Gets current time and formats into string. 

    Returns: 
        List of dictionaries with range and values as specified by API call.
    """
    curr_time = str(datetime.now())
    data = [
        {
            'range': update_time_range_name,
            'values': [[curr_time]]
        }
    ]
    return data

def get_crypto_prices(service, spreadsheet_id, crypto_symbol_range_name, to_write_range_name, data):
    """Reads symbols from spreadsheet and makes corresponding query to CMC.

    More exactly, this parses the symbols from the spreadsheet and pulls that information from CoinMarketCap.

    Returns:
        Dictionary with 'value' string as key and real prices for cells as value of KVP.
    """
    # Note, result here is a dictionary with the range specified and then all imp info as values
    print('Getting specified spreadsheet crypto symbols...') if DEBUG_MODE == True else None
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=crypto_symbol_range_name).execute()
    print('Done.') if DEBUG_MODE == True else None
    # Saying if values doesn't exist, return [] as default value. Values is a list of lists.
    # List per every row.
    read_values = result.get('values', []) 
    flat_list_of_symbols = [item for sublist in read_values for item in sublist]
    print('Flat list of symbols:',flat_list_of_symbols) if DEBUG_MODE == True else None
    values_to_write = get_prices_from_cmc(flat_list_of_symbols)
    crypto_dict = {
        'range': to_write_range_name,
        'values': values_to_write
    }
    data.append(crypto_dict)
    return data

def get_prices_from_cmc(flat_list_of_symbols):
    """Gets prices for symbols passed in and formats correctly for API call. 

    Returns:
        List of price in USD and market cap in USD
    """
    market = coinmarketcap.Market()
    crypto_market_info = market.ticker()

    print('Getting current USD price and market cap for specified crypto symbols...') if DEBUG_MODE == True else None
    values_to_write = []
    
    # Not the cleanest... but need to double iterate rather just checking if the symbol is 
    # in the list_of_symbols because the order does matter for API call. Filter the dict first to truncate iteration size.
    filt_crypto_market_info = list(filter(lambda d: d['symbol'] in flat_list_of_symbols, crypto_market_info))
    for crypto_symbol in flat_list_of_symbols:
        for crypto_info in filt_crypto_market_info:
            if crypto_symbol == crypto_info['symbol']:
                values_to_write.append([crypto_info['price_usd'], crypto_info['market_cap_usd'], crypto_info['percent_change_24h'] + '%']) # need to add % for excel
                break

    print('Done.') if DEBUG_MODE == True else None
    return values_to_write

def main():
    """Shows basic usage of the Sheets API.

    Returns:
        Nothing, just successfully updates the spreadsheet of your choice. 
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheet_info = get_spreadsheet_information(SPREADSHEET_INFO)
    spreadsheet_id = spreadsheet_info['spreadsheet_id']
    update_time_range_name = spreadsheet_info['update_time_range_name']
    crypto_ticker_range_name = spreadsheet_info['crypto_ticker_range_name']
    to_write_range_name = spreadsheet_info['to_write_range_name']
    value_input_option = spreadsheet_info['value_input_option']

    data = get_current_time_for_update(update_time_range_name)
    data = get_crypto_prices(service, spreadsheet_id, crypto_ticker_range_name, to_write_range_name, data)

    body = {
        'valueInputOption': value_input_option,
        'data': data
    }

    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id, body=body).execute()

    print('{0} cells updated.'.format(result.get('totalUpdatedCells'))) if DEBUG_MODE == True else None

if __name__ == '__main__':
    main()