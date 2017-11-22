import os
import sys


# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
# CLIENT_SECRET_FILE = 'client_secret.json'
CLIENT_SECRET_FILE = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'client_secret.json')
APPLICATION_NAME = 'Sheeter Fusion 360 Add-in'
DIRECTORY_NAME = 'FusionSheeter'
CREDENTIALS_JSON = 'fusion-sheeter.json'


def modules_path():
    if not hasattr(modules_path, 'path'):
        modules_path.path = os.path.join(os.path.dirname(__file__), 'lib')
    return modules_path.path


def modify_path():
    sys.path.insert(0, modules_path())


def revert_path():
    if modules_path() in sys.path:
        sys.path.remove(modules_path())


modify_path()
try:
    import httplib2

    from apiclient import discovery
    from oauth2client import client
    from oauth2client import tools
    from oauth2client.file import Storage
finally:
    revert_path()


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, DIRECTORY_NAME, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, CREDENTIALS_JSON)

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def get_sheets_service():

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    return service


# Get Values from a sheet
def sheets_get_range(spreadsheet_id, range_name):
    """ Returns a Value Range for the input ranges.

        See Sheets API Documnetation for range syntax
        Sample: 'Sheet1' returns all values on the sheet

        """

    service = get_sheets_service()

    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()

    if not result:
        # Todo change to raise?
        raise Exception('Failed to get sheet data')

    return result


# Get Value Ranges from a sheet with batch
def sheets_get_ranges(spreadsheet_id, sheet_ranges):
    """ Returns a Value Ranges for the input ranges.
    
    See Sheets API Documnetation for rangee syntax
    Sample: ['Sheet1', 'Sheet2] returns all values on these sheets
    
    """
    service = get_sheets_service()
    response = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id, ranges=sheet_ranges).execute()

    if not response:
        raise Exception('Failed to get sheet data')

    value_ranges = response.get("valueRanges", [])

    return value_ranges


# Update values in a Sheet
def sheets_update_values(sheet_id, sheet_range, range_body):
    service = get_sheets_service()
    request = service.spreadsheets().values().update(spreadsheetId=sheet_id, range=sheet_range, body=range_body,
                                                     valueInputOption='USER_ENTERED')
    response = request.execute()

    return response


# Update values in a Sheet
def sheets_append_values(sheet_id, sheet_range, range_body):
    service = get_sheets_service()
    request = service.spreadsheets().values().append(spreadsheetId=sheet_id, range=sheet_range, body=range_body,
                                                     valueInputOption='USER_ENTERED')
    response = request.execute()

    return response


# Update a spreadsheet
def sheets_update_spreadsheet(spreadsheet_id, update_request_list):

    service = get_sheets_service()

    batch_update_spreadsheet_request_body = {
        # A list of updates to apply to the spreadsheet.
        # Requests will be applied in the order they are specified.
        # If any request is not valid, no requests will be applied.
        'requests': update_request_list

    }

    request = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,
                                                 body=batch_update_spreadsheet_request_body)
    response = request.execute()

    return response


def sheets_copy_spreadsheet(spreadsheet_id):

    service = get_sheets_service()

    result = service.spreadsheets().get(spreadsheetId=spreadsheet_id, includeGridData=True).execute()

    new_result = service.spreadsheets().create(body=result).execute()

    return new_result
