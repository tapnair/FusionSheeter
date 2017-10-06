import adsk.core
import adsk.fusion
import traceback

from .Fusion360Utilities.Fusion360Utilities import get_app_objects
from .Fusion360Utilities.Fusion360CommandBase import Fusion360CommandBase

import os
import sys
import csv

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'lib'))

import httplib2

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse

    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/drive.file'
CLIENT_SECRET_FILE = csv_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'client_secret.json')
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


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
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
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


def get_sheet_data():
    app_objects = get_app_objects()
    um = app_objects['units_manager']
    ui = app_objects['units_manager']
    design = app_objects['design']

    spreadsheet_id_attribute = design.attributes.itemByName('FusionSheeter', 'spreadsheetId')

    if spreadsheet_id_attribute:
        spreadsheet_id = spreadsheet_id_attribute.value
    else:
        ui.messageBox('No Spreadsheet Associated with this model\n '
                      'Use the link button first to establish a linked sheet')
        return False

    service = get_sheets_service()

    # range_name = 'Sizes'
    range_name = 'Sheet1'

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_name).execute()

    if not result:
        ui.messageBox('Could not connect to Sheet.  Try again or make sure you have the correct Sheet ID')
        return False

    rows = result.get('values', [])

    dict_list = []

    for row in rows[1:]:
        row_dict = dict(zip(rows[0], row))
        dict_list.append(row_dict)

    # ao = get_app_objects()
    # ao['ui'].messageBox(str(rows))

    return dict_list


def update_parameters(size):
    app_objects = get_app_objects()
    um = app_objects['units_manager']
    # root_comp = app_objects['root_comp']

    # model_parameters = root_comp.modelParameters

    # # TODO iterate all components
    # for parameter in model_parameters:
    #
    #     new_value = size.get(parameter.name)
    #     if new_value is not None:
    #         parameter.value = um.evaluateExpression(new_value, 'in')

    all_parameters = app_objects['design'].allParameters

    for parameter in all_parameters:

        new_value = size.get(parameter.name)
        if new_value is not None:
            # TODO handle units
            if parameter.value != um.evaluateExpression(new_value, 'in'):
                parameter.value = um.evaluateExpression(new_value, 'in')


def create_sheet(all_params):
    app_objects = get_app_objects()
    um = app_objects['units_manager']
    design = app_objects['design']

    name = app_objects['app'].activeDocument.name
    spreadsheet_body = {
        "properties": {
            "title": name
        }
    }

    service = get_sheets_service()

    request = service.spreadsheets().create(body=spreadsheet_body)
    response = request.execute()

    app_objects['ui'].messageBox(str(response['spreadsheetId']))

    new_id = response['spreadsheetId']

    if all_params:
        parameters = design.allParameters
    else:
        parameters = design.userParameters

    headers = []
    dims = []

    headers.append('name')
    dims.append(design.rootComponent.name)

    headers.append('number')
    dims.append(design.rootComponent.partNumber)

    headers.append('description')
    dims.append(design.rootComponent.description)

    for parameter in parameters:
        headers.append(parameter.name)
        dims.append(um.formatInternalValue(parameter.value, "DefaultDistance", False))

    range_body = {"range": "Sheet1",
                  "values": [headers, dims]}

    request = service.spreadsheets().values().append(spreadsheetId=new_id, range='Sheet1', body=range_body,
                                                     valueInputOption='USER_ENTERED')
    response = request.execute()

    return new_id


# Class for a Fusion 360 Command
# Place your program logic here
# Delete the line that says "pass" for any method you want to use
class FusionSheeterCommand(Fusion360CommandBase):
    # Run whenever a user makes any change to a value or selection in the addin UI
    # Commands in here will be run through the Fusion processor and changes will be reflected in  Fusion graphics area
    def on_preview(self, command, inputs, args, input_values):

        # TODO add a 'current values' to top of list
        index = input_values['Size_input'].selectedItem.index

        sizes = get_sheet_data()

        if not sizes:
            return

        size = sizes[index]

        # ao = get_app_objects()
        # ao['ui'].messageBox(str(size))

        update_parameters(size)
        args.isValidResult = True

    # Run after the command is finished.
    # Can be used to launch another command automatically or do other clean up.
    def on_destroy(self, command, inputs, reason, input_values):
        pass

    # Run when any input is changed.
    # Can be used to check a value and then update the add-in UI accordingly
    def on_input_changed(self, command_, command_inputs, changed_input, input_values):
        pass

    # Run when the user presses OK
    # This is typically where your main program logic would go
    def on_execute(self, command, inputs, args, input_values):
        pass

    # Run when the user selects your command icon from the Fusion 360 UI
    # Typically used to create and display a command dialog box
    # The following is a basic sample of a dialog UI
    def on_create(self, command, command_inputs):
        sizes = get_sheet_data()

        size_drop_down = command_inputs.addDropDownCommandInput('Size', 'Which Size?',
                                                                adsk.core.DropDownStyles.LabeledIconDropDownStyle)
        ao = get_app_objects()
        ao['ui'].messageBox(str(sizes))
        for size in sizes:
            size_drop_down.listItems.add(size['description'], False)
        size_drop_down.listItems.item(0).isSelected = True


class FusionSheeterCreateCommand(Fusion360CommandBase):
    # Run whenever a user makes any change to a value or selection in the addin UI
    # Commands in here will be run through the Fusion processor and changes will be reflected in  Fusion graphics area
    def on_preview(self, command, inputs, args, input_values):
        pass

    # Run after the command is finished.
    # Can be used to launch another command automatically or do other clean up.
    def on_destroy(self, command, inputs, reason, input_values):
        pass

    # Run when any input is changed.
    # Can be used to check a value and then update the add-in UI accordingly
    def on_input_changed(self, command_, command_inputs, changed_input, input_values):

        if changed_input.id == 'new_or_existing':
            if changed_input.selectedItem.name == 'Create New Sheet':

                input_values['instructions_input'].isVisible = False
                input_values['existing_sheet_id_input'].isVisible = False
                command_inputs.itemById('param_title').isVisible = True
                input_values['parameters_option_input'].isVisible = True


            elif changed_input.selectedItem.name == 'Link to Existing Sheet':
                input_values['instructions_input'].isVisible = True
                input_values['existing_sheet_id_input'].isVisible = True
                command_inputs.itemById('param_title').isVisible = False
                input_values['parameters_option_input'].isVisible = False

    # Run when the user presses OK
    # This is typically where your main program logic would go
    def on_execute(self, command, inputs, args, input_values):

        app_objects = get_app_objects()
        design = app_objects['design']
        ui = app_objects['ui']

        all_params = True

        if input_values['parameters_option'] == 'User Parameters Only':
            all_params = False

        if input_values['new_or_existing'] == 'Create New Sheet':
            new_id = create_sheet(all_params)

        elif input_values['new_or_existing'] == 'Link to Existing Sheet':
            new_id = input_values['existing_sheet_id']
            # TODO add check to see if sheet id is valid

        else:
            new_id = ''
            ui.messageBox('Something went wrong creating sheet with your inputs')
            return

        design.attributes.add('FusionSheeter', 'spreadsheetId', new_id)

    # Run when the user selects your command icon from the Fusion 360 UI
    # Typically used to create and display a command dialog box
    # The following is a basic sample of a dialog UI
    def on_create(self, command, command_inputs):

        # new_or_existing_input = command_inputs.addButtonRowCommandInput('new_or_existing',
        #                                                       'Create New Sheet or Link to existing?', False)
        # new_or_existing_input.listItems.add('Create New Sheet', True)
        # new_or_existing_input.listItems.add('Link to Existing Sheet', False)

        command_inputs.addTextBoxCommandInput('new_title', '', '<b>Create New Sheet or Link to existing?</b>', 1, True)
        new_option_group = command_inputs.addRadioButtonGroupCommandInput('new_or_existing')

        new_option_group.listItems.add('Create New Sheet', True)
        new_option_group.listItems.add('Link to Existing Sheet', False)

        instructions_text = 'To link to an existing spreadsheet you need to input the spreadsheetID from your ' \
                            'existing sheets document. \n\n ' \
                            'This is found by examining the hyperlink displayed in your ' \
                            'browser when editing the document: \n\n' \
                            '     https://docs.google.com/spreadsheets/d/****spreadshetID*****/edit#gid=0 \n\n'\
                            'Copy just the long character string between d/ and /edit.\n\n\n'

        instructions_input = command_inputs.addTextBoxCommandInput('instructions', '',
                                                                   instructions_text, 16, True)
        instructions_input.isVisible = False

        existing_sheet_id_input = command_inputs.addStringValueInput('existing_sheet_id', 'Existing Sheet ID:', '')

        ao = get_app_objects()
        design = ao['design']
        spreadsheet_id_attribute = design.attributes.itemByName('FusionSheeter', 'spreadsheetId')

        if spreadsheet_id_attribute:
            existing_sheet_id_input.value = spreadsheet_id_attribute.value

        existing_sheet_id_input.isVisible = False

        param_title = command_inputs.addTextBoxCommandInput('param_title', '',
                                                            '<b>Parameters to include in new sheet:</b>', 1, True)
        param_title.isVisible = True

        parameters_option_group = command_inputs.addRadioButtonGroupCommandInput('parameters_option')
        parameters_option_group.listItems.add('All Parameters', True, './resources')
        parameters_option_group.listItems.add('User Parameters Only', False)
        parameters_option_group.isVisible = True


class FusionSheeterBuildCommand(Fusion360CommandBase):
    # Run whenever a user makes any change to a value or selection in the addin UI
    # Commands in here will be run through the Fusion processor and changes will be reflected in  Fusion graphics area
    def on_preview(self, command, inputs, args, input_values):
        pass

    # Run after the command is finished.
    # Can be used to launch another command automatically or do other clean up.
    def on_destroy(self, command, inputs, reason, input_values):
        pass

    # Run when any input is changed.
    # Can be used to check a value and then update the add-in UI accordingly
    def on_input_changed(self, command_, command_inputs, changed_input, input_values):
        pass

    # Run when the user presses OK
    # This is typically where your main program logic would go
    def on_execute(self, command, inputs, args, input_values):
        pass
        # TODO export all sizes to folder

    # Run when the user selects your command icon from the Fusion 360 UI
    # Typically used to create and display a command dialog box
    # The following is a basic sample of a dialog UI
    def on_create(self, command, command_inputs):
        pass
