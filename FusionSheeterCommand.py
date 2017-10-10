import adsk.core
import adsk.fusion
import traceback

from .Fusion360Utilities.Fusion360Utilities import get_app_objects
from .Fusion360Utilities.Fusion360CommandBase import Fusion360CommandBase

import os
import sys
import csv
import webbrowser

from collections import defaultdict, namedtuple

# Todo add and remove ass necessary only when making the call
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'lib'))

import httplib2

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

BOM_Item = namedtuple('BOM_Item', ('part_number', 'part_name', 'description', 'children', 'level'))

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


def get_sheet_id():
    app_objects = get_app_objects()
    ui = app_objects['ui']
    design = app_objects['design']

    spreadsheet_id_attribute = design.attributes.itemByName('FusionSheeter', 'spreadsheetId')

    if spreadsheet_id_attribute:
        spreadsheet_id = spreadsheet_id_attribute.value
    else:
        ui.messageBox('No Spreadsheet Associated with this model\n '
                      'Use the link button first to establish a linked sheet')
        return False

    return spreadsheet_id


def get_row_id():
    app_objects = get_app_objects()
    ui = app_objects['ui']
    design = app_objects['design']

    row_id_attribute = design.attributes.itemByName('FusionSheeter', 'parameter_row_index')

    if row_id_attribute:
        row_id = row_id_attribute.value
    else:
        return 0

    return int(row_id)


def get_sheet_data(range_name, spreadsheet_id):

    app_objects = get_app_objects()
    ui = app_objects['ui']

    service = get_sheets_service()

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_name).execute()

    if not result:
        ui.messageBox('Could not connect to Sheet.  Try again or make sure you have the correct Sheet ID')

    rows = result.get('values', [])

    dict_list = []

    for row in rows[1:]:
        row_dict = dict(zip(rows[0], row))
        dict_list.append(row_dict)

    return dict_list


def get_feature_sheet_data(range_name, spreadsheet_id):

    app_objects = get_app_objects()
    ui = app_objects['ui']

    service = get_sheets_service()

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_name).execute()

    if not result:
        ui.messageBox('Could not connect to Sheet.  Try again or make sure you have the correct Sheet ID')

    rows = result.get('values', [])

    list_of_lists = []

    for row in rows[1:]:
        row_list = list(zip(rows[0], row))
        list_of_lists.append(row_list)

    return list_of_lists


def update_parameters(size):

    app_objects = get_app_objects()
    um = app_objects['units_manager']

    # root_comp = app_objects['root_comp']

    # model_parameters = root_comp.modelParameters

    # # TODO iterate all components and also create / deal with BOM info
    # for parameter in model_parameters:
    #
    #     new_value = size.get(parameter.name)
    #     if new_value is not None:
    #         parameter.value = um.evaluateExpression(new_value, 'in')

    design = app_objects['design']

    all_parameters = design.allParameters

    for parameter in all_parameters:

        new_value = size.get(parameter.name)
        if new_value is not None:
            # TODO handle units with an attribute that is written on create.  Can be set for link
            if parameter.value != um.evaluateExpression(new_value, um.defaultLengthUnits):
                parameter.value = um.evaluateExpression(new_value, um.defaultLengthUnits)

    new_number = size.get('Part Number')
    if new_number is not None:
        design.rootComponent.partNumber = new_number

    new_description = size.get('Description')
    if new_description is not None:
        design.rootComponent.description = new_description


def update_features(feature_list):

    app_objects = get_app_objects()
    design = app_objects['design']
    ui = app_objects['ui']

    # ui.messageBox(str(feature_list))

    # Record feature suppression state
    time_line = design.timeline

    for index in reversed(range(time_line.count)):
        # ui.messageBox('timeline index: ' + str(index))
        time_line_object = time_line.item(index)

        new_state = get_tuple(feature_list, time_line_object.name)

        if new_state is not None:

            if new_state[1] == 'Unsuppressed' and time_line_object.isSuppressed:
                time_line_object.isSuppressed = False

            elif new_state[1] == 'Suppressed' and not time_line_object.isSuppressed:
                time_line_object.isSuppressed = True

    new_number = get_tuple(feature_list, 'Part Number')
    if new_number is not None:
        design.rootComponent.partNumber = new_number[1]

    new_description = get_tuple(feature_list, 'Description')
    if new_description is not None:
        design.rootComponent.description = new_description[1]


def get_tuple(feature_list, name):
    app_objects = get_app_objects()
    ui = app_objects['ui']
    feature_list = list(reversed(feature_list))

    for index, feature in enumerate(feature_list):
        # ui.messageBox('tuple index: ' + str(index))
        if feature[0] == name:
            # ui.messageBox('feature: ' + str(feature_list[index]))
            return feature_list.pop(index)

    return None



def create_sheet(name):

    spreadsheet_body = {
        "properties": {
            "title": name
        },
        'sheets': [
            {
                "properties": {
                    "title": 'Parameters',
                    'gridProperties': {
                        "frozenRowCount": 1
                    }
                }

            },
            {
                "properties": {
                    "title": 'BOM',
                    'gridProperties': {
                        "columnCount": 4,
                        "frozenRowCount": 1
                    }
                }
            },
            {
                "properties": {
                    "title": 'Features',
                    'gridProperties': {
                        "frozenRowCount": 1
                    }
                }
            }
        ]
    }

    service = get_sheets_service()

    request = service.spreadsheets().create(body=spreadsheet_body)
    response = request.execute()

    # app_objects['ui'].messageBox(str(response['spreadsheetId']))

    new_id = response['spreadsheetId']

    return new_id


def update_sheet_parameters(sheet_id, all_params):
    app_objects = get_app_objects()
    um = app_objects['units_manager']
    design = app_objects['design']

    service = get_sheets_service()

    if all_params:
        parameters = design.allParameters
    else:
        parameters = design.userParameters

    headers = []
    dims = []

    # headers.append('name')
    # dims.append(design.rootComponent.name)

    headers.append('Part Number')
    dims.append(design.rootComponent.partNumber)

    headers.append('Description')
    dims.append(design.rootComponent.description)

    for parameter in parameters:
        headers.append(parameter.name)
        dims.append(um.formatInternalValue(parameter.value, "DefaultDistance", False))

    range_body = {"range": "Parameters",
                  "values": [headers, dims]}

    request = service.spreadsheets().values().update(spreadsheetId=sheet_id, range='Parameters', body=range_body,
                                                     valueInputOption='USER_ENTERED')
    response = request.execute()


def update_sheet_suppression(sheet_id, all_features):

    app_objects = get_app_objects()
    um = app_objects['units_manager']
    design = app_objects['design']
    ui = app_objects['ui']

    service = get_sheets_service()


    headers = []
    dims = []

    headers.append('Part Number')
    dims.append(design.rootComponent.partNumber)

    headers.append('Description')
    dims.append(design.rootComponent.description)

    # Record feature suppression state
    time_line = design.timeline

    for index in range(time_line.count):

        time_line_object = time_line.item(index)

        if time_line_object.isGroup:
            state = 'Group'
        elif time_line_object.isSuppressed:
            state = 'Suppressed'
        else:
            state = 'Unsuppressed'

        # ui.messageBox(time_line_object.entity.objectType)
        feature_name = time_line_object.name
        # feature_name = time_line_object.entity.parentComponent.name + ':' + time_line_object.name
        headers.append(feature_name)
        dims.append(state)

    range_body = {"range": "Features",
                  "values": [headers, dims]}

    request = service.spreadsheets().values().update(spreadsheetId=sheet_id, range='Features', body=range_body,
                                                     valueInputOption='USER_ENTERED')
    response = request.execute()


def update_sheet_bom(sheet_id):
    app_objects = get_app_objects()
    um = app_objects['units_manager']
    design = app_objects['design']

    service = get_sheets_service()

    headers = []
    sheet_values = []

    headers.append('Part Name')
    headers.append('Description')
    headers.append('Part Number')
    headers.append('Quantity')
    headers.append('Level')

    sheet_values.append(headers)

    bom_map = defaultdict(list)
    root_comp = app_objects['root_comp']

    bom_builder(bom_map, root_comp.occurrences, 0)

    # app_objects['ui'].messageBox(str(bom_map))

    bom_to_sheet_values(sheet_values, bom_map)

    range_body = {"range": "BOM",
                  "values": sheet_values}

    request = service.spreadsheets().values().update(spreadsheetId=sheet_id, range='BOM', body=range_body,
                                                     valueInputOption='USER_ENTERED')
    response = request.execute()


def bom_to_sheet_values(sheet_values, bom_map):

    for component, occurrences in bom_map.items():

        bom_item = occurrences[0]
        sheet_values.append(
            [
                bom_item.part_name,
                bom_item.description,
                bom_item.part_number,
                str(len(occurrences)),
                bom_item.level
            ]

        )

        if len(bom_item.children) > 0:
            bom_to_sheet_values(sheet_values, bom_item.children)


def bom_builder(bom_map, occurrences, level):

    for occurrence in occurrences:

        new_item = BOM_Item(part_number=occurrence.component.partNumber, part_name=occurrence.component.name,
                            description=occurrence.component.description, children=defaultdict(list), level=level)

        bom_map[new_item.part_name].append(new_item)

        child_occurrences = occurrence.childOccurrences

        if child_occurrences.count > 0:
            bom_builder(new_item.children, child_occurrences, level+1)


def update_meta(items):

    app_objects = get_app_objects()
    design = app_objects['design']
    ui = app_objects['ui']

    change_list = ''

    all_components = design.allComponents

    for item in items:
        component = all_components.itemByName(item['Part Name'])
        if component.partNumber != item['Part Number']:
            component.partNumber = item['Part Number']
            change_list += ('Changed: ' + component.name + ' Part Number to: ' + item['Part Number'] + '\n')

        if component.description != item['Description']:
            component.description = item['Description']
            change_list += ('Changed: ' + component.name + ' Description to: ' + item['Description'] + '\n')

    if len(change_list) == 0:
        change_list += 'No Metadata changes were found'
    ui.messageBox(change_list)


# Class for a Fusion 360 Command
# Place your program logic here
# Delete the line that says "pass" for any method you want to use
class FusionSheeterSizeCommand(Fusion360CommandBase):
    # Run whenever a user makes any change to a value or selection in the addin UI
    # Commands in here will be run through the Fusion processor and changes will be reflected in  Fusion graphics area
    def on_preview(self, command, inputs, args, input_values):

        app_objects = get_app_objects()
        design = app_objects['design']

        # TODO add a 'current values' to top of list
        index = input_values['Size_input'].selectedItem.index

        sheet_id = get_sheet_id()

        sizes = get_sheet_data('Parameters', sheet_id)
        update_parameters(sizes[index])

        feature_list_of_lists = get_feature_sheet_data('Features', sheet_id)
        update_features(feature_list_of_lists[index])

        design.attributes.add('FusionSheeter', 'parameter_row_index', str(index))

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

        sheet_id = get_sheet_id()

        sizes = get_sheet_data('Parameters', sheet_id)

        size_drop_down = command_inputs.addDropDownCommandInput('Size', 'Which Size?',
                                                                adsk.core.DropDownStyles.LabeledIconDropDownStyle)
        # ao = get_app_objects()
        # ao['ui'].messageBox(str(sizes))

        for size in sizes:
            size_drop_down.listItems.add(size['Description'], False)

        row_id = get_row_id()

        size_drop_down.listItems.item(row_id).isSelected = True


class FusionSheeterBOMPullCommand(Fusion360CommandBase):

    def on_execute(self, command, inputs, args, input_values):
        sheet_id = get_sheet_id()
        items = get_sheet_data('BOM', sheet_id)
        update_meta(items)


# This needs a lot of thought and work
# Todo notify that it will overwrite any un-consumed changes in the sheet
class FusionSheeterBOMPushCommand(Fusion360CommandBase):

    def on_execute(self, command, inputs, args, input_values):

        sheet_id = get_sheet_id()

        update_sheet_bom(sheet_id)


class FusionSheeterParameterPullCommand(Fusion360CommandBase):

    def on_execute(self, command, inputs, args, input_values):

        sheet_id = get_sheet_id()
        row_id = get_row_id()
        items = get_sheet_data('Parameters', sheet_id)
        update_parameters(items[row_id])


class FusionSheeterFeaturePullCommand(Fusion360CommandBase):
    def on_execute(self, command, inputs, args, input_values):
        sheet_id = get_sheet_id()
        row_id = get_row_id()

        feature_list_of_lists = get_feature_sheet_data('Features', sheet_id)
        update_features(feature_list_of_lists[row_id])


# Todo - Does nothing still
class FusionSheeterParameterPushCommand(Fusion360CommandBase):

    def on_execute(self, command, inputs, args, input_values):

        sheet_id = get_sheet_id()
        row_id = get_row_id()
        items = get_sheet_data('Parameters', sheet_id)
        update_parameters(items[row_id])


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

        name = app_objects['app'].activeDocument.name
        name = name[:name.rfind(' v')]

        if input_values['new_or_existing'] == 'Create New Sheet':
            new_id = create_sheet(name)
            update_sheet_parameters(new_id, all_params)
            update_sheet_bom(new_id)
            # Todo optional only renamed features?
            update_sheet_suppression(new_id, True)

        elif input_values['new_or_existing'] == 'Link to Existing Sheet':
            new_id = input_values['existing_sheet_id']
            # TODO add check to see if sheet id is valid

        else:
            new_id = ''
            ui.messageBox('Something went wrong creating sheet with your inputs')
            return

        design.attributes.add('FusionSheeter', 'spreadsheetId', new_id)
        design.attributes.add('FusionSheeter', 'parameter_row_index', '0')

        url = 'https://docs.google.com/spreadsheets/d/%s/edit#gid=0' % new_id

        webbrowser.open(url, new=2)

    # Run when the user selects your command icon from the Fusion 360 UI
    # Typically used to create and display a command dialog box
    # The following is a basic sample of a dialog UI
    def on_create(self, command, command_inputs):

        # new_or_existing_input = command_inputs.addButtonRowCommandInput('new_or_existing',
        #                                                       'Create New Sheet or Link to existing?', False)
        # new_or_existing_input.listItems.add('Create New Sheet', True)
        # new_or_existing_input.listItems.add('Link to Existing Sheet', False)

        command.setDialogInitialSize(600, 800)

        command_inputs.addTextBoxCommandInput('new_title', '', '<b>Create New Sheet or Link to existing?</b>', 1, True)
        new_option_group = command_inputs.addRadioButtonGroupCommandInput('new_or_existing')

        new_option_group.listItems.add('Create New Sheet', True)
        new_option_group.listItems.add('Link to Existing Sheet', False)

        instructions_text = 'To link to an existing spreadsheet you need to input the spreadsheetID from your ' \
                            'existing sheets document. \n\n ' \
                            'This is found by examining the hyperlink displayed in your ' \
                            'browser when editing the document: \n\n' \
                            '     https://docs.google.com/spreadsheets/d/****spreadshetID*****/edit#gid=0 \n\n' \
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

        app_objects = get_app_objects()
        document = app_objects['document']
        ui = app_objects['ui']
        design = app_objects['design']

        if not document.isSaved:
            ui.messageBox('Please save document first')
            return

        folder = document.dataFile.parentFolder

        document.save('Auto Saved by Fusion Sheeter')

        sheet_id = get_sheet_id()

        sizes = get_sheet_data('Parameters', sheet_id)

        for index, size in enumerate(sizes):

            app_objects = get_app_objects()
            document = app_objects['document']
            ui = app_objects['ui']
            design = app_objects['design']

            update_parameters(size)
            design.attributes.add('FusionSheeter', 'parameter_row_index', str(index))

            document.saveAs(size['Description'], folder, 'Auto Generated by Fusion Sheeter', '')

    # Run when the user selects your command icon from the Fusion 360 UI
    # Typically used to create and display a command dialog box
    # The following is a basic sample of a dialog UI
    def on_create(self, command, command_inputs):
        pass

class FusionSheeterOpenSheetCommand(Fusion360CommandBase):
    def on_execute(self, command, inputs, args, input_values):
        sheet_id = get_sheet_id()
        url = 'https://docs.google.com/spreadsheets/d/%s/edit#gid=0' % sheet_id

        webbrowser.open(url, new=2)