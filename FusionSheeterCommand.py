import adsk.core
import adsk.fusion
import adsk.cam
import traceback

from .Fusion360Utilities.Fusion360Utilities import get_app_objects
from .Fusion360Utilities.Fusion360CommandBase import Fusion360CommandBase
from .SheetsService import sheets_get_ranges, get_sheets_service, sheets_update_values
from .SheeterUtilities import get_sheet_id, get_row_id, get_display_row_id
from .SheeterModelUtilities import push_parameters, get_parameters2, get_display, \
    get_parameters_matrix, get_features2, update_local_parameters, update_local_features, \
    update_local_display, get_time_line_object_name, get_display_matrix, push_display

from .SheeterBOMUtilities import bom_map_to_sheet_array, get_bom2, bom_builder, update_local_bom, get_bom_adv, \
    update_local_bom_adv

import webbrowser

from collections import defaultdict


# Create new Google Sheet
def sheets_create(name):
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
            },
            {
                "properties": {
                    "title": 'Display',
                    'gridProperties': {
                        "frozenRowCount": 1
                    }
                }
            },
            {
                "properties": {
                    "sheetId": 4,
                    "title": 'BOM-ADV',
                    'gridProperties': {
                        "frozenRowCount": 1
                    }
                }
            }
        ]
    }

    service = get_sheets_service()
    request = service.spreadsheets().create(body=spreadsheet_body)
    spreadsheet_response = request.execute()

    return spreadsheet_response


def create_sheet_bom_parts(spreadsheet, components: adsk.fusion.ComponentList):
    service = get_sheets_service()

    requests = []
    data = []
    validation_data = []
    component_number = 0

    for component in components:
        # get_app_objects()['ui'].messageBox(get_app_objects()['root_comp'].name + '\n' + component.name)
        if component.name == get_app_objects()['root_comp'].name:
            continue

        component_sheet = {
            "addSheet": {
                "properties": {
                    "title": 'BOM-' + component.name,
                    'gridProperties': {
                        "frozenRowCount": 1
                    }
                }

            }
        }

        requests.append(component_sheet)

        headers = []
        dims = []

        headers.append('Part Number')
        dims.append(component.partNumber)

        headers.append('Description')
        dims.append(component.description)

        range_body = {
            "range": 'BOM-' + component.name,
            "values": [headers, dims]
        }

        validation_range_body = {
            "setDataValidation": {
                "range": {
                    "sheetId": 4,
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 2 + component_number,
                    # "startColumnIndex": 2,
                    "endColumnIndex": 3 + component_number
                    # "endColumnIndex": 3
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_RANGE",
                        "values": [
                            {
                                "userEnteredValue": '=\'BOM-' + component.name + '\'!B2:B1000'
                            }
                        ]
                    },
                    "inputMessage": "Pick from the list",
                    "strict": False,
                    "showCustomUi": True
                },
            }
        }

        # TODO add some dims somehow?
        data.append(range_body)
        validation_data.append(validation_range_body)

        component_number += 1

    # Create the component sheets
    create_component_sheets_body = {
        'requests': requests
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                  body=create_component_sheets_body).execute()

    # Add the values to the component sheets
    batch_update_values_request_body = {
        # How the input data should be interpreted.
        'value_input_option': 'USER_ENTERED',
        'data': data

    }
    response = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                           body=batch_update_values_request_body).execute()

    # Add the values to the component sheets
    batch_update_validation_request_body = {
        # How the input data should be interpreted.
        'requests': validation_data

    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                  body=batch_update_validation_request_body).execute()

    # TODO make dropdowns
    # https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#SetDataValidationRequest


# Not currently working for some reason
def sheets_add_protected_ranges(spreadsheet):
    service = get_sheets_service()

    sheets = spreadsheet['sheets']

    sheet_ids = {}

    # TODO sheet names should bee configurable in general
    for sheet in sheets:
        sheet_id = sheet['properties']['sheetId']
        title = sheet['properties']['title']
        sheet_ids[title] = sheet_id

    requests = [
        {
            "addProtectedRange": {
                'protectedRange': {
                    "range": {
                        "sheetId": sheet_ids['BOM'],
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0
                    },
                    "description": "BOM Headers must not be changed",
                    "warningOnly": True
                }
            }
        },
        {
            "addProtectedRange": {
                'protectedRange': {
                    "range": {
                        "sheetId": sheet_ids['Parameters'],
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0
                    },
                    "description": "Headers must match Parameter names in Fusion 360",
                    "warningOnly": True

                }
            }
        },
        {
            "addProtectedRange": {
                'protectedRange': {
                    "range": {
                        "sheetId": sheet_ids['Features'],
                        "startRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 2,

                    },
                    "description": "Name and Number driven from Parameters Sheet",
                    "warningOnly": True

                }
            }
        },
        {
            "addProtectedRange": {
                'protectedRange': {
                    "range": {
                        "sheetId": sheet_ids['Features'],
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0

                    },
                    "description": "Headers must match feature names in Fusion 360",
                    "warningOnly": True

                }
            }
        },
        {
            "addProtectedRange": {
                'protectedRange': {
                    "range": {
                        "sheetId": sheet_ids['Display'],
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0
                    },
                    "description": "Headers must match Component Occurrence names in Fusion 360",
                    "warningOnly": True

                }
            }
        },
        {
            "addProtectedRange": {
                'protectedRange': {
                    "range": {
                        "sheetId": sheet_ids['BOM-ADV'],
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0
                    },
                    "description": "Headers must match Component names in Fusion 360",
                    "warningOnly": True

                }
            }
        }

    ]

    body = {
        'requests': requests
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'],
                                                  body=body).execute()


# Create Parameters Sheet
def create_sheet_parameters(sheet_id, all_params):
    app_objects = get_app_objects()
    um = app_objects['units_manager']
    design = app_objects['design']

    if all_params:
        parameters = design.allParameters
    else:
        parameters = design.userParameters

    headers = []
    dims = []

    # TODO implement specifically for builder?
    # headers.append('name')
    # dims.append(design.rootComponent.name)

    headers.append('Part Number')
    dims.append(design.rootComponent.partNumber)

    headers.append('Description')
    dims.append(design.rootComponent.description)

    for parameter in parameters:
        headers.append(parameter.name)
        dims.append(um.formatInternalValue(parameter.value, parameter.unit, False))

    range_body = {"range": "Parameters",
                  "values": [headers, dims]}

    sheet_range = 'Parameters'

    sheets_update_values(sheet_id, sheet_range, range_body)


# Create Feature Suppresion Sheet
def create_sheet_suppression(sheet_id, all_features):
    app_objects = get_app_objects()
    um = app_objects['units_manager']
    design = app_objects['design']
    ui = app_objects['ui']

    headers = []
    dims = []

    headers.append('Part Number')
    headers.append('Description')

    dims.append('=Parameters!A2')
    dims.append('=Parameters!B2')

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

        feature_name = get_time_line_object_name(time_line_object)

        headers.append(feature_name)
        dims.append(state)

    values = [headers, dims]

    # Create Link to Parameters Sheet
    for i in range(3, 100):
        values.append(['=Parameters!A%i' % i, '=Parameters!B%i' % i])

    # Execute Request
    range_body = {"range": "Features", "values": values}

    sheet_range = 'Features'

    sheets_update_values(sheet_id, sheet_range, range_body)


# Create display sheet
def create_sheet_display(sheet_id):
    app_objects = get_app_objects()
    design = app_objects['design']

    headers = []
    dims = []

    headers.append('Display Name')
    dims.append('Current Display')

    for occurrence in design.rootComponent.allOccurrences:
        headers.append(occurrence.fullPathName)
        dims.append(str(occurrence.isLightBulbOn))

    range_body = {"range": "Display",
                  "values": [headers, dims]}

    sheet_range = 'Display'

    sheets_update_values(sheet_id, sheet_range, range_body)


# Create BOM Sheet from Current design Assembly
# Also called during push update (overwrites everything)
def create_sheet_bom(sheet_id):
    app_objects = get_app_objects()

    headers = []
    sheet_values = []

    headers.append('Part Name')
    headers.append('Description')
    headers.append('Part Number')
    headers.append('Quantity')
    headers.append('Level')

    sheet_values.append(headers)

    # Build BOM Definition
    bom_map = defaultdict(list)
    root_comp = app_objects['root_comp']
    bom_builder(bom_map, root_comp.occurrences, 0)
    bom_map_to_sheet_array(sheet_values, bom_map)

    # Create Request
    range_body = {"range": "BOM", "values": sheet_values}
    sheet_range = 'BOM'
    sheets_update_values(sheet_id, sheet_range, range_body)


# Create BOM Sheet from Current design Assembly
# Also called during push update (overwrites everything)
def create_sheet_bom_adv(sheet_id):
    app_objects = get_app_objects()
    um = app_objects['units_manager']
    design = app_objects['design']
    ui = app_objects['ui']

    headers = []
    dims = []

    headers.append('Part Number')
    headers.append('Description')

    dims.append('=Parameters!A2')
    dims.append('=Parameters!B2')

    for component in design.allComponents:
        if component.name == get_app_objects()['root_comp'].name:
            continue
        headers.append(component.name)
        dims.append(str(component.description))

    values = [headers, dims]

    # Create Link to Parameters Sheet
    for i in range(3, 100):
        values.append(['=Parameters!A%i' % i, '=Parameters!B%i' % i])

    range_body = {"range": "BOM-ADV",
                  "values": values}

    sheet_range = 'BOM-ADV'

    sheets_update_values(sheet_id, sheet_range, range_body)


# Create a dropdown based on descriptions on Parameters page
def build_sizes_dropdown(command_inputs, value_ranges):
    # sizes = get_parameters('Parameters', sheet_id)

    parameters = get_parameters2(value_ranges)

    size_drop_down = command_inputs.addDropDownCommandInput('Size', 'Current Size (Associated Sheet Row)',
                                                            adsk.core.DropDownStyles.LabeledIconDropDownStyle)

    for size in parameters:
        size_drop_down.listItems.add(size['Description'], False)

    row_id = get_row_id()

    size_drop_down.listItems.item(row_id).isSelected = True

    return size_drop_down


# Create a dropdown based on descriptions on Parameters page
def build_display_dropdown(command_inputs, value_ranges, name):
    display_captures = get_display(value_ranges)

    display_drop_down = command_inputs.addDropDownCommandInput(name, 'Current Display Capture (Associated Sheet Row)',
                                                               adsk.core.DropDownStyles.LabeledIconDropDownStyle)

    for capture in display_captures:
        display_drop_down.listItems.add(capture['Display Name'], False)

    row_id = get_display_row_id()

    display_drop_down.listItems.item(row_id).isSelected = True

    return display_drop_down


# Switch current design to a different size
class FusionSheeterSizeCommand(Fusion360CommandBase):
    def __init__(self, cmd_def, debug):
        super().__init__(cmd_def, debug)
        self.value_ranges = []

    # All size changes done during preview.  Ok commits switch
    def on_preview(self, command, inputs, args, input_values):
        app_objects = get_app_objects()
        design = app_objects['design']

        # TODO add a 'current values' to top of list
        index = input_values['Size_input'].selectedItem.index

        value_ranges = self.value_ranges

        # TODO if BOM is configurable
        # bom_items = get_bom2(value_ranges)
        # update_local_bom(bom_items)

        parameters = get_parameters2(value_ranges)
        update_local_parameters(parameters[index])

        features = get_features2(value_ranges)
        update_local_features(features[index])

        design.attributes.add('FusionSheeter', 'parameter_row_index', str(index))

        args.isValidResult = True

    def on_create(self, command, command_inputs):
        spreadsheet_id = get_sheet_id()

        value_ranges = sheets_get_ranges(spreadsheet_id, ['Parameters', 'Features'])

        self.value_ranges = value_ranges

        build_sizes_dropdown(command_inputs, value_ranges)


# Switch current design to a different size
class FusionSheeterDisplayCommand(Fusion360CommandBase):
    def __init__(self, cmd_def, debug):
        super().__init__(cmd_def, debug)
        self.value_ranges = []

    # All size changes done during preview.  Ok commits switch
    def on_preview(self, command, inputs, args, input_values):
        app_objects = get_app_objects()
        design = app_objects['design']

        # TODO add a 'current values' to top of list
        index = inputs.itemById('display_input').selectedItem.index

        value_ranges = self.value_ranges

        display_captures = get_display(value_ranges)
        update_local_display(display_captures[index])

        design.attributes.add('FusionSheeter', 'display_row_index', str(index))

        args.isValidResult = True

    def on_create(self, command, command_inputs):
        spreadsheet_id = get_sheet_id()

        # value_ranges = sheets_get2(spreadsheet_id)
        value_ranges = sheets_get_ranges(spreadsheet_id, ['Display'])

        self.value_ranges = value_ranges

        build_display_dropdown(command_inputs, value_ranges, 'display_input')


# Switch current design to a different size
class FusionSheeterDisplayCreateCommand(Fusion360CommandBase):
    def on_execute(self, command, inputs, args, input_values):

        app_objects = get_app_objects()
        design = app_objects['design']

        spreadsheet_id = get_sheet_id()
        display_row_id = get_display_row_id()

        value_ranges = sheets_get_ranges(spreadsheet_id, ['Display'])

        display_matrix = get_display_matrix(value_ranges)

        if input_values['new_existing'] == 'Create New Display Capture?':
            row_index = len(display_matrix)
            display_matrix.append([None] * len(display_matrix[0]))

            number_index = display_matrix[0].index('Display Name')
            display_matrix[-1][number_index] = input_values['new_name']

            design.attributes.add('FusionSheeter', 'display_row_index', str(row_index - 1))

        else:
            row_index = display_row_id + 1

        push_display(display_matrix, row_index, spreadsheet_id)

    def on_create(self, command, command_inputs):
        new_existing_title = command_inputs.addTextBoxCommandInput('new_existing_title', '',
                                                                   '<b>Update Existing Display Capture or Create New one? </b>',
                                                                   1, True)

        new_existing_option = command_inputs.addRadioButtonGroupCommandInput('new_existing')
        new_existing_option.listItems.add('Create New Display Capture?', True)
        new_existing_option.listItems.add('Update Existing Display (Row)?', False)

        new_name = command_inputs.addStringValueInput('new_name', 'New Capture Name', '--Name--')

    def on_input_changed(self, command_, command_inputs, changed_input, input_values):

        if changed_input.id == 'new_existing':
            if changed_input.selectedItem.name == 'Create New Display Capture?':
                command_inputs.itemById('new_name').isVisible = True
            else:
                command_inputs.itemById('new_name').isVisible = False


class FusionSheeterSyncCommand(Fusion360CommandBase):
    def __init__(self, cmd_def, debug):
        super().__init__(cmd_def, debug)
        self.value_ranges = []

    # Dialog for sync feature
    def on_create(self, command, command_inputs):

        command.setDialogInitialSize(600, 800)

        command_inputs.addTextBoxCommandInput('sync_direction_title', '', '<b>Sync Direction</b>', 1, True)
        sync_direction_group = command_inputs.addRadioButtonGroupCommandInput('sync_direction')
        sync_direction_group.listItems.add('Pull Sheets Data into Design', True)
        sync_direction_group.listItems.add('Push Design Data to Sheets', False)

        type_group = command_inputs.addGroupCommandInput('type_group', 'Sync Options')

        type_group.children.addBoolValueInput('sync_parameters', 'Sync Design Parameters?', True, '', True)
        type_group.children.addBoolValueInput('sync_bom', 'Sync Design BOM?', True, '', True)
        type_group.children.addBoolValueInput('sync_features', 'Sync Design Feature Suppression?', True, '', True)

        spreadsheet_id = get_sheet_id()

        # value_ranges = sheets_get2(spreadsheet_id)
        value_ranges = sheets_get_ranges(spreadsheet_id, ['Parameters', 'BOM', 'Features', 'BOM-ADV'])

        self.value_ranges = value_ranges

        size_drop_down = build_sizes_dropdown(command_inputs, value_ranges)

        type_group.isExpanded = False

        new_existing_title = command_inputs.addTextBoxCommandInput('new_existing_title', '',
                                                                   '<b>For Parameters and/or Features: </b>', 1, True)
        new_existing_title.isVisible = False

        new_existing_option = command_inputs.addRadioButtonGroupCommandInput('new_existing')
        new_existing_option.listItems.add('Update Existing Size?', True)
        new_existing_option.listItems.add('Create New Size?', False)
        new_existing_option.isVisible = False

        warning_text = '<br><b>Please Note: Currently Fusion Sheeter only supports pushing BOM data.</b><br>' \
                       'For significant changes to the design it may be easier to simply create a new Sheet<br><br>' \
                       'Also Note: Pushing BOM data will <b>overwrite  existing BOM values</b> in the sheet.<br>' \
                       'If you have made changes in the sheet, first pull them locally before pushing'

        warning_input = command_inputs.addTextBoxCommandInput('warning_input', '', warning_text, 6, True)
        warning_input.isVisible = False

    # Todo update displayed fields
    def on_input_changed(self, command_, command_inputs, changed_input, input_values):

        if changed_input.id == 'sync_direction':
            if changed_input.selectedItem.name == 'Push Design Data to Sheets':

                # TODO when get push working remove these
                # command_inputs.itemById('sync_parameters').value = False
                # command_inputs.itemById('sync_parameters').isEnabled = False
                command_inputs.itemById('sync_features').value = False
                command_inputs.itemById('sync_features').isEnabled = False

                command_inputs.itemById('Size').isVisible = False

                command_inputs.itemById('warning_input').isVisible = True

                # TODO when get push working set these to True
                command_inputs.itemById('new_existing_title').isVisible = True
                command_inputs.itemById('new_existing').isVisible = True

            elif changed_input.selectedItem.name == 'Pull Sheets Data into Design':

                # TODO when get push working remove these
                # command_inputs.itemById('sync_parameters').value = True
                # command_inputs.itemById('sync_parameters').isEnabled = True
                command_inputs.itemById('sync_features').value = True
                command_inputs.itemById('sync_features').isEnabled = True

                command_inputs.itemById('Size').isVisible = True

                command_inputs.itemById('warning_input').isVisible = False

                command_inputs.itemById('new_existing_title').isVisible = False
                command_inputs.itemById('new_existing').isVisible = False

    def on_execute(self, command, inputs, args, input_values):

        app_objects = get_app_objects()
        design = app_objects['design']
        ui = app_objects['ui']

        spreadsheet_id = get_sheet_id()
        value_ranges = self.value_ranges

        if input_values['sync_direction'] == 'Push Design Data to Sheets':

            row_id = get_row_id()

            if input_values['sync_bom']:
                create_sheet_bom(spreadsheet_id)

            if input_values['sync_parameters']:

                parameter_matrix = get_parameters_matrix(value_ranges)

                if input_values['new_existing'] == 'Create New Size?':
                    row_index = len(parameter_matrix)
                    parameter_matrix.append([None] * len(parameter_matrix[0]))

                    number_index = parameter_matrix[0].index('Part Number')
                    parameter_matrix[-1][number_index] = design.rootComponent.partNumber

                    description_index = parameter_matrix[0].index('Description')
                    parameter_matrix[-1][description_index] = design.rootComponent.description

                    design.attributes.add('FusionSheeter', 'parameter_row_index', str(row_index - 1))

                else:
                    row_index = row_id + 1

                push_parameters(parameter_matrix, row_index, spreadsheet_id)

            if input_values['sync_features']:
                # TODO add features push
                pass

        elif input_values['sync_direction'] == 'Pull Sheets Data into Design':

            # TODO add a 'current values' to top of list
            row_id = input_values['Size_input'].selectedItem.index

            if input_values['sync_bom']:

                # TODO Making changes here for BOM Branch
                # bom_items = get_bom2(value_ranges)
                # all_components = design.allComponents
                # change_list = update_local_bom(bom_items, all_components)

                bom_components = get_bom_adv(value_ranges)
                all_components = design.allComponents
                change_list = update_local_bom_adv(bom_components[row_id], all_components, spreadsheet_id)

                if len(change_list) != 0:
                    # change_list += 'No Metadata changes were found'
                    ui.messageBox(change_list)

            if input_values['sync_parameters']:
                # items = get_parameters('Parameters', sheet_id)
                # update_local_parameters(items[index])
                parameters = get_parameters2(value_ranges)
                update_local_parameters(parameters[row_id])

                design.attributes.add('FusionSheeter', 'parameter_row_index', str(row_id))

            if input_values['sync_features']:
                # feature_list_of_lists = get_features('Features', sheet_id)
                # update_local_features(feature_list_of_lists[index])
                features = get_features2(value_ranges)
                update_local_features(features[row_id])

                design.attributes.add('FusionSheeter', 'parameter_row_index', str(row_id))


# Command to create a new Google Sheet or link to an existing one
class FusionSheeterCreateCommand(Fusion360CommandBase):
    # Update dialog based on user selections
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

        if changed_input.id == 'parameters_option':
            if changed_input.selectedItem.name == 'All Parameters':
                input_values['warning_input'].isVisible = True
            elif changed_input.selectedItem.name == 'User Parameters Only':
                input_values['warning_input'].isVisible = False

    # Execute
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
            spreadsheet = sheets_create(name)
            new_id = spreadsheet['spreadsheetId']

            create_sheet_parameters(new_id, all_params)
            design.attributes.add('FusionSheeter', 'parameter_row_index', str(0))

            # Todo consolidate into less API calls.  Not critical
            create_sheet_bom(new_id)
            create_sheet_bom_adv(new_id)

            # Todo optional only renamed features?
            create_sheet_suppression(new_id, True)

            create_sheet_display(new_id)

            sheets_add_protected_ranges(spreadsheet)

            # add_bom_part_sheet(spreadsheet, design.allComponents[:-1])
            create_sheet_bom_parts(spreadsheet, design.allComponents)

        elif input_values['new_or_existing'] == 'Link to Existing Sheet':
            new_id = input_values['existing_sheet_id']
            # TODO add check to see if sheet id is valid

        else:
            new_id = ''
            ui.messageBox('Something went wrong creating sheet with your inputs')
            return

        design.attributes.add('FusionSheeter', 'spreadsheetId', new_id)
        design.attributes.add('FusionSheeter', 'parameter_row_index', '0')
        design.attributes.add('FusionSheeter', 'display_row_index', '0')

        url = 'https://docs.google.com/spreadsheets/d/%s/edit#gid=0' % new_id

        webbrowser.open(url, new=2)

    # Sheet creation dialog
    # Everything defined, some off by default and depend on other selections
    def on_create(self, command, command_inputs):

        command.setDialogInitialSize(600, 800)

        command_inputs.addTextBoxCommandInput('new_title', '', '<b>Create New Sheet or Link to existing?</b>', 1, True)
        new_option_group = command_inputs.addRadioButtonGroupCommandInput('new_or_existing')

        new_option_group.listItems.add('Create New Sheet', True)
        new_option_group.listItems.add('Link to Existing Sheet', False)

        instructions_text = '<br> <b>You need the spreadsheetID from your existing sheets document. </b><br><br>' \
                            'This is found by examining the hyperlink in your browser:<br><br>' \
                            '<i>https://docs.google.com/spreadsheets/d/<b>****spreadshetID*****</b>/edit#gid=0 </i>' \
                            '<br><br>Copy just the long character string in place of ****spreadshetID*****.<br><br>'

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
        parameters_option_group.listItems.add('All Parameters', False, './resources')
        parameters_option_group.listItems.add('User Parameters Only', True)
        parameters_option_group.isVisible = True

        warning_text = '<br> <b>Warning! </b><br><br>' \
                       'Equations in dimension expressions are not currently supported:<br><br>' \
                       '<i>For Example: d1 = d2 + 4 in </i>' \
                       '<br><br>Use expressions in sheets directly to create this behavior<br><br>'

        warning_input = command_inputs.addTextBoxCommandInput('warning', '',
                                                              warning_text, 16, True)
        warning_input.isVisible = False


# Simply open the associated sheet in a browser
class FusionSheeterOpenSheetCommand(Fusion360CommandBase):
    def on_execute(self, command, inputs, args, input_values):
        sheet_id = get_sheet_id()
        url = 'https://docs.google.com/spreadsheets/d/%s/edit#gid=0' % sheet_id
        webbrowser.open(url, new=2)


# Creates palette with sheet
class FusionSheeterPaletteCommand(Fusion360CommandBase):
    def on_execute(self, command, inputs, args, input_values):

        app = adsk.core.Application.get()
        ui = app.userInterface

        sheet_id = get_sheet_id()

        self.palette_id = 'sheets_palette'

        # Create and display the palette.
        palette = ui.palettes.itemById(self.palette_id)

        url = 'https://docs.google.com/spreadsheets/d/%s/edit' % sheet_id

        if not palette:
            palette = ui.palettes.add(self.palette_id, 'Fusion Sheeter', url, True, True, True, 300, 300)

            # Dock the palette to the bottom of the Fusion window.
            palette.dockingState = adsk.core.PaletteDockingStates.PaletteDockStateBottom
            # # Add handler to HTMLEvent of the palette.
            # onHTMLEvent = MyHTMLEventHandler()
            # palette.incomingFromHTML.add(onHTMLEvent)
            # handlers.append(onHTMLEvent)
            #
            # # Add handler to CloseEvent of the palette.
            # onClosed = MyCloseEventHandler()
            # palette.closed.add(onClosed)
            # handlers.append(onClosed)
        else:
            palette.htmlFileURL = url
            palette.isVisible = True


class FusionSheeterQuickPullCommand(Fusion360CommandBase):
    def on_execute(self, command, inputs, args, input_values):
        row_id = get_row_id()
        spreadsheet_id = get_sheet_id()

        value_ranges = sheets_get_ranges(spreadsheet_id, ['Parameters', 'Features'])

        # Todo if BOM is configurable
        # bom_items = get_bom2(value_ranges)
        # update_local_bom(bom_items)

        parameters = get_parameters2(value_ranges)
        update_local_parameters(parameters[row_id])

        # Todo ? Was too slow previously
        # features = get_features2(value_ranges)
        # update_local_features(features[row_id])
