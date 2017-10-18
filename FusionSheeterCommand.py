import adsk.core
import adsk.fusion
import adsk.cam
import traceback

from .Fusion360Utilities.Fusion360Utilities import get_app_objects
from .Fusion360Utilities.Fusion360CommandBase import Fusion360CommandBase

import os
import sys
import csv
import webbrowser

from collections import defaultdict, namedtuple

# Todo add and remove ass necessary only when making the call
# sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'lib'))

# Globals for multiple command execution
operation_list = []
post_name = None
output_folder = None
the_file_name = None
gcode_index = 0
gcode_test = 25
params_list = []
feature_list = []

BOM_Item = namedtuple('BOM_Item', ('part_number', 'part_name', 'description', 'children', 'level'))

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/drive.file'
CLIENT_SECRET_FILE = csv_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'client_secret.json')
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


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


def get_sheets_service():

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    return service


# Get Values from a sheet
def sheets_get(spreadsheet_id, range_name):

    service = get_sheets_service()
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()

    if not result:
        # Todo change to raise?
        app_objects = get_app_objects()
        ui = app_objects['ui']
        ui.messageBox('Could not connect to Sheet.  Try again or make sure you have the correct Sheet ID')
        adsk.terrminate()

    return result


# Update values in a Sheet
def sheets_update(sheet_id, sheet_range, range_body):
    service = get_sheets_service()
    request = service.spreadsheets().values().update(spreadsheetId=sheet_id, range=sheet_range, body=range_body,
                                                     valueInputOption='USER_ENTERED')
    response = request.execute()


def get_parameters(range_name, spreadsheet_id):

    result = sheets_get(spreadsheet_id, range_name)

    rows = result.get('values', [])

    dict_list = []

    for row in rows[1:]:
        row_dict = dict(zip(rows[0], row))
        dict_list.append(row_dict)

    return dict_list


def get_features(range_name, spreadsheet_id):

    result = sheets_get(spreadsheet_id, range_name)

    rows = result.get('values', [])

    list_of_lists = []

    for row in rows[1:]:
        row_list = list(zip(rows[0], row))
        list_of_lists.append(row_list)

    return list_of_lists


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


def update_local_parameters(size):
    app_objects = get_app_objects()
    um = app_objects['units_manager']

    design = app_objects['design']

    all_parameters = design.allParameters

    for parameter in all_parameters:

        new_value = size.get(parameter.name)
        if new_value is not None:

            if um.isValidExpression(new_value, um.defaultLengthUnits):
                sheet_value = um.evaluateExpression(new_value, um.defaultLengthUnits)

                # TODO handle units with an attribute that is written on create.  Can be set for link
                if parameter.value != sheet_value:
                    parameter.value = sheet_value

    new_number = size.get('Part Number')
    if new_number is not None:
        design.rootComponent.partNumber = new_number

    new_description = size.get('Description')
    if new_description is not None:
        design.rootComponent.description = new_description

        # Todo create and display change list (not during size change, need variable to control)


# Update local feature suppression state from sheets data
def update_local_features(feature_list):
    app_objects = get_app_objects()
    design = app_objects['design']

    # Record feature suppression state
    time_line = design.timeline

    # Going to iterate time line in reverse
    feature_list = list(reversed(feature_list))

    # Walk timeline in reverse order
    for index in reversed(range(time_line.count)):

        time_line_object = time_line.item(index)

        feature_name = get_time_line_object_name(time_line_object)

        new_state = find_list_item(feature_list, feature_name)

        # Set suppression state from sheet if different than current
        if new_state is not None:

            if new_state[1] == 'Unsuppressed' and time_line_object.isSuppressed:
                time_line_object.isSuppressed = False

            elif new_state[1] == 'Suppressed' and not time_line_object.isSuppressed:
                time_line_object.isSuppressed = True

    # Update Meta Data
    # TODO probably remove this from here
    new_number = find_list_item(feature_list, 'Part Number')
    if new_number is not None:
        design.rootComponent.partNumber = new_number[1]

    new_description = find_list_item(feature_list, 'Description')
    if new_description is not None:
        design.rootComponent.description = new_description[1]

        # Todo create and display change list (not during size change, need variable to control)


# Search sheets feature list for item pops it from list
def find_list_item(feature_list, name):
    for index, feature in enumerate(feature_list):
        if feature[0] == name:
            return feature_list.pop(index)

    return None


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
            }
        ]
    }

    service = get_sheets_service()
    request = service.spreadsheets().create(body=spreadsheet_body)
    spreadsheet_response = request.execute()

    return spreadsheet_response


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
        dims.append(um.formatInternalValue(parameter.value, "DefaultDistance", False))

    range_body = {"range": "Parameters",
                  "values": [headers, dims]}

    sheet_range = 'Parameters'

    sheets_update(sheet_id, sheet_range, range_body)


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

    sheets_update(sheet_id, sheet_range, range_body)



# Create Unique name for time line objects
def get_time_line_object_name(time_line_object):
    app_objects = get_app_objects()
    root_comp = app_objects['root_comp']
    feature_name = ''

    try:
        component = time_line_object.entity.parentComponent
        if root_comp.revisionId != component.revisionId:
            feature_name += component.name + ':'
    except:
        pass

    feature_name += time_line_object.name

    return feature_name


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
    sheets_update(sheet_id, sheet_range, range_body)


# Appends BOM information to sheet values array
def bom_map_to_sheet_array(sheet_values, bom_map):
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
            bom_map_to_sheet_array(sheet_values, bom_item.children)


# Maps current design components to BOM structured as dict
def bom_builder(bom_map, occurrences, level):
    for occurrence in occurrences:

        new_item = BOM_Item(part_number=occurrence.component.partNumber, part_name=occurrence.component.name,
                            description=occurrence.component.description, children=defaultdict(list), level=level)

        bom_map[new_item.part_name].append(new_item)

        child_occurrences = occurrence.childOccurrences

        if child_occurrences.count > 0:
            bom_builder(new_item.children, child_occurrences, level + 1)


# Update local BOM metadata based on sheets info
def update_local_bom(items):
    app_objects = get_app_objects()
    design = app_objects['design']
    ui = app_objects['ui']

    change_list = ''

    all_components = design.allComponents

    for item in items:

        # TODO Allow for editing component names?
        component = all_components.itemByName(item['Part Name'])
        if component is not None:
            if component.partNumber != item['Part Number']:
                component.partNumber = item['Part Number']
                change_list += ('Changed: ' + component.name + ' Part Number to: ' + item['Part Number'] + '\n')

            if component.description != item['Description']:
                component.description = item['Description']
                change_list += ('Changed: ' + component.name + ' Description to: ' + item['Description'] + '\n')

    if len(change_list) != 0:
        # change_list += 'No Metadata changes were found'
        ui.messageBox(change_list)


# Create a dropdown based on descriptions on Parameters page
def build_sizes_dropdown(command_inputs, sheet_id):

    sizes = get_parameters('Parameters', sheet_id)

    size_drop_down = command_inputs.addDropDownCommandInput('Size', 'Current Size (Associated Sheet Row)',
                                                            adsk.core.DropDownStyles.LabeledIconDropDownStyle)

    for size in sizes:
        size_drop_down.listItems.add(size['Description'], False)

    row_id = get_row_id()

    size_drop_down.listItems.item(row_id).isSelected = True

    return size_drop_down


# Updates tool paths and outputs new nc file
def update_g_code(op_name_, post_name_, output_folder_, index, prefix):
    global gcode_index

    # TODO use app objects, probably need to switch workspaces?
    ao = get_app_objects()
    app = adsk.core.Application.get()
    doc = app.activeDocument
    products = doc.products
    product = products.itemByProductType('CAMProductType')
    cam = adsk.cam.CAM.cast(product)
    adsk.doEvents()

    to_post = None
    # Iterate through CAM objects for operation, folder or setup
    # Currently doesn't handle duplicate in names
    for setup in cam.setups:
        if setup.name == op_name_:
            to_post = setup
        else:
            for folder in setup.folders:
                if folder.name == op_name_:
                    to_post = folder

    for operation in cam.allOperations:
        if operation.name == op_name_:
            to_post = operation

    if to_post is None:
        return

    # Update tool path
    future = cam.generateToolpath(to_post)
    check = 0
    while not future.isGenerationCompleted:
        adsk.doEvents()
        import time

        time.sleep(1)
        check += 1
        if check > 10:
            ao['ui'].messageBox('Timeout')
            break

    # Set the post options
    post_config = os.path.join(cam.genericPostFolder, post_name_)
    units = adsk.cam.PostOutputUnitOptions.DocumentUnitsOutput

    program_name = prefix + '_' + op_name_

    # create the postInput object
    post_input = adsk.cam.PostProcessInput.create(program_name, post_config, output_folder_, units)
    post_input.isOpenInEditor = False

    cam.postProcess(to_post, post_input)


# Get default directory
def get_default_model_dir(model_name):
    # Get user's home directory
    default_dir = os.path.expanduser("~")

    model_name = model_name[:model_name.rfind(' v')]

    # Create a subdirectory for this application
    default_dir = os.path.join(default_dir, 'FusionSheeterOutput', '')

    # Create sub directory for specific model
    default_dir = os.path.join(default_dir, model_name, '')

    # Create the folder if it does not exist
    if not os.path.exists(default_dir):
        os.makedirs(default_dir)

    return default_dir


# Switch Current Workspace
def switch_workspace(workspace_name):

    ao = get_app_objects()
    workspace = ao['ui'].workspaces.itemById(workspace_name)
    workspace.activate()


# Execute a command based on ID
def execute_next_command(cmd_id):

    app_objects = get_app_objects()
    next_command = app_objects['ui'].commandDefinitions.itemById(cmd_id)
    next_command.execute()


# Will update visibility of 3 selection dropdowns based on radio selection
# Also updates radio selection which is only really useful when command is first launched.
def gcode_drop_down_set_visibility(changed_input, inputs):
    drop_down = inputs.itemById(changed_input.id[6:])

    if changed_input.value:
        drop_down.isVisible = True
    else:

        for item in drop_down.listItems:
            item.isSelected = False
        drop_down.isVisible = False


def gcode_add_operations(input_list, op_list):
    for item in input_list:
        if item.isSelected:
            op_list.append(item.name)


def dup_check(name):
    if os.path.exists(name):
        base, ext = os.path.splitext(name)
        base += '-dup'
        name = base + ext
        dup_check(name)
    return name


def export_active_doc(folder, file_types, write_version, index):
    app = adsk.core.Application.get()
    design = app.activeProduct
    export_mgr = design.exportManager

    export_functions = [export_mgr.createIGESExportOptions,
                        export_mgr.createSTEPExportOptions,
                        export_mgr.createSATExportOptions,
                        export_mgr.createSMTExportOptions,
                        export_mgr.createFusionArchiveExportOptions,
                        export_mgr.createSTLExportOptions]
    export_extensions = ['.igs', '.step', '.sat', '.smt', '.f3d', '.stl']

    for i in range(file_types.count):

        if file_types.item(i).isSelected:

            doc_name = app.activeDocument.name

            if not write_version:
                doc_name = doc_name[:doc_name.rfind(' v')]

            export_name = folder + doc_name + '_' + str(index) + export_extensions[i]
            export_name = dup_check(export_name)
            export_options = export_functions[i](export_name)
            export_mgr.execute(export_options)


# Switch current design to a different size
class FusionSheeterSizeCommand(Fusion360CommandBase):
    # All size changes done during preview.  Ok commits switch
    def on_preview(self, command, inputs, args, input_values):
        app_objects = get_app_objects()
        design = app_objects['design']

        # TODO add a 'current values' to top of list
        index = input_values['Size_input'].selectedItem.index

        sheet_id = get_sheet_id()

        sizes = get_parameters('Parameters', sheet_id)
        update_local_parameters(sizes[index])

        feature_list_of_lists = get_features('Features', sheet_id)
        update_local_features(feature_list_of_lists[index])

        design.attributes.add('FusionSheeter', 'parameter_row_index', str(index))

        args.isValidResult = True

    def on_create(self, command, command_inputs):
        sheet_id = get_sheet_id()

        size_drop_down = build_sizes_dropdown(command_inputs, sheet_id)


class FusionSheeterSyncCommand(Fusion360CommandBase):
    # Dialog for sync feature
    def on_create(self, command, command_inputs):

        sheet_id = get_sheet_id()

        command.setDialogInitialSize(600, 800)

        command_inputs.addTextBoxCommandInput('sync_direction_title', '', '<b>Sync Direction</b>', 1, True)
        sync_direction_group = command_inputs.addRadioButtonGroupCommandInput('sync_direction')
        sync_direction_group.listItems.add('Pull Sheets Data into Design', True)
        sync_direction_group.listItems.add('Push Design Data to Sheets', False)

        type_group = command_inputs.addGroupCommandInput('type_group', 'Sync Options')

        type_group.children.addBoolValueInput('sync_parameters', 'Sync Design Parameters?', True, '', True)
        type_group.children.addBoolValueInput('sync_bom', 'Sync Design BOM?', True, '', True)
        type_group.children.addBoolValueInput('sync_features', 'Sync Design Feature Suppression?', True, '', True)

        size_drop_down = build_sizes_dropdown(type_group.children, sheet_id)

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
                command_inputs.itemById('sync_parameters').value = False
                command_inputs.itemById('sync_parameters').isEnabled = False
                command_inputs.itemById('sync_features').value = False
                command_inputs.itemById('sync_features').isEnabled = False

                command_inputs.itemById('Size').isVisible = False

                command_inputs.itemById('warning_input').isVisible = True

                # TODO when get push working set these to True
                command_inputs.itemById('new_existing_title').isVisible = False
                command_inputs.itemById('new_existing').isVisible = False

            elif changed_input.selectedItem.name == 'Pull Sheets Data into Design':

                # TODO when get push working remove these
                command_inputs.itemById('sync_parameters').value = True
                command_inputs.itemById('sync_parameters').isEnabled = True
                command_inputs.itemById('sync_features').value = True
                command_inputs.itemById('sync_features').isEnabled = True

                command_inputs.itemById('Size').isVisible = True

                command_inputs.itemById('warning_input').isVisible = False

                command_inputs.itemById('new_existing_title').isVisible = False
                command_inputs.itemById('new_existing').isVisible = False

    def on_execute(self, command, inputs, args, input_values):

        app_objects = get_app_objects()
        design = app_objects['design']

        # TODO add a 'current values' to top of list
        index = input_values['Size_input'].selectedItem.index

        sheet_id = get_sheet_id()

        if input_values['sync_direction'] == 'Push Design Data to Sheets':

            if input_values['sync_bom']:
                create_sheet_bom(sheet_id)

            if input_values['sync_parameters']:
                # Todo add ability to push parameter and feature data
                # Todo handle create new vs. update existing
                #
                pass

            if input_values['sync_features']:
                pass

        elif input_values['sync_direction'] == 'Pull Sheets Data into Design':

            if input_values['sync_bom']:
                items = get_parameters('BOM', sheet_id)
                update_local_bom(items)

            if input_values['sync_parameters']:
                items = get_parameters('Parameters', sheet_id)

                update_local_parameters(items[index])

                design.attributes.add('FusionSheeter', 'parameter_row_index', str(index))

            if input_values['sync_features']:
                feature_list_of_lists = get_features('Features', sheet_id)

                update_local_features(feature_list_of_lists[index])

                design.attributes.add('FusionSheeter', 'parameter_row_index', str(index))


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

            create_sheet_bom(new_id)

            # Todo optional only renamed features?
            create_sheet_suppression(new_id, True)

            sheets_add_protected_ranges(spreadsheet)

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


# Command to build all sizes in a sheet
# Todo better control for naming and for destination
class FusionSheeterBuildCommand(Fusion360CommandBase):
    # Execute model creation
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

        sizes = get_parameters('Parameters', sheet_id)

        feature_list_of_lists = get_features('Features', sheet_id)

        for index, size in enumerate(sizes):

            update_local_parameters(size)
            update_local_features(feature_list_of_lists[index])

            design.attributes.add('FusionSheeter', 'parameter_row_index', str(index))

            document.saveAs(size['Description'], folder, 'Auto Generated by Fusion Sheeter', '')


# Command to build all sizes in a sheet
# Todo better control for naming and for destination
class FusionSheeterExportCommand(Fusion360CommandBase):

    # Execute model creation
    def on_execute(self, command, inputs, args, input_values):

        sheet_id = get_sheet_id()

        sizes = get_parameters('Parameters', sheet_id)
        feature_list_of_lists = get_features('Features', sheet_id)

        folder = input_values['output_folder']
        file_types = input_values['file_types_input'].listItems
        write_version = input_values['write_version']

        for index, size in enumerate(sizes):
            update_local_parameters(size)
            update_local_features(feature_list_of_lists[index])

            export_active_doc(folder, file_types, write_version, index)

    def on_create(self, command, command_inputs):
        app = adsk.core.Application.get()
        default_dir = get_default_model_dir(app.activeDocument.name)

        command_inputs.addStringValueInput('output_folder', 'Output Folder:', default_dir)

        drop_input_list = command_inputs.addDropDownCommandInput('file_types', 'Export Types',
                                                                 adsk.core.DropDownStyles.CheckBoxDropDownStyle)

        drop_input_list = drop_input_list.listItems
        drop_input_list.add('IGES', False)
        drop_input_list.add('STEP', True)
        drop_input_list.add('SAT', False)
        drop_input_list.add('SMT', False)
        drop_input_list.add('F3D', False)
        # drop_input_list.add('STL', False)

        command_inputs.addBoolValueInput('write_version', 'Write versions to output file names?', True)


# Simply open the associated sheet in a browser
class FusionSheeterOpenSheetCommand(Fusion360CommandBase):
    def on_execute(self, command, inputs, args, input_values):
        sheet_id = get_sheet_id()
        url = 'https://docs.google.com/spreadsheets/d/%s/edit#gid=0' % sheet_id
        webbrowser.open(url, new=2)


# Class for initial Model Definition and import.
class FusionSheeterGCodeCommand(Fusion360CommandBase):
    # This is typically where your main program logic would go
    def on_execute(self, command, inputs, args, input_values):
        global operation_list
        global post_name
        global output_folder
        global gcode_index
        global gcode_test
        global params_list
        global feature_list

        operation_list.clear()

        # Get the values from the user input
        # operation_list = input_values['setup_name']
        post_name = input_values['post_name']
        output_folder = input_values['output_folder']

        gcode_add_operations(input_values['setups'], operation_list)
        gcode_add_operations(input_values['folders'], operation_list)
        gcode_add_operations(input_values['operations'], operation_list)

        gcode_index = 0

        sheet_id = get_sheet_id()
        params_list = get_parameters('Parameters', sheet_id)
        feature_list = get_features('Features', sheet_id)

        gcode_test = len(params_list)

        update_local_parameters(params_list[gcode_index])
        update_local_features(feature_list[gcode_index])

        execute_next_command('cmdID_FusionSheeterGCodeCommand2')

    # Creates a dialog UI
    def on_create(self, command, command_inputs):

        app = adsk.core.Application.get()
        ui = app.userInterface
        doc = app.activeDocument
        products = doc.products
        product = products.itemByProductType('CAMProductType')

        default_dir = get_default_model_dir(app.activeDocument.name)

        # Check if the document has a CAMProductType. It will not if there are no CAM operations in it.
        if product == None:
            ui.messageBox('There are no CAM operations in the active document')
            return
            # Cast the CAM product to a CAM object (a subtype of product).
        cam = adsk.cam.CAM.cast(product)

        # Create a few inputs in the UI
        # command_inputs.addStringValueInput('setup_name', 'CAM Setup Name', 'Setup1')
        # Todo auto populate post list
        command_inputs.addStringValueInput('post_name', 'Post to Use:', 'grbl.cps')
        command_inputs.addStringValueInput('output_folder', 'Output Folder:', default_dir)

        # What to select from?  Setups, Folders, Operations?
        group_input = command_inputs.addGroupCommandInput("showOperations", 'What to Post?')
        group_input.children.addBoolValueInput("group_setups", "Setups", True, '', True)
        group_input.children.addBoolValueInput("group_folders", "Folders", True, '', False)
        group_input.children.addBoolValueInput("group_operations", "Operations", True, '', False)

        # Drop down for Setups
        setup_drop_down = command_inputs.addDropDownCommandInput('setups', 'Select Setup(s):',
                                                                 adsk.core.DropDownStyles.CheckBoxDropDownStyle)
        # Drop down for Folders
        folder_drop_down = command_inputs.addDropDownCommandInput('folders', 'Select Folder(s):',
                                                                  adsk.core.DropDownStyles.CheckBoxDropDownStyle)
        # Drop down for Operations
        op_drop_down = command_inputs.addDropDownCommandInput('operations', 'Select Operation(s):',
                                                              adsk.core.DropDownStyles.CheckBoxDropDownStyle)

        # Populate values in drop downs based on current document:

        for setup in cam.setups:
            setup_drop_down.listItems.add(setup.name, False)
            for folder in setup.folders:
                folder_drop_down.listItems.add(folder.name, False)
        for operation in cam.allOperations:
            op_drop_down.listItems.add(operation.name, False)

        op_drop_down.isVisible = False
        folder_drop_down.isVisible = False

    # Update dialog based on user selections
    def on_input_changed(self, command_, command_inputs, changed_input, input_values):
        # Check to see if the post type has changed and show appropriate drop down
        if changed_input.parentCommandInput:
            if changed_input.parentCommandInput.id == 'showOperations':
                gcode_drop_down_set_visibility(changed_input, command_inputs)


class FusionSheeterGCodeCommand2(Fusion360CommandBase):
    # Update the G-Code
    def on_execute(self, command, inputs, args, input_values):
        global operation_list
        global post_name
        global output_folder
        global gcode_index
        global gcode_test
        global params_list
        global feature_list

        switch_workspace('CAMEnvironment')

        # app = adsk.core.Application.get()
        # ui = app.userInterface
        # ui.messageBox(operation_list)

        # Todo make this configurable
        prefix = params_list[gcode_index]['Part Number'] + '_' + params_list[gcode_index]['Description']
        for operation in operation_list:
            update_g_code(operation, post_name, output_folder, gcode_index, prefix)

        switch_workspace('FusionSolidEnvironment')

        gcode_index += 1

        if gcode_index < gcode_test:
            execute_next_command('cmdID_FusionSheeterGCodeCommand3')


# Class for initial Model Definition and import.
class FusionSheeterGCodeCommand3(Fusion360CommandBase):
    # Run when the user presses OK
    # This is typically where your main program logic would go
    def on_execute(self, command, inputs, args, input_values):
        global operation_list
        global post_name
        global output_folder
        global the_file_name
        global params_list
        global gcode_index
        global feature_list

        update_local_parameters(params_list[gcode_index])
        update_local_features(feature_list[gcode_index])

        execute_next_command('cmdID_FusionSheeterGCodeCommand2')


# Class for initial Model Definition and import.
class FusionSheeterPaletteCommand(Fusion360CommandBase):
    # Run when the user presses OK
    # This is typically where your main program logic would go
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

            # Dock the palette to the right side of Fusion window.
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


# # Class for initial Model Definition and import.
# class FusionSheeterQuickPullCommand(Fusion360CommandBase):
#     # Run when the user presses OK
#     # This is typically where your main program logic would go
#     def on_execute(self, command, inputs, args, input_values):
#         row_id = get_row_id()
#
#         items = get_parameters('BOM', sheet_id)
#         update_local_bom(items)
#
#         items = get_parameters('Parameters', sheet_id)
#
#         update_local_parameters(items[index])
#
#
#         feature_list_of_lists = get_features('Features', sheet_id)
#
#         update_local_features(feature_list_of_lists[index])
#
#         result = sheets_get(spreadsheet_id, range_name)