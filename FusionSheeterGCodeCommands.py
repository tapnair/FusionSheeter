
import adsk.core
import adsk.fusion
import adsk.cam
import traceback

import time
import os

from .Fusion360Utilities.Fusion360Utilities import get_app_objects
from .Fusion360Utilities.Fusion360CommandBase import Fusion360CommandBase
from .SheeterUtilities import get_sheet_id, get_default_model_dir
from .SheeterModelUtilities import get_parameters2, get_features2, update_local_parameters, update_local_features
from .SheetsService import sheets_get_ranges


# Globals for multiple command execution
operation_list = []
post_name = None
output_folder = None
the_file_name = None
gcode_index = 0
gcode_test = 25
params_list = []
feature_list = []

operation_collection = None


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


# Switch Current Workspace
def switch_workspace(workspace_name):

    ao = get_app_objects()
    workspace = ao['ui'].workspaces.itemById(workspace_name)
    workspace.activate()


def gcode_regen_tool_paths(operation_collection_):
    app = adsk.core.Application.get()
    doc = app.activeDocument
    ui = app.userInterface
    products = doc.products
    product = products.itemByProductType('CAMProductType')
    cam = adsk.cam.CAM.cast(product)

    adsk.doEvents()

    # Update tool path
    future = cam.generateToolpath(operation_collection_)
    check = 0
    while not future.isGenerationCompleted:
        adsk.doEvents()

        time.sleep(.1)
        check += 1
        if check > 1000:
            ui.messageBox('Timeout')
            break


def check_valid_container(container):

    valid = False

    if container.objectType == adsk.cam.Operation.classType():
        if container.operationState == adsk.cam.OperationStates.IsValidOperationState:
            valid = True

    elif container.objectType == adsk.cam.Setup.classType() or container.objectType == adsk.cam.CAMFolder.classType():
        for operation in container.allOperations:
            if operation.operationState == adsk.cam.OperationStates.IsValidOperationState:
                valid = True

    return valid


def build_op_collection(operation_list_):
    # TODO use app objects, probably need to switch workspaces?
    app = adsk.core.Application.get()
    doc = app.activeDocument
    products = doc.products
    product = products.itemByProductType('CAMProductType')
    cam = adsk.cam.CAM.cast(product)

    operation_collection_ = adsk.core.ObjectCollection.create()

    for operation_name in operation_list_:
        for setup in cam.setups:
            if setup.name == operation_name:
                operation_collection_.add(setup)
            else:
                for folder in setup.folders:
                    if folder.name == operation_name:
                        operation_collection_.add(folder)

        for operation in cam.allOperations:
            if operation.name == operation_name:
                operation_collection_.add(operation)

    return operation_collection_


# Updates tool paths and outputs new nc file
def g_code_post_operation(operation, post_name_, output_folder_, prefix):

    # TODO use app objects, probably need to switch workspaces?
    app = adsk.core.Application.get()
    doc = app.activeDocument
    products = doc.products
    product = products.itemByProductType('CAMProductType')
    cam = adsk.cam.CAM.cast(product)

    post_config = os.path.join(cam.genericPostFolder, post_name_)
    units = adsk.cam.PostOutputUnitOptions.DocumentUnitsOutput

    program_name = prefix + '_' + operation.name

    program_name = program_name.replace(" ", "")

    post_input = adsk.cam.PostProcessInput.create(program_name, post_config, output_folder_, units)
    post_input.isOpenInEditor = False

    result = cam.postProcess(operation, post_input)


class FusionSheeterGCodeCommand(Fusion360CommandBase):
    def on_execute(self, command, inputs, args, input_values):

        global operation_list
        global post_name
        global output_folder
        global gcode_index
        global gcode_test
        global params_list
        global feature_list
        global operation_collection

        operation_list.clear()

        # Get the values from the user input
        post_name = input_values['post_name']
        output_folder = input_values['output_folder']

        gcode_add_operations(input_values['setups'], operation_list)
        gcode_add_operations(input_values['folders'], operation_list)
        gcode_add_operations(input_values['operations'], operation_list)

        # Debug testing
        # ao = get_app_objects()
        # ao['ui'].messageBox(str(operation_list))

        operation_collection = build_op_collection(operation_list)

        gcode_index = 0

        spreadsheet_id = get_sheet_id()

        value_ranges = sheets_get_ranges(spreadsheet_id, ['Parameters', 'BOM', 'Features', 'Display'])

        # Populate global list of parameters to send to next command
        params_list = get_parameters2(value_ranges)
        feature_list = get_features2(value_ranges)

        # Build 1st model size in list
        update_local_parameters(params_list[gcode_index])
        update_local_features(feature_list[gcode_index])

        # Global to control iteration count
        gcode_test = len(params_list)

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
        if product is None:
            ui.messageBox('There are no CAM operations in the active document')
            return

        # Cast the CAM product to a CAM object (a subtype of product).
        cam = adsk.cam.CAM.cast(product)

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
        # global operation_list
        global post_name
        global output_folder
        global gcode_index
        global gcode_test
        global params_list
        # global feature_list
        global operation_collection

        switch_workspace('CAMEnvironment')

        # Todo make this configurable
        prefix = params_list[gcode_index]['Part Number'] + '_' + params_list[gcode_index]['Description']

        gcode_regen_tool_paths(operation_collection)

        for operation in operation_collection:

            # check if operation is valid
            if check_valid_container(operation):
                g_code_post_operation(operation, post_name, output_folder, prefix)

        switch_workspace('FusionSolidEnvironment')

        gcode_index += 1

        if gcode_index < gcode_test:
            execute_next_command('cmdID_FusionSheeterGCodeCommand3')


class FusionSheeterGCodeCommand3(Fusion360CommandBase):

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

