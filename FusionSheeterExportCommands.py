import adsk.core
import adsk.fusion
import adsk.cam

import os

from .Fusion360Utilities.Fusion360Utilities import get_app_objects
from .Fusion360Utilities.Fusion360CommandBase import Fusion360CommandBase
from .SheeterUtilities import get_sheet_id, get_default_model_dir, get_row_id
from .SheetsService import sheets_get_ranges
from .SheeterModelUtilities import get_parameters2, get_features2, update_local_parameters, update_local_features

name_index_map = {}

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


def dup_check(name):
    if os.path.exists(name):
        base, ext = os.path.splitext(name)
        base += '-dup'
        name = base + ext
        dup_check(name)
    return name


# Command to build all sizes in a sheet
# Todo better control for naming and for destination
class FusionSheeterBuildCommand(Fusion360CommandBase):

    def __init__(self, cmd_def, debug):
        super().__init__(cmd_def, debug)
        self.value_ranges = []

    # Execute model creation
    def on_execute(self, command, inputs, args, input_values):

        app_objects = get_app_objects()
        document = app_objects['document']
        ui = app_objects['ui']
        design = app_objects['design']

        name_field = input_values['name_field']
        # create_shares = input_values['create_shares']

        if not document.isSaved:
            ui.messageBox('Please save document first')
            return

        folder = document.dataFile.parentFolder

        document.save('Auto Saved by Fusion Sheeter')

        spreadsheet_id = get_sheet_id()

        value_ranges = sheets_get_ranges(spreadsheet_id, ['Parameters', 'Features'])

        # TODO if BOM becomes configurable
        # bom_items = get_bom2(value_ranges)
        # update_local_bom(bom_items)

        parameters = get_parameters2(value_ranges)
        features = get_features2(value_ranges)

        quick_quit = 0

        for i, size in enumerate(parameters):

            update_local_parameters(parameters[i])
            update_local_features(features[i])

            design.attributes.add('FusionSheeter', 'parameter_row_index', str(i))

            document.saveAs(size[name_field], folder, 'Auto Generated by Fusion Sheeter', '')

            name_index_map[size[name_field]] = i + 1

            # # TODO Delete me
            # quick_quit += 1
            # if quick_quit > 3:

        # if create_shares:
        #     adsk.doEvents()
        #     ui.commandDefinitions.itemById('cmdID_FusionSheeterPublicShareCommand').execute()

    def on_create(self, command, command_inputs: adsk.core.CommandInputs):
        drop_input_list = command_inputs.addDropDownCommandInput('name_field', 'File Name Parameter',
                                                                 adsk.core.DropDownStyles.TextListDropDownStyle)

        # TODO add custom field option (with dropdown)
        drop_input_list = drop_input_list.listItems
        drop_input_list.add('Description', True)
        drop_input_list.add('Part Number', True)

        # command_inputs.addBoolValueInput('create_shares', "Create Public Share Links?", True, '', False)

        # TODO Ability to select which rows


# Command to build all sizes in a sheet
# Todo better control for naming and for destination
class FusionSheeterExportCommand(Fusion360CommandBase):

    # Execute model creation
    def on_execute(self, command, inputs, args, input_values):
        folder = input_values['output_folder']
        file_types = input_values['file_types_input'].listItems
        write_version = input_values['write_version']

        spreadsheet_id = get_sheet_id()
        row_id = get_row_id()

        value_ranges = sheets_get_ranges(spreadsheet_id, ['Parameters', 'Features'])

        parameters = get_parameters2(value_ranges)
        features = get_features2(value_ranges)

        for index, size in enumerate(parameters):
            update_local_parameters(parameters[index])
            update_local_features(features[index])

            # Todo handle naming
            export_active_doc(folder, file_types, write_version, index)

        # Revert the model back to last saved size
        update_local_parameters(parameters[row_id])
        update_local_features(features[row_id])

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


