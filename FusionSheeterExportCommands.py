import adsk.core
import adsk.fusion
import adsk.cam
import traceback

import os

from .Fusion360Utilities.Fusion360Utilities import get_app_objects
from .Fusion360Utilities.Fusion360CommandBase import Fusion360CommandBase
from .SheeterUtilities import get_sheet_id, get_default_model_dir, get_row_id
from .SheetsService import sheets_get_ranges
from .SheeterModelUtilities import get_parameters2, get_features2, update_local_parameters, update_local_features


def export_active_doc(folder, file_types, write_version, output_name):
    app = adsk.core.Application.get()
    design = app.activeProduct
    export_mgr = design.exportManager

    export_functions = [
        export_mgr.createIGESExportOptions,
        export_mgr.createSTEPExportOptions,
        export_mgr.createSATExportOptions,
        export_mgr.createSMTExportOptions,
        export_mgr.createFusionArchiveExportOptions,
        export_mgr.createSTLExportOptions
    ]
    export_extensions = ['.igs', '.step', '.sat', '.smt', '.f3d', '.stl']

    for i in range(file_types.count-1):

        if file_types.item(i).isSelected:
            export_name = folder + output_name + export_extensions[i]
            export_name = dup_check(export_name)
            export_options = export_functions[i](export_name)
            export_mgr.execute(export_options)

    if file_types.item(file_types.count - 1).isSelected:
        stl_export_name = folder + output_name + '.stl'
        stl_options = export_mgr.createSTLExportOptions(design.rootComponent, stl_export_name)
        export_mgr.execute(stl_options)


def dup_check(name):
    if os.path.exists(name):
        base, ext = os.path.splitext(name)
        base += '-dup'
        name = base + ext
        dup_check(name)
    return name


def get_name(write_version, index, size, option, column_name):
    app = adsk.core.Application.get()
    design = app.activeProduct
    output_name = ''

    if option == 'Document Name':

        doc_name = app.activeDocument.name

        if not write_version:
            doc_name = doc_name[:doc_name.rfind(' v')]

        output_name = doc_name + '_' + str(index)

    elif option == 'Description':
        column_name = option
        output_name = size.get('Description')

    elif option == 'Part Number':
        column_name = option
        output_name = size.get('Part Number')

    elif option == 'Custom':
        output_name = size.get(column_name)

    else:
        raise ValueError('Something strange happened')

    if output_name is None:
        raise AttributeError(
            'There is no column in the sheet with the name {}.  Aborting operation.'.format(column_name))

    elif output_name.isspace() or (len(output_name) < 1):
        raise ValueError('Skipping row {} as there is no value for the column {}.'.format(column_name, index + 2))

    else:
        return output_name


def add_name_inputs(command_inputs):
    name_option_group = command_inputs.addRadioButtonGroupCommandInput('name_option_id', 'File Name Option')
    name_option_group.listItems.add('Document Name', True)
    name_option_group.listItems.add('Description', False)
    name_option_group.listItems.add('Part Number', False)
    name_option_group.listItems.add('Custom', False)
    name_option_group.isVisible = True
    name_option_group.isFullWidth = True

    custom_name_input = command_inputs.addStringValueInput('column_name_id', 'Custom Column:', 'My Column')
    custom_name_input.isVisible = False

    warning_text = '<br> <b>Warning! </b><br><br>' \
                   'Custom Column must be present on Parameters and Features sheets<br><br>' \
                   '<i>Column name is case sensitive</i><br><br>'

    warning_input = command_inputs.addTextBoxCommandInput('name_warning_id', '',
                                                          warning_text, 16, True)
    warning_input.isVisible = False

    version_input = command_inputs.addBoolValueInput('write_version', 'Write versions to output file names?', True)
    version_input.isVisible = False


def update_name_inputs(command_inputs, selection):
    command_inputs.itemById('column_name_id').isVisible = False
    command_inputs.itemById('name_warning_id').isVisible = False
    command_inputs.itemById('write_version').isVisible = False

    if selection == 'Custom':
        command_inputs.itemById('column_name_id').isVisible = True
        command_inputs.itemById('name_warning_id').isVisible = True

    elif selection == 'Document Name':
        command_inputs.itemById('write_version').isVisible = True


# Command to build all sizes in a sheet
# Todo better control for naming and for destination
class FusionSheeterBuildCommand(Fusion360CommandBase):
    # Execute model creation
    def on_execute(self, command, inputs, args, input_values):

        app_objects = get_app_objects()
        document = app_objects['document']
        ui = app_objects['ui']
        design = app_objects['design']

        write_version = input_values['write_version']
        name_option = input_values['name_option_id']
        column_name = input_values['column_name_id']

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

        for i, size in enumerate(parameters):

            try:
                output_name = get_name(write_version, i, size, name_option, column_name)

                update_local_parameters(parameters[i])
                update_local_features(features[i])

                design.attributes.add('FusionSheeter', 'parameter_row_index', str(i))

                document.saveAs(output_name, folder, 'Auto Generated by Fusion Sheeter', '')

            except ValueError as e:
                ui.messageBox(str(e))
            except AttributeError as e:
                ui.messageBox(str(e))
                break

    def on_input_changed(self, command_, command_inputs, changed_input, input_values):
        if changed_input.id == 'name_option_id':
            update_name_inputs(command_inputs, changed_input.selectedItem.name)

    def on_create(self, command, command_inputs):
        add_name_inputs(command_inputs)
        update_name_inputs(command_inputs, 'Document Name')


# Command to build all sizes in a sheet
# Todo better control for naming and for destination
class FusionSheeterExportCommand(Fusion360CommandBase):

    def on_input_changed(self, command_, command_inputs, changed_input, input_values):
        if changed_input.id == 'name_option_id':
            update_name_inputs(command_inputs, changed_input.selectedItem.name)

    # Execute model creation
    def on_execute(self, command, inputs, args, input_values):
        app_objects = get_app_objects()
        ui = app_objects['ui']

        folder = input_values['output_folder']
        file_types = input_values['file_types_input'].listItems
        write_version = input_values['write_version']
        name_option = input_values['name_option_id']
        column_name = input_values['column_name_id']

        spreadsheet_id = get_sheet_id()
        row_id = get_row_id()

        value_ranges = sheets_get_ranges(spreadsheet_id, ['Parameters', 'Features'])

        parameters = get_parameters2(value_ranges)
        features = get_features2(value_ranges)

        for index, size in enumerate(parameters):

            try:
                update_local_parameters(parameters[index])
                update_local_features(features[index])
                output_name = get_name(write_version, index, size, name_option, column_name)
                export_active_doc(folder, file_types, write_version, output_name)

            except ValueError as e:
                ui.messageBox(str(e))

            except AttributeError as e:
                ui.messageBox(str(e))
                break
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
        drop_input_list.add('STL', False)

        add_name_inputs(command_inputs)
        update_name_inputs(command_inputs, 'Document Name')
