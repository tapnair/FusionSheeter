import traceback

import adsk.core
import adsk.fusion

handlers = []


# Returns a dictionary for all inputs. Very useful for creating quick Fusion 360 Add-ins
def get_inputs(command_inputs):
    value_types = [adsk.core.BoolValueCommandInput.classType(), adsk.core.DistanceValueCommandInput.classType(),
                   adsk.core.FloatSliderCommandInput.classType(), adsk.core.FloatSpinnerCommandInput.classType(),
                   adsk.core.IntegerSliderCommandInput.classType(), adsk.core.IntegerSpinnerCommandInput.classType(),
                   adsk.core.ValueCommandInput.classType(), adsk.core.SliderCommandInput.classType(),
                   adsk.core.StringValueCommandInput.classType()]

    list_types = [adsk.core.ButtonRowCommandInput.classType(), adsk.core.DropDownCommandInput.classType(),
                  adsk.core.RadioButtonGroupCommandInput.classType()]

    selection_types = [adsk.core.SelectionCommandInput.classType()]

    input_values = {}
    input_values.clear()

    for command_input in command_inputs:

        # If the input type is in this list the value of the input is returned
        if command_input.objectType in value_types:
            input_values[command_input.id] = command_input.value
            input_values[command_input.id + '_input'] = command_input

        # TODO need to account for radio and button multi select also
        # If the input type is in this list the name of the selected list item is returned
        elif command_input.objectType in list_types:
            if command_input.objectType == adsk.core.DropDownCommandInput.classType():
                if command_input.dropDownStyle == adsk.core.DropDownStyles.CheckBoxDropDownStyle:
                    input_values[command_input.id] = command_input.listItems
                    input_values[command_input.id + '_input'] = command_input

                else:
                    if command_input.selectedItem is not None:
                        input_values[command_input.id] = command_input.selectedItem.name
                        input_values[command_input.id + '_input'] = command_input
            else:
                if command_input.selectedItem is not None:
                    input_values[command_input.id] = command_input.selectedItem.name
                    input_values[command_input.id + '_input'] = command_input

        # If the input type is a selection an array of entities is returned
        elif command_input.objectType in selection_types:
            if command_input.selectionCount > 0:
                selections = []
                for i in range(0, command_input.selectionCount):
                    selections.append(command_input.selection(i).entity)

                input_values[command_input.id] = selections
                input_values[command_input.id + '_input'] = command_input

        else:
            input_values[command_input.id] = command_input.name
            input_values[command_input.id + '_input'] = command_input

    return input_values


# Finds command definition in active UI
def command_definition_by_id(cmd_id, ui):

    command_definitions = ui.commandDefinitions
    command_definition = command_definitions.itemById(cmd_id)
    return command_definition


# Find command control by id in nav bar
def cmd_control_in_nav_bar(cmd_id, ui):

    toolbars_ = ui.toolbars
    nav_toolbar = toolbars_.itemById('NavToolbar')
    nav_toolbar_controls = nav_toolbar.controls
    cmd_control = nav_toolbar_controls.itemById(cmd_id)

    if cmd_control is not None:
        return cmd_control


# Destroys a given object
def destroy_object(obj_to_be_deleted):
    app = adsk.core.Application.cast(adsk.core.Application.get())
    ui = app.userInterface

    if ui and obj_to_be_deleted:
        if obj_to_be_deleted.isValid:
            obj_to_be_deleted.deleteMe()
        else:
            ui.messageBox(obj_to_be_deleted.id + 'is not a valid object')


# Returns the id of a Toolbar Panel in the given Workspace
def toolbar_panel_by_id_in_workspace(workspace_id, toolbar_panel_id):
    app = adsk.core.Application.cast(adsk.core.Application.get())
    ui = app.userInterface

    all_workspaces = ui.workspaces
    this_workspace = all_workspaces.itemById(workspace_id)

    if this_workspace is None:
        ui.messageBox(toolbar_panel_id + 'is not a valid workspace')
        raise ValueError

    all_toolbar_panels = this_workspace.toolbarPanels
    toolbar_panel = all_toolbar_panels.itemById(toolbar_panel_id)

    if toolbar_panel is None:

        toolbar_panel = all_toolbar_panels.add(toolbar_panel_id, toolbar_panel_id)

    return toolbar_panel

        # ui.messageBox(toolbar_panel_id + 'is not a valid tool bar')
        # raise ValueError


# Returns the Command Control from the given panel
def command_control_by_id_in_panel(cmd_id, toolbar_panel, ui):
    if not cmd_id:
        ui.messageBox('Command Control:  ' + cmd_id + '  is not specified')
        return None

    cmd_control = toolbar_panel.controls.itemById(cmd_id)

    if cmd_control is not None:
        return cmd_control

    else:
        raise ValueError


# Get Controls in workspace panel or nav bar
def get_controls(command_in_nav_bar, workspace, toolbar_panel_id, ui):

    # Add command in nav bar
    if command_in_nav_bar:

        toolbars_ = ui.toolbars
        nav_bar = toolbars_.itemById('NavToolbar')
        controls = nav_bar.controls

    # Get Controls from a workspace panel
    else:
        toolbar_panel = toolbar_panel_by_id_in_workspace(workspace, toolbar_panel_id)
        controls = toolbar_panel.controls

    if controls is not None:
        return controls
    else:
        raise RuntimeError


# Base Class for creating Fusion 360 Commands
class Fusion360CommandBase:
    def __init__(self, cmd_def, debug):

        self.cmd_name = cmd_def.get('cmd_name', 'Default Command Name')
        self.cmd_description = cmd_def.get('cmd_description', 'Default Command Description')
        self.cmd_resources = cmd_def.get('cmd_resources', './resources')
        self.cmd_id = cmd_def.get('cmd_id', 'Default Command ID')

        self.workspace = cmd_def.get('workspace', 'FusionSolidEnvironment')
        self.toolbar_panel_id = cmd_def.get('toolbar_panel_id', 'SolidScriptsAddinsPanel')

        self.add_to_drop_down = cmd_def.get('add_to_drop_down', False)
        self.drop_down_cmd_id = cmd_def.get('drop_down_cmd_id', 'Default_DC_CmdId')
        self.drop_down_resources = cmd_def.get('drop_down_resources', './resources')
        self.drop_down_name = cmd_def.get('drop_down_name', 'Drop Name')

        self.command_in_nav_bar = cmd_def.get('command_in_nav_bar', False)

        self.command_visible = cmd_def.get('command_visible', True)

        self.command_promoted = cmd_def.get('command_promoted', False)

        self.debug = debug

        # global set of event handlers to keep them referenced for the duration of the command
        self.handlers = []

    def on_preview(self, command, inputs, args, input_values):
        pass

    def on_destroy(self, command, inputs, reason, input_values):
        pass

    def on_input_changed(self, command_, command_inputs, changed_input, input_values):
        pass

    def on_execute(self, command, inputs, args, input_values):
        pass

    def on_create(self, command, inputs):
        pass

    def on_run(self):
        global handlers

        app = adsk.core.Application.cast(adsk.core.Application.get())
        ui = app.userInterface

        try:

            cmd_definitions = ui.commandDefinitions

            controls_to_add_to = get_controls(self.command_in_nav_bar, self.workspace, self.toolbar_panel_id, ui)

            # Add to a drop down
            if self.add_to_drop_down:

                drop_control = controls_to_add_to.itemById(self.drop_down_cmd_id)

                if not drop_control:
                    drop_control = controls_to_add_to.addDropDown(self.drop_down_name, self.drop_down_resources,
                                                                  self.drop_down_cmd_id)

                controls_to_add_to = drop_control.controls

            new_control = controls_to_add_to.itemById(self.cmd_id)

            # If control does not exist, create it
            if not new_control:
                cmd_definition = cmd_definitions.itemById(self.cmd_id)
                if not cmd_definition:
                    cmd_definition = cmd_definitions.addButtonDefinition(self.cmd_id,
                                                                         self.cmd_name,
                                                                         self.cmd_description,
                                                                         self.cmd_resources)

                on_command_created_handler = CommandCreatedEventHandler(self)
                cmd_definition.commandCreated.add(on_command_created_handler)
                handlers.append(on_command_created_handler)

                new_control = controls_to_add_to.addCommand(cmd_definition)

                if self.command_visible:
                    new_control.isVisible = True
                else:
                    new_control.isVisible = False

                if self.command_promoted:
                    new_control.isPromoted = True
                else:
                    new_control.isPromoted = False


        except:
            if ui:
                ui.messageBox('AddIn Start Failed: {}'.format(traceback.format_exc()))

    def on_stop(self):
        app = adsk.core.Application.cast(adsk.core.Application.get())
        ui = app.userInterface

        try:

            controls_to_delete_from = get_controls(self.command_in_nav_bar, self.workspace, self.toolbar_panel_id, ui)

            # If it is in a drop down
            if self.add_to_drop_down:
                drop_down_control = controls_to_delete_from.itemById(self.drop_down_cmd_id)
                controls_to_delete_from = drop_down_control.controls

            cmd_control = controls_to_delete_from.itemById(self.cmd_id)
            cmd_definition = command_definition_by_id(self.cmd_id, ui)

            destroy_object(cmd_control)
            destroy_object(cmd_definition)

            if self.add_to_drop_down:
                if drop_down_control.controls.count == 0:
                    drop_down_definition = command_definition_by_id(self.drop_down_cmd_id, ui)

                    destroy_object(drop_down_control)
                    destroy_object(drop_down_definition)

        except:
            if ui:
                ui.messageBox('AddIn Stop Failed: {}'.format(traceback.format_exc()))


class ExecutePreviewHandler(adsk.core.CommandEventHandler):
    def __init__(self, cmd_object):
        super().__init__()
        self.cmd_object_ = cmd_object
        self.args = None

    def notify(self, args):
        app = adsk.core.Application.cast(adsk.core.Application.get())
        ui = app.userInterface

        try:

            command_ = args.firingEvent.sender
            command_inputs = command_.commandInputs
            if self.cmd_object_.debug:
                ui.messageBox('***Debug *** Preview: {} execute preview event triggered'.
                              format(command_.parentCommandDefinition.id))
            input_values = get_inputs(command_inputs)
            self.cmd_object_.on_preview(command_, command_inputs, args, input_values)

        except:
            if ui:
                ui.messageBox('Input changed event failed: {}'.format(traceback.format_exc()))


class DestroyHandler(adsk.core.CommandEventHandler):
    def __init__(self, cmd_object):
        super().__init__()
        self.cmd_object_ = cmd_object

    def notify(self, args):
        app = adsk.core.Application.cast(adsk.core.Application.get())
        ui = app.userInterface

        try:
            command_ = args.firingEvent.sender
            command_inputs = command_.commandInputs
            reason_ = args.terminationReason

            if self.cmd_object_.debug:
                ui.messageBox('***Debug ***Command: {} destroyed'.format(command_.parentCommandDefinition.id))
                ui.messageBox("***Debug ***Reason for termination= " + str(reason_))

            input_values = get_inputs(command_inputs)

            self.cmd_object_.on_destroy(command_, command_inputs, reason_, input_values)

        except:
            if ui:
                ui.messageBox('Input changed event failed: {}'.format(traceback.format_exc()))


class InputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self, cmd_object):
        super().__init__()
        self.cmd_object_ = cmd_object

    def notify(self, args):
        app = adsk.core.Application.cast(adsk.core.Application.get())
        ui = app.userInterface

        try:
            command_ = args.firingEvent.sender
            command_inputs = command_.commandInputs
            changed_input = args.input

            if self.cmd_object_.debug:
                ui.messageBox('***Debug Input: {} changed event triggered'.format(command_.parentCommandDefinition.id))
                ui.messageBox('***Debug The Input: {} was the command'.format(changed_input.id))

            input_values = get_inputs(command_inputs)

            self.cmd_object_.on_input_changed(command_, command_inputs, changed_input, input_values)

        except:
            if ui:
                ui.messageBox('Input changed event failed: {}'.format(traceback.format_exc()))


class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self, cmd_object):
        super().__init__()
        self.cmd_object_ = cmd_object

    def notify(self, args):
        try:
            app = adsk.core.Application.cast(adsk.core.Application.get())
            ui = app.userInterface
            command_ = args.firingEvent.sender
            command_inputs = command_.commandInputs

            if self.cmd_object_.debug:
                ui.messageBox('***Debug command: {} executed successfully'.format(command_.parentCommandDefinition.id))

            input_values = get_inputs(command_inputs)

            self.cmd_object_.on_execute(command_, command_inputs, args, input_values)

        except:
            if ui:
                ui.messageBox('command executed failed: {}'.format(traceback.format_exc()))


class CommandCreatedEventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self, cmd_object):
        super().__init__()
        self.cmd_object_ = cmd_object

    def notify(self, args):
        app = adsk.core.Application.cast(adsk.core.Application.get())
        ui = app.userInterface

        try:
            global handlers

            command_ = args.command
            inputs_ = command_.commandInputs

            on_execute_handler = CommandExecuteHandler(self.cmd_object_)
            command_.execute.add(on_execute_handler)
            handlers.append(on_execute_handler)

            on_input_changed_handler = InputChangedHandler(self.cmd_object_)
            command_.inputChanged.add(on_input_changed_handler)
            handlers.append(on_input_changed_handler)

            on_destroy_handler = DestroyHandler(self.cmd_object_)
            command_.destroy.add(on_destroy_handler)
            handlers.append(on_destroy_handler)

            on_execute_preview_handler = ExecutePreviewHandler(self.cmd_object_)
            command_.executePreview.add(on_execute_preview_handler)
            handlers.append(on_execute_preview_handler)

            if self.cmd_object_.debug:
                ui.messageBox('***Debug ***Panel command created successfully')

            self.cmd_object_.on_create(command_, inputs_)

        except:
            if ui:
                ui.messageBox('Command created failed: {}'.format(traceback.format_exc()))