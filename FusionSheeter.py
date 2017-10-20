# Importing sample Fusion Command
# Could import multiple Command definitions here
from .FusionSheeterCommand import FusionSheeterSizeCommand, FusionSheeterCreateCommand, \
    FusionSheeterBuildCommand, FusionSheeterSyncCommand, FusionSheeterOpenSheetCommand, \
    FusionSheeterGCodeCommand, FusionSheeterGCodeCommand2, FusionSheeterGCodeCommand3, FusionSheeterExportCommand, \
    FusionSheeterPaletteCommand, FusionSheeterQuickPullCommand

commands = []
command_definitions = []

# Define parameters for 1st command
cmd = {
    'cmd_name': 'Link Design to Google Sheet',
    'cmd_description': 'Creates a new Google Sheets Document in your Google Drive based on the current active model.  '
                       'Establishes a link from the current design to the new spreadsheet.',
    'cmd_id': 'cmdID_FusionSheeterCreateCommand',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Sheeter',
    'command_promoted': False,
    'class': FusionSheeterCreateCommand
}
command_definitions.append(cmd)

# Define parameters for 1st command
cmd = {
    'cmd_name': 'Sync With Google Sheet',
    'cmd_description': 'Synchronize Design Parameters, BOM Data and Feature Suppression with a Google Sheet',
    'cmd_id': 'cmdID_FusionSheeterSyncCommand',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Sheeter',
    'command_promoted': True,
    'add_separator': True,
    'class': FusionSheeterSyncCommand
}
command_definitions.append(cmd)

# Define parameters for 1st command
cmd = {
    'cmd_name': 'Open Sheet',
    'cmd_description': 'Open Sheet Associated with this Fusion 360 Design',
    'cmd_id': 'cmdID_FusionSheeterOpenSheetCommand',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Sheeter',
    'command_promoted': False,
    'class': FusionSheeterOpenSheetCommand
}
command_definitions.append(cmd)

# Define parameters for 1st command
cmd = {
    'cmd_name': 'Open Sheet in Fusion360',
    'cmd_description': 'Edit the Google Sheet Document in context of Fusion 360',
    'cmd_id': 'cmdID_FusionSheeterPaletteCommand',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Sheeter',
    'command_promoted': False,
    'class': FusionSheeterPaletteCommand
}
command_definitions.append(cmd)

# Define parameters for command
cmd = {
    'cmd_name': 'Change Size',
    'cmd_description': 'Change association of active model to a different row in Sheets Document',
    'cmd_id': 'cmdID_FusionSheeterCommand',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Sheeter',
    'add_separator': True,
    'class': FusionSheeterSizeCommand
}
command_definitions.append(cmd)

# Define parameters for command
cmd = {
    'cmd_name': 'Generate G Code for all sizes',
    'cmd_description': 'Generate NC files for all sizes in associated Sheet',
    'cmd_id': 'cmdID_FusionSheeterGCodeCommand',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Sheeter',
    'command_visible': True,
    'class': FusionSheeterGCodeCommand
}
command_definitions.append(cmd)

# Define parameters for 1st command
cmd = {
    'cmd_name': 'Create Fusion 360 Design for all sizes',
    'cmd_description': 'Create Unique Component for each row in Google Sheet Document in the current project',
    'cmd_id': 'cmdID_FusionSheeterBuildCommand',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Sheeter',
    'command_promoted': False,
    'class': FusionSheeterBuildCommand
}
command_definitions.append(cmd)

# Define parameters for command
cmd = {
    'cmd_name': 'Hidden NC 2',
    'cmd_description': 'Generate NC files for all sizes in associated Sheet',
    'cmd_id': 'cmdID_FusionSheeterGCodeCommand2',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Sheeter',
    'command_visible': False,
    'class': FusionSheeterGCodeCommand2
}
command_definitions.append(cmd)

# Define parameters for command
cmd = {
    'cmd_name': 'Hidden NC 3',
    'cmd_description': 'Generate NC files for all sizes in associated Sheet',
    'cmd_id': 'cmdID_FusionSheeterGCodeCommand3',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Sheeter',
    'command_visible': False,
    'class': FusionSheeterGCodeCommand3
}
command_definitions.append(cmd)

# Define parameters for command
cmd = {
    'cmd_name': 'Export local file for all sizes',
    'cmd_description': 'Generate local export files for all sizes in associated Sheet',
    'cmd_id': 'cmdID_FusionSheeterExportCommand',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Sheeter',
    'command_visible': True,
    'class': FusionSheeterExportCommand
}
command_definitions.append(cmd)




# Define parameters for 1st command
cmd = {
    'cmd_name': 'Quick Pull',
    'cmd_description': 'Sync data from Google Sheets into model',
    'cmd_id': 'cmdID_FusionSheeterQuickPullCommand',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Sheeter',
    'command_in_nav_bar': True,
    'class': FusionSheeterQuickPullCommand
}
command_definitions.append(cmd)
# Set to True to display various useful messages when debugging your app
debug = False

# Don't change anything below here:
for cmd_def in command_definitions:
    command = cmd_def['class'](cmd_def, debug)
    commands.append(command)


def run(context):
    for run_command in commands:
        run_command.on_run()


def stop(context):
    for stop_command in commands:
        stop_command.on_stop()
