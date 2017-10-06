# Importing sample Fusion Command
# Could import multiple Command definitions here
from .FusionSheeterCommand import FusionSheeterCommand, FusionSheeterCreateCommand

commands = []
command_definitions = []

# Define parameters for 1st command
cmd = {
    'cmd_name': 'Change Size',
    'cmd_description': 'Fusionto Google Sheets Connector',
    'cmd_id': 'cmdID_FusionSheeterCommand',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Sheeter',
    'class': FusionSheeterCommand
}
command_definitions.append(cmd)

# Define parameters for 1st command
cmd = {
    'cmd_name': 'Link Design to Google Sheet',
    'cmd_description': 'Fusion to Google Sheets Creator',
    'cmd_id': 'cmdID_FusionSheeterCreateCommand',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Sheeter',
    'command_promoted': True,
    'class': FusionSheeterCreateCommand
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
