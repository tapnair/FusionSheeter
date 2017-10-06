# Importing sample Fusion Command
# Could import multiple Command definitions here
from .FusionBolterCommand import Demo1Command
from .FusionSheeterCommand import Demo2Command

commands = []
command_definitions = []

# Define parameters for 1st command
cmd = {
    'cmd_name': 'Fusion Demo Command 1',
    'cmd_description': 'Fusion Demo Command 1 Description',
    'cmd_id': 'cmdID_demo44',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'SolidScriptsAddinsPanel',
    'class': Demo1Command
}
command_definitions.append(cmd)

# Define parameters for 2nd command
cmd = {
    'cmd_name': 'Fusion Demo Command 2',
    'cmd_description': 'Fusion Demo Command 2 Description',
    'cmd_id': 'cmdID_demo2',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'SolidScriptsAddinsPanel',
    'class': Demo2Command
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
