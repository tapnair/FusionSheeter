import os

from .Fusion360Utilities.Fusion360Utilities import get_app_objects


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


def get_display_row_id():
    app_objects = get_app_objects()
    ui = app_objects['ui']
    design = app_objects['design']

    display_row_id_attribute = design.attributes.itemByName('FusionSheeter', 'display_row_index')

    if display_row_id_attribute:
        row_id = display_row_id_attribute.value
    else:
        return 0

    return int(row_id)


# Get default directory
def get_default_app_dir():
    # Get user's home directory
    default_dir = os.path.expanduser("~")

    # Create a subdirectory for this application
    default_dir = os.path.join(default_dir, 'FusionSheeter', '')

    # Create the folder if it does not exist
    if not os.path.exists(default_dir):
        os.makedirs(default_dir)

    return default_dir


# Get default directory
def get_default_model_dir(model_name):

    default_dir = get_default_app_dir()

    model_name = model_name[:model_name.rfind(' v')]

    # Create sub directory for specific model
    default_dir = os.path.join(default_dir, 'Output', model_name, '')

    # Create the folder if it does not exist
    if not os.path.exists(default_dir):
        os.makedirs(default_dir)

    return default_dir


