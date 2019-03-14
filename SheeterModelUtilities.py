
import adsk.core
import adsk.fusion
import adsk.cam
import traceback

from .Fusion360Utilities.Fusion360Utilities import get_app_objects
from .SheetsService import sheets_update_values


def push_parameters(parameter_matrix, row_id, spreadsheet_id, all_params=False):

    app_objects = get_app_objects()
    um = app_objects['units_manager']

    design = app_objects['design']

    if all_params:
        parameters = design.allParameters
    else:
        parameters = design.userParameters

    for parameter in parameters:

        try:
            index = parameter_matrix[0].index(parameter.name)
            parameter_matrix[row_id][index] = um.formatInternalValue(parameter.value, parameter.unit, False)
        except ValueError:
            parameter_matrix[0].append(parameter.name)
            for i in range(1, len(parameter_matrix)):
                parameter_matrix[i].append(um.formatInternalValue(parameter.value, parameter.unit, False))

    range_body = {"range": "Parameters",
                  "values": parameter_matrix}

    sheet_range = 'Parameters'

    sheets_update_values(spreadsheet_id, sheet_range, range_body)


def push_display(display_matrix, row_id, spreadsheet_id):

    app_objects = get_app_objects()
    design = app_objects['design']

    all_occurrences = design.rootComponent.allOccurrences

    for occurrence in all_occurrences:

        try:
            index = display_matrix[0].index(occurrence.fullPathName)

            display_matrix[row_id][index] = occurrence.isLightBulbOn
        except ValueError:
            display_matrix[0].append(occurrence.fullPathName)
            for i in range(1, len(display_matrix)):
                display_matrix[i].append(occurrence.isLightBulbOn)

    range_body = {"range": "Display",
                  "values": display_matrix}

    sheet_range = 'Display'

    sheets_update_values(spreadsheet_id, sheet_range, range_body)


def find_range(value_ranges, target_range):

    for value_range in value_ranges:
        if target_range in value_range['range']:
            return value_range.get('values', [])

    return None


def get_parameters2(value_ranges):

    rows = find_range(value_ranges, 'Parameters')

    dict_list = []

    for row in rows[1:]:
        row_dict = dict(zip(rows[0], row))
        dict_list.append(row_dict)

    return dict_list


def get_display(value_ranges):

    rows = find_range(value_ranges, 'Display')

    dict_list = []

    for row in rows[1:]:
        row_dict = dict(zip(rows[0], row))
        dict_list.append(row_dict)

    return dict_list


def get_parameters_matrix(value_ranges):

    parameter_matrix = find_range(value_ranges, 'Parameters')

    return parameter_matrix


def get_display_matrix(value_ranges):

    display_matrix = find_range(value_ranges, 'Display')

    return display_matrix


def get_features2(value_ranges):

    rows = find_range(value_ranges, 'Features')

    list_of_lists = []

    for row in rows[1:]:
        row_list = list(zip(rows[0], row))
        list_of_lists.append(row_list)

    return list_of_lists


def update_local_parameters(size):
    app_objects = get_app_objects()
    um = app_objects['units_manager']

    design = app_objects['design']

    if design.designType == adsk.fusion.DesignTypes.ParametricDesignType:

        all_parameters = design.allParameters

        for parameter in all_parameters:

            new_value = size.get(parameter.name)
            if new_value is not None:
                unit_type = parameter.unit

                if len(unit_type) > 0:

                    if um.isValidExpression(new_value, unit_type):
                        sheet_value = um.evaluateExpression(new_value, unit_type)
                    else:
                        continue
                else:
                    sheet_value = float(new_value)

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
def update_local_features(feature_list_input):
    app_objects = get_app_objects()
    design = app_objects['design']

    if design.designType == adsk.fusion.DesignTypes.ParametricDesignType:
        # Record feature suppression state
        time_line = design.timeline

        # Going to iterate time line in reverse
        reverse_feature_list = list(reversed(feature_list_input))

        # Walk time line in reverse order
        for index in reversed(range(time_line.count)):

            time_line_object = time_line.item(index)

            feature_name = get_time_line_object_name(time_line_object)

            new_state = find_list_item(reverse_feature_list, feature_name)

            # Set suppression state from sheet if different than current
            if new_state is not None:

                if new_state[1] == 'Unsuppressed' and time_line_object.isSuppressed:
                    time_line_object.isSuppressed = False

                elif new_state[1] == 'Suppressed' and not time_line_object.isSuppressed:
                    time_line_object.isSuppressed = True

        # Todo create and display change list (not during size change, need variable to control)


# Search sheets feature list for item pops it from list
def find_list_item(feature_list_input, name):
    for index, feature in enumerate(feature_list_input):
        if feature[0] == name:
            return feature_list_input.pop(index)

    return None


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

    try:
        feature_name += time_line_object.name
    except:
        feature_name += 'No Name'

    return feature_name


def update_local_display(display_capture):
    app_objects = get_app_objects()
    design = app_objects['design']

    all_occurrences = design.rootComponent.allOccurrences

    for occurrence in all_occurrences:

        new_value = display_capture.get(occurrence.fullPathName)

        if new_value is not None:
            if new_value == 'TRUE':
                occurrence.isLightBulbOn = True
            elif new_value == 'FALSE':
                occurrence.isLightBulbOn = False