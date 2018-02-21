
from collections import defaultdict, namedtuple
from .SheeterModelUtilities import find_range, update_only_parameters
from .SheetsService import sheets_get_range

BOM_Item = namedtuple('BOM_Item', ('part_number', 'part_name', 'description', 'children', 'level'))


def get_bom2(value_ranges):

    # rows = value_ranges[1].get('values', [])
    rows = find_range(value_ranges, 'BOM')

    dict_list = []

    for row in rows[1:]:
        row_dict = dict(zip(rows[0], row))
        dict_list.append(row_dict)

    return dict_list


def get_bom_adv(value_ranges):

    # rows = value_ranges[1].get('values', [])
    rows = find_range(value_ranges, 'BOM-ADV')

    dict_list = []

    for row in rows[1:]:
        row_dict = dict(zip(rows[0], row))
        dict_list.append(row_dict)

    return dict_list


def get_bom_adv_component(component_name, spreadsheet_id):

    # rows = value_ranges[1].get('values', [])
    value_range = sheets_get_range(spreadsheet_id, 'BOM-' + component_name)

    rows = value_range.get('values', [])
    dict_list = []

    for row in rows[1:]:
        row_dict = dict(zip(rows[0], row))
        dict_list.append(row_dict)

    return dict_list


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
def update_local_bom(items, all_components):

    change_list = ''

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

    return change_list


# Update local BOM metadata based on sheets info
def update_local_bom_adv(selected_components, all_components, spreadsheet_id):

    change_list = ''

    for component in all_components:

        new_number = selected_components.get(component.name)

        if new_number is not None:
            if new_number != component.partNumber:
                component.partNumber = new_number
                change_list += ('Changed: ' + component.name + ' Part Number to: ' + new_number + '\n')

                component_configs = get_bom_adv_component(component.name, spreadsheet_id)

                # Config is row zipped from components sheet
                for config in component_configs:
                    if config['Part Number'] == component.partNumber:
                        component.description = config['Description']
                        # TODO add something here to also read any parameters
                        update_only_parameters(config)

    return change_list
