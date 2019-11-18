import adsk.core
import adsk.fusion
import adsk.cam

import http.client
import urllib.parse

from .Fusion360Utilities.Fusion360CommandBase import Fusion360CommandBase
from .SheeterModelUtilities import get_properties_matrix, update_custom_properties, push_properties
from .SheeterUtilities import get_sheet_id
from .SheetsService import sheets_get_ranges


def un_shorten_url(url):
    parsed = urllib.parse.urlparse(url)
    h = http.client.HTTPConnection(parsed.netloc)
    resource = parsed.path
    if parsed.query != "":
        resource += "?" + parsed.query
    h.request('HEAD', resource)
    response = h.getresponse()
    if response.status//100 == 3 and response.getheader('Location'):
        # return un_shorten_url(response.getheader('Location')) # changed to process chains of short urls
        return response.getheader('Location')
    else:
        return url


class FusionSheeterPublicShareCommand(Fusion360CommandBase):

    # Execute model creation
    def on_execute(self, command, inputs, args, input_values):

        app = adsk.core.Application.get()

        spreadsheet_id = get_sheet_id()
        value_ranges = sheets_get_ranges(spreadsheet_id, ['Custom Properties'])

        property_matrix = get_properties_matrix(value_ranges)

        # TODO fix folder situation
        # for data_file in app.data.activeProject.rootFolder.dataFiles:
        for data_file in app.activeDocument.dataFile.parentFolder.dataFiles:

            # if data_file.fileExtension == "f3d":
            name = data_file.name
            row_index = next((i for i, r in enumerate(property_matrix) if (r[0] == name) or (r[1] == name)), None)
            if row_index is not None:
                short_public_link = data_file.publicLink
                public_link = un_shorten_url(short_public_link)
                custom_properties = {
                    "short_public_link": short_public_link,
                    "public_link": public_link,
                    "public_link_id": public_link.split("/")[-1],
                    "forge_urn": data_file.id,
                    "forge_id": data_file.id.split(":")[-1]
                }

                update_custom_properties(property_matrix, row_index, custom_properties)

        # app.userInterface.messageBox(str(property_matrix))
        push_properties(property_matrix, spreadsheet_id)
