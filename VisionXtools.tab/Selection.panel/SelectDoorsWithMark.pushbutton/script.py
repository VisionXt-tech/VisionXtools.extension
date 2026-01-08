"""Select all doors in the current view and list their Mark parameters"""

__title__ = "Select Doors\nWith Mark"
__author__ = "Luca Rosati"

from pyrevit import forms, script
import clr
import System
from System.Collections.Generic import List

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


def ParaInst(element, paraname):
    """Legge parametro da elemento gestendo tutti i tipi di dato"""
    param = element.LookupParameter(paraname)
    if param is None:
        return "N/A"

    if param.StorageType == StorageType.Double:
        value = param.AsDouble()
    elif param.StorageType == StorageType.ElementId:
        value = param.AsElementId()
    elif param.StorageType == StorageType.String:
        value = param.AsString()
    elif param.StorageType == StorageType.Integer:
        value = param.AsInteger()
    elif param.StorageType == None:
        value = "Da Compilare"
    else:
        value = "N/A"

    return value


# Get all doors in current view
doors = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(
    BuiltInCategory.OST_Doors
).WhereElementIsNotElementType().ToElements()

if len(doors) == 0:
    forms.alert('No doors found in current view', exitscript=True)

# Extract door data
door_data = []
door_ids = []

for door in doors:
    door_id = door.Id
    mark_value = ParaInst(door, "Mark")

    # Handle empty or None values
    if mark_value is None or mark_value == "":
        mark_value = "<Empty>"

    door_data.append({
        'id': door_id,
        'mark': mark_value,
        'element': door
    })
    door_ids.append(door_id)

# Select doors in the UI
collection = List[ElementId](door_ids)
uidoc.Selection.SetElementIds(collection)

# Output results
output = script.get_output()
output.set_height(600)
output.print_md('# \U0001F6AA Doors Selected in Current View')
output.print_md('## Total Doors Found: {}'.format(len(doors)))
output.print_md('---')
output.print_md('## Door Marks List:')

# Sort by mark value for better readability
door_data_sorted = sorted(door_data, key=lambda x: str(x['mark']))

for data in door_data_sorted:
    mark_display = data['mark']
    door_id = data['id']

    # Print with linkified element ID
    print('Mark: {} - '.format(mark_display), end='')
    output.linkify(door_id)

output.print_md('---')
output.print_md('\U00002705 **Selection completed successfully**')
