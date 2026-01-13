"""Create sections for rooms in selected phase and place them on A2 sheets with optimized packing"""

__title__ = "Room Sections\nto Sheets"
__author__ = "Luca Rosati"

import clr
import System
from System.Collections.Generic import List

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import *

from pyrevit import forms, script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


def ParaInst(element, paraname):
    """Read parameter value from element (VisionXtools standard)"""
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
    else:
        value = "N/A"

    return value


def create_section_for_room(room, section_type, counter, active_view, orientation='vertical'):
    """
    Create a single section view for a room.

    Args:
        room: Room element
        section_type: ViewFamilyType for section
        counter: Progressive counter
        active_view: Active view for bounding box
        orientation: 'vertical' (major axis) or 'horizontal' (minor axis)

    Returns:
        ViewSection or None if creation fails
    """
    # Get room properties
    room_name = ParaInst(room, "Name")
    room_number = ParaInst(room, "Number")

    # Handle empty names
    if not room_name or room_name == "N/A" or room_name == "":
        room_name = "Room"
    if not room_number or room_number == "N/A" or room_number == "":
        room_number = str(counter)

    # Get bounding box
    bbox = room.get_BoundingBox(active_view)
    if bbox is None:
        return None

    # Calculate room dimensions
    center = (bbox.Max + bbox.Min) / 2.0
    width_x = bbox.Max.X - bbox.Min.X
    width_y = bbox.Max.Y - bbox.Min.Y
    height_z = bbox.Max.Z - bbox.Min.Z

    print("DEBUG SECTION: Room '{}' - Center: ({}, {}, {})".format(room_name, center.X, center.Y, center.Z))
    print("DEBUG SECTION: Room dimensions - X: {} ft, Y: {} ft, Z: {} ft".format(width_x, width_y, height_z))

    # Determine major and minor axes
    if width_x >= width_y:
        major_width = width_x
        minor_width = width_y
        major_is_x = True
    else:
        major_width = width_y
        minor_width = width_x
        major_is_x = False

    # Offset for section depth - convert 500mm to project units
    # Get project units
    try:
        units = doc.GetUnits().GetFormatOptions(UnitType.UT_Length).DisplayUnits
        if units == DisplayUnitType.DUT_MILLIMETERS:
            offset = 500.0 / 304.8  # mm to feet
        elif units == DisplayUnitType.DUT_METERS:
            offset = 0.5 / 0.3048  # meters to feet
        else:
            offset = 1.64  # Default to feet
    except:
        units = doc.GetUnits().GetFormatOptions(SpecTypeId.Length).GetUnitTypeId()
        if units == UnitTypeId.Millimeters:
            offset = 500.0 / 304.8
        elif units == UnitTypeId.Meters:
            offset = 0.5 / 0.3048
        else:
            offset = 1.64

    print("DEBUG SECTION: Orientation: {}, Major axis: {}, Offset: {} ft".format(orientation, "X" if major_is_x else "Y", offset))

    # Set up transform based on orientation
    # For a section view:
    # - BasisX = right direction (along section width)
    # - BasisY = up direction (always vertical)
    # - BasisZ = view direction (perpendicular to section) = BasisX cross BasisY

    up = XYZ.BasisZ  # Always up

    if orientation == 'vertical':
        # Section along major axis
        if major_is_x:
            # Section runs along X axis
            right = XYZ.BasisX
            section_width = major_width
            section_depth = minor_width  # Depth perpendicular to section
        else:
            # Section runs along Y axis
            right = XYZ.BasisY
            section_width = major_width
            section_depth = minor_width  # Depth perpendicular to section

        suffix = "V"
    else:
        # Section along minor axis (horizontal)
        if major_is_x:
            # Section runs along Y (minor)
            right = XYZ.BasisY
            section_width = minor_width
            section_depth = major_width  # Depth perpendicular to section
        else:
            # Section runs along X (minor)
            right = XYZ.BasisX
            section_width = minor_width
            section_depth = major_width  # Depth perpendicular to section

        suffix = "H"

    # Calculate view direction as cross product to ensure right-handed system
    # BasisZ = BasisX cross BasisY
    view_dir = right.CrossProduct(up)

    # Create transform with proper right-handed coordinate system
    t = Transform.Identity
    t.Origin = center
    t.BasisX = right       # Along section width
    t.BasisY = up          # Up direction
    t.BasisZ = view_dir    # View direction (computed from cross product)

    # Create bounding box for section
    # Section passes through center with dynamic depth
    # Far Clip = perpendicular room dimension + offset (to see borders)
    # BoundingBox is in local coordinates (Transform.Origin = center)
    bbox_section = BoundingBoxXYZ()
    bbox_section.Transform = t
    bbox_section.Min = XYZ(-section_width/2 - offset, -height_z/2 - offset, 0)
    bbox_section.Max = XYZ(section_width/2 + offset, height_z/2 + offset, section_depth + offset)

    print("DEBUG BBOX: Width: {} ft, Height: {} ft (from {} to {}), Depth: {} ft".format(
        section_width + 2*offset, height_z + 2*offset, -height_z/2 - offset, height_z/2 + offset, section_depth + offset))

    # Create section view
    try:
        section = ViewSection.CreateSection(doc, section_type.Id, bbox_section)
        section.Scale = 50  # 1:50 scale
        # Clean name (remove special characters that might cause issues)
        clean_name = "{}_{}_{}_{}".format(
            room_name.replace(" ", "_"),
            room_number.replace(" ", "_"),
            counter,
            suffix
        )
        section.Name = clean_name
        return section
    except Exception as e:
        return None


def pack_viewports_on_sheet(viewport_data, available_width, available_height, margin):
    """
    Simple row-based packing algorithm for viewports.

    Args:
        viewport_data: List of dicts with 'view', 'width', 'height', 'placed' keys
        available_width: Available width on sheet
        available_height: Available height on sheet
        margin: Margin from sheet edges

    Returns:
        List of placement dicts with 'viewport', 'x', 'y' keys
    """
    placements = []
    current_x = margin
    current_y = margin
    row_height = 0
    gap = 0.033  # 10mm gap between viewports

    print("DEBUG PACKING: Starting - avail_width: {}, avail_height: {}".format(available_width, available_height))

    for vp in viewport_data:
        if vp['placed']:
            continue

        vp_w = vp['width']
        vp_h = vp['height']

        print("DEBUG PACKING: Viewport {} x {} at pos ({}, {})".format(vp_w, vp_h, current_x, current_y))

        # Check if fits in current row
        if current_x + vp_w > available_width + margin:
            print("DEBUG PACKING: Moving to next row")
            # Move to next row
            current_x = margin
            current_y += row_height + gap
            row_height = 0

        # Check if fits vertically
        if current_y + vp_h > available_height + margin:
            print("DEBUG PACKING: Sheet full - y:{} + h:{} > limit:{}".format(current_y, vp_h, available_height + margin))
            # Sheet is full
            break

        # Place viewport
        print("DEBUG PACKING: Placed OK")
        placements.append({
            'viewport': vp,
            'x': current_x + vp_w / 2,  # Center point
            'y': current_y + vp_h / 2
        })
        vp['placed'] = True

        # Update position for next viewport
        current_x += vp_w + gap
        row_height = max(row_height, vp_h)

    print("DEBUG PACKING: Placed {} viewports".format(len(placements)))
    return placements


# ===== PHASE 1: PHASE SELECTION =====
output = script.get_output()

# Extract all phases from document
phase_collector = FilteredElementCollector(doc).OfClass(Phase)
phase_list = list(phase_collector.ToElements())

if len(phase_list) == 0:
    forms.alert('No phases found in document', exitscript=True)

# Extract phase names for dropdown
phase_names = [p.Name for p in phase_list]

# Set default to "Stato di Progetto" if exists, otherwise first phase
default_phase = "Stato di Progetto" if "Stato di Progetto" in phase_names else phase_names[0]

selected_phase_name = forms.ask_for_one_item(
    phase_names,
    default=default_phase,
    prompt='Select Phase for Rooms',
    title='Room Section Creator'
)

if selected_phase_name is None:
    forms.alert('No phase selected', exitscript=True)

# Find corresponding Phase object
target_phase = None
for p in phase_list:
    if p.Name == selected_phase_name:
        target_phase = p
        break

if target_phase is None:
    forms.alert('Phase not found', exitscript=True)

# ===== PHASE 2: ROOM COLLECTION =====
# Collect ALL rooms in document
all_rooms = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_Rooms)\
    .WhereElementIsNotElementType()\
    .ToElements()

# Filter rooms by selected phase and area > 0
rooms = []
for room in all_rooms:
    # Get room phase parameter
    phase_param = room.get_Parameter(BuiltInParameter.ROOM_PHASE)
    if phase_param and phase_param.AsElementId() == target_phase.Id:
        # Filter out rooms with zero area
        if room.Area > 0:
            rooms.append(room)

# Validate we found rooms
if len(rooms) == 0:
    forms.alert(
        "No rooms found in phase '{}'\n\nPlease verify that rooms exist in this phase.".format(selected_phase_name),
        exitscript=True
    )

# ===== PHASE 3: SECTION TYPE SELECTION =====
view_types = FilteredElementCollector(doc)\
    .OfClass(ViewFamilyType)\
    .WhereElementIsElementType()\
    .ToElements()

section_types = []
section_names = []
for vt in view_types:
    if vt.ViewFamily == ViewFamily.Section:
        section_types.append(vt)
        section_names.append(vt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString())

if len(section_types) == 0:
    forms.alert('No section view types found in document', exitscript=True)

selected_section_name = forms.ask_for_one_item(
    section_names,
    default=section_names[0],
    prompt='Select Section View Type',
    title='Room Section Creator'
)

if selected_section_name is None:
    forms.alert('No section type selected', exitscript=True)

# Find selected section type
section_type = None
for name, vt in zip(section_names, section_types):
    if name == selected_section_name:
        section_type = vt
        break

# ===== PHASE 4: CREATE SECTIONS =====
created_sections = []
section_errors = []
active_view = doc.ActiveView

t = Transaction(doc, "Create Room Sections")
t.Start()

try:
    for counter, room in enumerate(rooms, start=1):
        # Create vertical section (along major axis)
        section_v = create_section_for_room(room, section_type, counter, active_view, 'vertical')
        if section_v:
            created_sections.append(section_v)
        else:
            room_name = ParaInst(room, "Name")
            section_errors.append("Room '{}' - Failed to create vertical section".format(room_name))

        # Create horizontal section (along minor axis)
        section_h = create_section_for_room(room, section_type, counter, active_view, 'horizontal')
        if section_h:
            created_sections.append(section_h)
        else:
            room_name = ParaInst(room, "Name")
            section_errors.append("Room '{}' - Failed to create horizontal section".format(room_name))
except Exception as e:
    t.RollBack()
    import traceback
    print("ERROR DETAILS:")
    print(str(e))
    print(traceback.format_exc())
    forms.alert('Error creating sections:\n\n{}'.format(str(e)), exitscript=True)
else:
    t.Commit()

if len(created_sections) == 0:
    forms.alert('No sections could be created', exitscript=True)

# ===== PHASE 5: TITLE BLOCK SELECTION =====
# Collect ALL title block family symbols (no filtering by size)
title_blocks = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_TitleBlocks)\
    .WhereElementIsElementType()\
    .ToElements()

if len(title_blocks) == 0:
    forms.alert('No title blocks found in project.\n\nPlease load a title block family.', exitscript=True)

# Build list of title blocks with dimensions for user selection
tb_list = []
tb_names = []
for tb in title_blocks:
    # Try to get sheet width and height parameters by name
    try:
        # Use LookupParameter with string names instead of BuiltInParameter
        width_param = tb.LookupParameter("Sheet Width")
        height_param = tb.LookupParameter("Sheet Height")

        # Get family and symbol names
        family_name = tb.FamilyName if hasattr(tb, 'FamilyName') else 'Unknown'
        symbol_param = tb.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        symbol_name = symbol_param.AsString() if symbol_param else 'Unknown'

        # Try to get dimensions
        if width_param and height_param:
            try:
                width = width_param.AsDouble()
                height = height_param.AsDouble()

                # Convert to mm for display (feet to mm: * 304.8)
                width_mm = int(width * 304.8)
                height_mm = int(height * 304.8)

                # Show with dimensions
                display_name = "{} : {} ({} x {} mm)".format(family_name, symbol_name, width_mm, height_mm)
            except:
                # Dimensions not available, show without
                display_name = "{} : {}".format(family_name, symbol_name)
        else:
            # Parameters not found, show without dimensions
            display_name = "{} : {}".format(family_name, symbol_name)

        tb_list.append(tb)
        tb_names.append(display_name)
    except:
        # Skip this title block if any error occurs
        continue

if len(tb_list) == 0:
    forms.alert('No title blocks found in project.\n\nPlease load a title block family and try again.', exitscript=True)

# Let user select any title block
selected_tb_name = forms.ask_for_one_item(
    tb_names,
    default=tb_names[0],
    prompt='Select Title Block',
    title='Room Section Creator'
)

if selected_tb_name is None:
    forms.alert('No title block selected', exitscript=True)

# Find selected title block
title_block = None
for name, tb in zip(tb_names, tb_list):
    if name == selected_tb_name:
        title_block = tb
        break

# Get title block dimensions and calculate available area
# Access custom Sheet Width/Height parameters from title block type
width_param = title_block.LookupParameter("Sheet Width")
height_param = title_block.LookupParameter("Sheet Height")

if width_param is None or height_param is None:
    forms.alert('Selected title block does not have Sheet Width/Height parameters.\n\nPlease add these parameters to the title block family.', exitscript=True)

# Get dimensions (values are in feet - Revit internal units)
tb_width = width_param.AsDouble()
tb_height = height_param.AsDouble()

# Account for margins (50mm = 0.164ft on each side)
margin = 0.164
# Reserve space for title block at bottom (50mm = 0.164ft)
title_block_height = 0.164

available_width = tb_width - 2 * margin
available_height = tb_height - 2 * margin - title_block_height

# ===== PHASE 6: CALCULATE VIEWPORT SIZES =====
viewport_data = []
print("DEBUG: Calculating viewport sizes...")
print("DEBUG: Title block dimensions - Width: {} ft, Height: {} ft".format(tb_width, tb_height))
print("DEBUG: Available area - Width: {} ft, Height: {} ft".format(available_width, available_height))

for section in created_sections:
    # Get section CropBox to calculate viewport size
    crop_box = section.CropBox
    if crop_box:
        # Calculate model space dimensions
        crop_min = crop_box.Min
        crop_max = crop_box.Max

        # Calculate width and height in model space
        model_width = crop_max.X - crop_min.X
        model_height = crop_max.Y - crop_min.Y

        # Convert to paper space using scale (scale is 1:20, so divide by 20)
        scale = section.Scale
        paper_width = model_width / scale
        paper_height = model_height / scale

        # Add small padding (10mm = 0.033ft)
        padding = 0.033
        vp_width = paper_width + padding
        vp_height = paper_height + padding

        print("DEBUG: Section '{}' - Model: {}x{} ft, Scale: 1:{}, Paper: {}x{} ft".format(
            section.Name, model_width, model_height, scale, vp_width, vp_height))

        viewport_data.append({
            'view': section,
            'width': vp_width,
            'height': vp_height,
            'placed': False
        })
    else:
        print("DEBUG: Section '{}' has no CropBox!".format(section.Name))

print("DEBUG: Total viewports to place: {}".format(len(viewport_data)))

# ===== PHASE 7: CREATE SHEETS AND PLACE VIEWPORTS =====
created_sheets = []
sheet_counter = 1
viewport_errors = []

t2 = Transaction(doc, "Create Sheets and Place Viewports")
t2.Start()

try:
    while not all(vp['placed'] for vp in viewport_data):
        # Pack viewports for current sheet
        placements = pack_viewports_on_sheet(viewport_data, available_width, available_height, margin)

        if len(placements) == 0:
            # No more viewports can be placed
            break

        # Create sheet
        sheet = ViewSheet.Create(doc, title_block.Id)
        sheet.SheetNumber = "RS-{:03d}".format(sheet_counter)
        sheet.Name = "Room Sections - Sheet {}".format(sheet_counter)
        created_sheets.append(sheet)

        # Place viewports on sheet
        for placement in placements:
            vp_view = placement['viewport']['view']
            vp_x = placement['x']
            vp_y = placement['y']

            # Create viewport
            location = XYZ(vp_x, vp_y, 0)
            try:
                viewport = Viewport.Create(doc, sheet.Id, vp_view.Id, location)
            except Exception as e:
                viewport_errors.append("Failed to place viewport for view '{}'".format(vp_view.Name))

        sheet_counter += 1
except:
    t2.RollBack()
    forms.alert('Error creating sheets or placing viewports', exitscript=True)
else:
    t2.Commit()

# ===== PHASE 8: GENERATE REPORT =====
output.set_height(700)
output.print_md('# Room Section Creation Report')
output.print_md('---')
output.print_md('## Summary')
output.print_md('- **Phase:** {}'.format(selected_phase_name))
output.print_md('- **Rooms Processed:** {}'.format(len(rooms)))
output.print_md('- **Sections Created:** {}'.format(len(created_sections)))
output.print_md('- **Sheets Created:** {}'.format(len(created_sheets)))
output.print_md('---')

if len(created_sections) > 0:
    output.print_md('## Created Sections (Scale 1:20)')
    for section in created_sections:
        output.print_md('- {}'.format(section.Name))
    output.print_md('---')

if len(created_sheets) > 0:
    output.print_md('## Created Sheets')
    for sheet in created_sheets:
        output.print_md('- {} - {}'.format(sheet.SheetNumber, sheet.Name))
    output.print_md('---')

# Show errors if any
if len(section_errors) > 0 or len(viewport_errors) > 0:
    output.print_md('## Errors Encountered')
    for error in section_errors:
        output.print_md('- {}'.format(error))
    for error in viewport_errors:
        output.print_md('- {}'.format(error))
    output.print_md('---')

output.print_md('**Operation completed successfully**')
