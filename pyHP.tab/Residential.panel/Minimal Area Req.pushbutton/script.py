__title__ = "Min area"
__doc__ = "Fill in minimal area requirement based on a formatted excel file. This script fill in the value of Minimal Area Requirement based on Room Name (ex. Bedroom) and Unit Type. Please make sure these parameters "

from pyrevit import revit, DB, forms
import xlrd
from rpw.ui.forms import (FlexForm, Label, ComboBox, Separator, Button)
import sys

def convert_to_internal(from_units):
    # convert project units to internal
    d_units = DB.Document.GetUnits(revit.doc).GetFormatOptions(DB.UnitType.UT_Area).DisplayUnits
    converted = DB.UnitUtils.ConvertToInternalUnits(from_units, d_units)
    return converted

coll_rooms = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_Rooms).ToElements()

# take only placed rooms
good_rooms = [r for r in coll_rooms if r.Area != 0]

element_parameter_set = good_rooms[0].Parameters

room_params = [p.Definition.Name for p in element_parameter_set if
               p.StorageType.ToString() == "Double" and p.IsReadOnly == False and p.Definition.Name not in ["Limit Offset", "Base Offset"]]

if not room_params:
    forms.alert(msg="No suitable parameter", \
        sub_msg="There is no suitable parameter to use for Minimal Area Requirement. Please add a parameter 'Area Requirement' of Area Type", \
        ok=True, \
        warn_icon=True, exitscript=True)


# pick excel file and read
path = forms.pick_file(file_ext='xlsx',init_dir="M:\BIM\BIM Manual\Minimal area requirement table")
book = xlrd.open_workbook(path)
worksheet = book.sheet_by_index(0)

# create dictionary with min requirements, of format [unit type][room name] : min area
area_dict = {}
for i in range(1, worksheet.ncols):
    area_dict[worksheet.cell_value(0, i)] = {}
    for j in range(1, worksheet.nrows):
        area_dict[worksheet.cell_value(0, i)][worksheet.cell_value(j, 0)] = worksheet.cell_value(j, i)


# a list of variations for Living / Dining / Kitchen room name
lkd_var = ["LKD",
           "LDK",
           "KLD",
           "KDL",
           "DKL",
           "DLK",
           "Living",
           "Kitchen",
           "Dining",
           "L/K/D",
           "L/D/K",
           "K/L/D",
           "K/D/L",
           "D/K/L",
           "D/L/K",
           "L-K-D",
           "L-D-K",
           "K-L-D",
           "K-D-L",
           "D-K-L",
           "D-L-K",
           ]



# check there's a Unit Type parameter
# gather and organize Room parameters: (only editable text params)
room_parameter_set = good_rooms[0].Parameters
room_params_text = [p.Definition.Name for p in room_parameter_set if
                        p.StorageType.ToString() == "String" and p.IsReadOnly == False]


#forms.select_parameters(src_element=good_rooms[0], multiple = False, include_instance = True, include_type = False)
room_params.sort()
# construct rwp UI
components = [
    Label("Which Room parameter is used for Unit Type:"),
    ComboBox(name="unit_type_param", options=room_params_text),
    Label("Select Room parameter to populate"),
    ComboBox("area_req_param", room_params),
    Button("Select")]
form = FlexForm("Unit Type", components)
form.show()
# assign chosen parameters
chosen_room_param1 = form.values["unit_type_param"]
selected_parameter = form.values["area_req_param"]

if not chosen_room_param1:
    sys.exit()



#components = [Label("Select room parameter to populate"), ComboBox("room_tx_params", room_params), Button ("Select")]
#form = FlexForm("Select Parameter", components)
#form.show()
#selected_parameter = form.values["room_tx_params"]

if selected_parameter:
    with revit.Transaction("Write Parameter", revit.doc):
        for room in good_rooms:
            # get room parameters
            area_req = room.LookupParameter(selected_parameter)
            unit_type = room.LookupParameter(chosen_room_param1).AsString()
            if not unit_type:
                unit_type = "Blank"
            room_name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
            # check if Living/Kitchen/Dining is written differently
            if room_name.split()[0] in lkd_var or room_name.split("/")[0] in lkd_var:
                room_name = "Living / Dining / Kitchen"
            #format unit type
            if "1B1P" in unit_type or "1B 1P" in unit_type:
                unit_type = "1B1P"
            elif "1B2P" in unit_type or "1B 2P" in unit_type :
                unit_type = "1B2P"
            elif "2B3P" in unit_type or "2B 3P" in unit_type:
                unit_type = "2B3P"
            elif "2B4P" in unit_type or "2B 4P" in unit_type:
                unit_type = "2B4P"
            elif "3B5P" in unit_type or "3B 5P" in unit_type:
                unit_type = "3B5P"
            elif "3B6P" in unit_type or "3B 6P" in unit_type:
                unit_type = "3B6P"
            elif "4B6P" in unit_type or "4B 6P" in unit_type:
                unit_type = "4B6P"

            # look for room in dictionary and set Area Requirement value
            try:
                get_req = area_dict[unit_type][room_name]
                area_req.Set(convert_to_internal(get_req))
            except:
                area_req.Set(0)



