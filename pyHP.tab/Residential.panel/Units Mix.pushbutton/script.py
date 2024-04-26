from pyrevit import revit, DB, script, forms, HOST_APP
import sys

FEC = DB.FilteredElementCollector
BIC = DB.BuiltInCategory
doc = revit.doc
active_view = revit.active_view
output = script.get_output()
HABITABLE_ROOMS_PARAMETER_NAME = "Habitable Rooms"
TENURE_PARAMETER_NAME = "Tenure"

def get_description(element):
    description = doc.GetElement(element.GetTypeId()).get_Parameter(DB.BuiltInParameter.ALL_MODEL_DESCRIPTION).AsValueString()
    return description


# get the design option of the active view
def get_design_option_of_view(view):
    units_in_view = FEC(doc, active_view.Id).OfCategory(
        BIC.OST_GenericModel).WhereElementIsNotElementType().ToElements()
    for unit in units_in_view:
        try:
            if get_description(unit) == "Flat":
                return unit
        except:
            return None


one_visible_unit = get_design_option_of_view(active_view)
design_option = one_visible_unit.DesignOption

collect_all_units = FEC(doc).OfCategory(BIC.OST_GenericModel).WhereElementIsNotElementType().ToElements()
units_of_active_design_option = []

# iterate through all units and gather all units matching the design option
for unit in collect_all_units:
    try:
        if unit.DesignOption.Id == design_option.Id and get_description(unit):
            units_of_active_design_option.append(unit)
    except AttributeError:
        pass

if len(units_of_active_design_option) == 0:
    output.print_md("Zero units found")
    sys.exit()

# counters
private_hab_rooms = 0
shared_hab_rooms = 0
social_hab_rooms = 0

for unit in units_of_active_design_option:
    unit_type = doc.GetElement(unit.GetTypeId())
    try:
        tenure = unit_type.LookupParameter(TENURE_PARAMETER_NAME).AsValueString()
        hab_rooms = unit_type.LookupParameter(HABITABLE_ROOMS_PARAMETER_NAME).AsValueString()

        if tenure == "Private":
            private_hab_rooms += int(hab_rooms)
        elif tenure == "S/O" or tenure == "Shared Ownership":
            shared_hab_rooms += int(hab_rooms)
        if tenure == "Rented" or tenure == "Social Rent":
            social_hab_rooms += int(hab_rooms)
    except:
        pass

# calculate
affordable_room_count = social_hab_rooms + shared_hab_rooms
total_room_count = private_hab_rooms + affordable_room_count

private_percent = float(private_hab_rooms * 100.0 / total_room_count)
affordable_percent = float(affordable_room_count * 100.0 / total_room_count)
shared_percent = float(shared_hab_rooms * 100.0 / affordable_room_count)
social_percent = float(social_hab_rooms * 100.0 / affordable_room_count)

private_colour = "#7E95B3"
affordable_colour = "#B3B133"
shared_colour = "#559050"
social_colour = "#FFA037"

output.print_md("# Units Mix - Option {}".format(design_option.Name))
output.print_md("--------")
output.print_md("Total units in option - {}".format(len(units_of_active_design_option)))

# Private / Affordable split
chartPrivateAffordable = output.make_doughnut_chart()
chartPrivateAffordable.options.title = {
    "display": True,
    "text": "Private/Affordable split by Habitable Room",
    "fontSize": 15,
    "fontStyle": "bold",
    "position": "top"
}
chartPrivateAffordable.options.legend = {"position": "left", "fullWidth": False}
chartPrivateAffordable.data.labels = ["Private", "Affordable"]
dataset = chartPrivateAffordable.data.new_dataset("Not Standard")

dataset.data = [private_hab_rooms, shared_hab_rooms + social_hab_rooms]
dataset.backgroundColor = [private_colour, affordable_colour]
chartPrivateAffordable.set_height = 1
chartPrivateAffordable.draw()

output.print_md(" . . . . . Habitable Rooms | Percent")
output.print_md("Private . . . . . . . . {} | {:.2f} %".format(private_hab_rooms, private_percent))
output.print_md("Affordable . . . . . . {} | {:.2f} %".format(affordable_room_count, affordable_percent))
output.print_md("TOTALS . . . . . . . . {}".format(total_room_count))
output.print_md("--------")

# Affordable split chart
chartAffordableSplit = output.make_doughnut_chart()
chartAffordableSplit.options.title = {
    "display": True,
    "text": "Affordable split by Habitable Room",
    "fontSize": 15,
    "fontStyle": "bold",
    "position": "top"
}
chartAffordableSplit.options.legend = {"position": "left", "fullWidth": False}
chartAffordableSplit.data.labels = ["Shared Ownership", "Social Rent"]
set = chartAffordableSplit.data.new_dataset("Not Standard")

set.data = [shared_hab_rooms, social_hab_rooms]
set.backgroundColor = [shared_colour, social_colour]
chartAffordableSplit.set_height = 1
chartAffordableSplit.draw()

output.print_md(". . . . . . . . . .Habitable Rooms     |         Percent")
output.print_md(
    "Shared Ownership. . . . . . . . {}            |            {:.2f} %".format(shared_hab_rooms, shared_percent))
output.print_md(
    "Social Rent . . . . . . . . . . {}            |            {:.2f} %".format(str(social_hab_rooms), social_percent))
output.print_md("TOTALS . . . . . . . . {}".format(affordable_room_count))
output.print_md("--------")
