from pyrevit import revit, DB, script, forms, HOST_APP
from rpw.ui.forms import (FlexForm, Label, ComboBox, Separator, Button)
import tempfile
import re
import sys

doc = revit.doc
collect_units = DB.FilteredElementCollector(doc).OfCategory(
    DB.BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

# counters
private_hab_rooms = 0
shared_hab_rooms = 0
social_hab_rooms = 0

for unit in collect_units:
    unit_type = doc.GetElement(unit.GetTypeId())
    tenure = unit_type.LookupParameter("Tenure").AsValueString()
    print (tenure)
    hab_rooms = unit_type.LookupParameter("Habitable Rooms").AsValueString()

    if tenure == "Private":
        private_hab_rooms += int(hab_rooms)
    elif tenure == "S/O" or tenure == "Shared Ownership":
        shared_hab_rooms += int(hab_rooms)
    if tenure == "Rented" or tenure == "Social Rent":
        social_hab_rooms += int(hab_rooms)


# calculate
affordable_room_count = social_hab_rooms + shared_hab_rooms
total_room_count = private_hab_rooms + affordable_room_count
private_percent = private_hab_rooms* 100/total_room_count
affordable_percent = affordable_room_count* 100/total_room_count
shared_percent = shared_hab_rooms* 100/affordable_room_count
social_percent = social_hab_rooms* 100/affordable_room_count


private_colour = "#7E95B3"
affordable_colour = "#B3B133"
shared_colour = "#559050"
social_colour = "#FFA037"

output = script.get_output()

output.print_md("# Units Mix")
output.print_md("--------")

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

dataset.data = [private_hab_rooms, shared_hab_rooms+social_hab_rooms]
dataset.backgroundColor = [private_colour, affordable_colour]
chartPrivateAffordable.set_height = 1
chartPrivateAffordable.draw()


output.print_md(" . . . . . Habitable Rooms | Percent")
output.print_md("Private . . . . . . . . {0} | {1} %".format(private_hab_rooms, private_percent))
output.print_md("Affordable . . . . . . {0} | {1} %".format(affordable_room_count, affordable_percent))

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
output.print_md("Shared Ownership. . . . . . . . {0}            |            {1} %".format(shared_hab_rooms, shared_percent))
output.print_md("Social Rent . . . . . . . . . . {0}            |            {1} %".format(str(social_hab_rooms), social_percent))
