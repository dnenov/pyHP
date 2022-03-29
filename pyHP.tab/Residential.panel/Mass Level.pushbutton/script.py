import sys

from pyrevit import revit, DB, forms
from rpw.ui.forms import FlexForm, Label, TextBox, Button, ComboBox, CheckBox, Separator


masses = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_Mass).WhereElementIsNotElementType().ToElements()

params = [p.Definition.Name for p in masses[0].Parameters if p.StorageType.ToString() == "String"]

components = [
    Label("Select Mass parameter"),
    ComboBox(name="p", options=params),
    Button("Select")
    ]

form = FlexForm("Select", components)
ok = form.show()

if ok:
    # match the variables with user input
    chosen_param = form.values["p"]
else:
    sys.exit()
# selected_param = forms.select_parameters(masses[0], title="Select Parameter for Level", multiple = False, include_instance=True, include_type=False)
# print (selected_param.definition)
with revit.Transaction("Fill in Mass Levels",revit.doc):
    for m in masses:
        workplane = m.Host.get_Parameter(DB.BuiltInParameter.DATUM_TEXT).AsString()
        correct_name = workplane.replace("Level ", "")
        m.LookupParameter(chosen_param).Set(workplane)
        # print (workplane)
