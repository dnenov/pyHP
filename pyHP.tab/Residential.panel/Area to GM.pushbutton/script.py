__title__ = "Area to\n Generic Model"
__doc__ = "Calculate area of a Generic Model families. Carries over unit parameters"

# Update Unit Area Instance from Volume for SELECTED Generic Models
# Area = Volume / Height (height from bounding box)
# Does nothing if nothing is selected.


from pyrevit import revit, DB, script

doc = revit.doc
uidoc = revit.uidoc
app = revit.doc.Application  # Changed from revit.app
app = revit.doc.Application
logger = script.get_logger()
output = script.get_output()


@@ -24,20 +23,36 @@ if not selection_ids:

elems = [doc.GetElement(eid) for eid in selection_ids]

# Step 1: Add Unit Area Instance as project parameter (if not already exists)
output.print_md("### Step 1: Adding Unit Area Instance as Project Parameter\n")
# Step 1: Check if Unit Area Instance parameter exists (as project or family parameter)
output.print_md("### Step 1: Checking for Unit Area Instance Parameter\n")

# Check if parameter already exists as project parameter
param_exists = False
# Check if parameter exists as project parameter
param_exists_as_project = False
binding_map = doc.ParameterBindings
iterator = binding_map.ForwardIterator()
while iterator.MoveNext():
    definition = iterator.Key
    if definition.Name == AREA_PARAM_NAME:
        param_exists = True
        param_exists_as_project = True
        break

if not param_exists:
# Check if parameter exists on any selected element (could be family parameter)
param_exists_on_element = False
for elem in elems:
    if elem.Category is None or elem.Category.Id.IntegerValue != int(DB.BuiltInCategory.OST_GenericModel):
        continue
    area_param = elem.LookupParameter(AREA_PARAM_NAME)
    if area_param:
        param_exists_on_element = True
        break

if param_exists_as_project:
    output.print_md("**'Unit Area Instance' already exists as project parameter.**\n")
elif param_exists_on_element:
    output.print_md("**'Unit Area Instance' already exists as family parameter. Using existing parameter.**\n")
else:
    # Parameter doesn't exist, add it as project parameter
    output.print_md("**Adding 'Unit Area Instance' as project parameter...**\n")
    try:
        # Get shared parameter file
        shared_param_file = app.OpenSharedParameterFile()


@@ -81,8 +96,6 @@ if not param_exists:
    except Exception as e:
        output.print_md("**Error accessing shared parameter file: {}**\n".format(e))
        logger.debug("Error: {}".format(e))
else:
    output.print_md("**'Unit Area Instance' parameter already exists as project parameter.**\n")

# Step 2: Calculate area (existing functionality)
output.print_md("### Step 2: Calculating Unit Area\n")

