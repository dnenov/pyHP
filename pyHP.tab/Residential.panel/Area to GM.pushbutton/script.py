__title__ = "Area to\n Generic Model"
__doc__ = "Calculate area of a Generic Model families. Carries over unit parameters"

# Update Unit Area Instance from Volume for SELECTED Generic Models
# Area = Volume / Height (height from bounding box)
# Does nothing if nothing is selected.

from pyrevit import revit, DB, script

doc = revit.doc
uidoc = revit.uidoc
app = revit.doc.Application
logger = script.get_logger()
output = script.get_output()

AREA_PARAM_NAME = "Unit Area Instance"   # Project shared parameter (Area)

# Get selection
selection_ids = list(uidoc.Selection.GetElementIds())
if not selection_ids:
    output.print_md("**No elements selected. Select Generic Models and run again.**")
    script.exit()

elems = [doc.GetElement(eid) for eid in selection_ids]

# Step 1: Check if Unit Area Instance parameter exists (as project or family parameter)
output.print_md("### Step 1: Checking for Unit Area Instance Parameter\n")

# Check if parameter exists as project parameter
param_exists_as_project = False
binding_map = doc.ParameterBindings
iterator = binding_map.ForwardIterator()
while iterator.MoveNext():
    definition = iterator.Key
    if definition.Name == AREA_PARAM_NAME:
        param_exists_as_project = True
        break

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
        if shared_param_file is None:
            output.print_md("**Warning: No shared parameter file is set. Please set it in Revit Options.**")
            script.exit()
        
        # Find Unit Area Instance parameter definition
        area_def = None
        for group in shared_param_file.Groups:
            for defn in group.Definitions:
                if defn.Name == AREA_PARAM_NAME:
                    area_def = defn
                    break
            if area_def:
                break
        
        if not area_def:
            output.print_md("**Warning: 'Unit Area Instance' shared parameter not found in shared parameter file.**")
            script.exit()
        
        # Create category set for Generic Model
        cat_set = DB.CategorySet()
        generic_model_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_GenericModel)
        cat_set.Insert(generic_model_cat)
        
        # Create instance binding
        binding = app.Create.NewInstanceBinding(cat_set)
        
        # Add as project parameter
        t = DB.Transaction(doc, "Add Unit Area Instance Project Parameter")
        t.Start()
        try:
            binding_map.Insert(area_def, binding)
            t.Commit()
            output.print_md("**Successfully added 'Unit Area Instance' as project parameter for Generic Model category.**\n")
        except Exception as e:
            t.RollBack()
            output.print_md("**Error adding project parameter: {}**\n".format(e))
            logger.debug("Error: {}".format(e))
    except Exception as e:
        output.print_md("**Error accessing shared parameter file: {}**\n".format(e))
        logger.debug("Error: {}".format(e))

# Step 2: Calculate area (existing functionality)
output.print_md("### Step 2: Calculating Unit Area\n")

def get_volume_param(elem):
    # Try built-in volume first
    vol = elem.get_Parameter(DB.BuiltInParameter.HOST_VOLUME_COMPUTED)
    if vol and vol.HasValue:
        return vol
    # Fallback: parameter literally called "Volume"
    vol2 = elem.LookupParameter("Volume")
    if vol2 and vol2.HasValue:
        return vol2
    return None

updated = 0
skipped = 0

t = DB.Transaction(doc, "Update Unit Area from Volume (Selection Only)")
t.Start()

for elem in elems:
    # Only Generic Models
    if elem.Category is None or elem.Category.Id.IntegerValue != int(DB.BuiltInCategory.OST_GenericModel):
        skipped += 1
        continue

    vol_param = get_volume_param(elem)
    area_param = elem.LookupParameter(AREA_PARAM_NAME)

    if not vol_param or not area_param:
        skipped += 1
        continue

    # Get bounding box to infer height
    bbox = elem.get_BoundingBox(None)
    if not bbox:
        skipped += 1
        continue

    height_internal = bbox.Max.Z - bbox.Min.Z
    if height_internal <= 0:
        skipped += 1
        continue

    try:
        vol_val = vol_param.AsDouble()  # internal units 

        # Area in internal units 
        area_val = vol_val / height_internal

        area_param.Set(area_val)
        updated += 1

    except Exception as e:
        logger.debug("Failed on element {}: {}".format(elem.Id, e))
        skipped += 1

t.Commit()

output.print_md(
    "**Step 2 Complete:**\n"
    "- Selected elements: **{}**\n"
    "- Updated Generic Models: **{}**\n"
    "- Skipped (wrong category / missing params / invalid height): **{}**\n".format(
        len(elems), updated, skipped
    )
)

output.print_md("\n### All Steps Complete!")