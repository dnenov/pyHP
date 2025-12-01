__title__ = "Area to\n Generic Model"
__doc__ = "Calculate area of Generic Model families. Make sure shared parameter Unit Area Instance is set to dimensions in the revit family and Formula tab is empty" 

from pyrevit import revit, DB, script, forms

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
output = script.get_output()

AREA_PARAM_NAME = "Unit Area Instance"   # Parameter name

# Get selection
selection_ids = list(uidoc.Selection.GetElementIds())
if not selection_ids:
    forms.alert("No elements selected. Select Generic Models and run again.", title="Area Calculation")
    script.exit()

elems = [doc.GetElement(eid) for eid in selection_ids]

# Calculate Area
output.print_md("### Calculating Unit Area\n")

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

    # Get bounding box to infer height (no Height parameter needed)
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

        # Area in internal units = Volume / Height (from bounding box)
        area_val = vol_val / height_internal

        area_param.Set(area_val)
        updated += 1

    except Exception as e:
        logger.debug("Failed on element {}: {}".format(elem.Id, e))
        skipped += 1

t.Commit()

output.print_md(
    "**Calculation Complete:**\n"
    "- Selected elements: **{}**\n"
    "- Updated Generic Models: **{}**\n"
    "- Skipped: **{}**\n".format(
        len(elems), updated, skipped
    )
)

if updated > 0:
    forms.alert(
        "Area calculated for {} Generic Model(s)!\n\n"
        "Selected: {}\n"
        "Updated: {}\n"
        "Skipped: {}".format(updated, len(elems), updated, skipped),
        title="Area Calculation Complete"
    )
else:
    forms.alert(
        "No Generic Models were updated.\n\n"
        "Selected: {}\n"
        "Skipped: {}\n\n"
        "Make sure you have selected Generic Model elements with volume and Unit Area Instance parameters.".format(
            len(elems), skipped
        ),
        title="Area Calculation"

    )
