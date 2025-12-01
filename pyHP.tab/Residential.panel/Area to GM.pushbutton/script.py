__title__ = "Area to\n Generic Model"
__doc__ = "Calculate area of a Generic Model families. Carries over unit parameters"

# Update Unit Area Instance from Volume for SELECTED Generic Models
# Area = Volume / Height (height from bounding box)
# Does nothing if nothing is selected.

from pyrevit import revit, DB, script

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
output = script.get_output()

AREA_PARAM_NAME = "Unit Area Instance"   # Project shared parameter (Area)


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


# Get selection
selection_ids = list(uidoc.Selection.GetElementIds())
if not selection_ids:
    output.print_md("**No elements selected. Select Generic Models and run again.**")
    script.exit()

elems = [doc.GetElement(eid) for eid in selection_ids]

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
        vol_val = vol_param.AsDouble()  # internal units (ft³)

        # Area in internal units (ft²) = V (ft³) / h (ft)
        area_val = vol_val / height_internal

        area_param.Set(area_val)
        updated += 1

    except Exception as e:
        logger.debug("Failed on element {}: {}".format(elem.Id, e))
        skipped += 1

t.Commit()

output.print_md(
    f"### Unit Area Update Complete (Selection Only)\n"
    f"- Selected elements: **{len(elems)}**\n"
    f"- Updated Generic Models: **{updated}**\n"
    f"- Skipped (wrong category / missing params / invalid height): **{skipped}**"
)
