__title__ = "Area to\n Generic Model"
__doc__ = "Calculate area of a Generic Model families. Carries over unit parameters"

# Update Unit Area Instance from Volume for SELECTED Generic Models
# Area = Volume / Height (height from bounding box)
# Does nothing if nothing is selected.
from pyrevit import revit, DB, script, forms
import os

doc = revit.doc
uidoc = revit.uidoc
app = revit.doc.Application
logger = script.get_logger()
output = script.get_output()

AREA_PARAM_NAME = "Unit Area Instance"   # Shared parameter to add to family

# Show menu to select action
menu_options = [
    "1. Load Family Parameters",
    "2. Calculate Area"
]

selected_option = forms.SelectFromList.show(
    menu_options,
    title="Area to Generic Model",
    button_name="Select",
    multiselect=False
)

if not selected_option:
    script.exit()

# Get selection (needed for both options)
selection_ids = list(uidoc.Selection.GetElementIds())
if not selection_ids:
    forms.alert("No elements selected. Select Generic Models and run again.", title="Area Calculation")
    script.exit()

elems = [doc.GetElement(eid) for eid in selection_ids]

# Option 1: Load Family Parameters
if selected_option == "1. Load Family Parameters":
    families_processed = set()
    families_updated = 0
    families_skipped = 0

    output.print_md("### Loading Unit Area Instance Parameter into Families\n")

    for elem in elems:
        # Only Generic Models
        if elem.Category is None or elem.Category.Id.IntegerValue != int(DB.BuiltInCategory.OST_GenericModel):
            continue
        
        if not isinstance(elem, DB.FamilyInstance):
            continue
        
        family = elem.Symbol.Family
        family_id = family.Id
        
        # Skip if we've already processed this family
        if family_id in families_processed:
            continue
        
        families_processed.add(family_id)
        
        # Check if family is editable (not in-place)
        if family.IsInPlace:
            logger.debug("Skipping in-place family: {}".format(family.Name))
            families_skipped += 1
            continue
        
        # Get family document path
        try:
            # Try to get family file path
            family_path = DB.FamilyPathUtils.GetFamilyPath(doc, family_id)
            if not os.path.exists(family_path):
                logger.debug("Family file not found: {}".format(family_path))
                families_skipped += 1
                continue
            
            # Open family document
            family_doc = app.OpenDocumentFile(family_path)
        except Exception as e:
            logger.debug("Could not open family document for {}: {}".format(family.Name, e))
            families_skipped += 1
            continue
        
        # Check for existing Unit Area Instance parameter and its group
        family_mgr = family_doc.FamilyManager
        existing_param = None
        param_in_dimensions = False
        
        for param in family_mgr.Parameters:
            if param.Definition.Name == AREA_PARAM_NAME:
                existing_param = param
                # Check if it's in Dimensions/Geometry group
                if param.Definition.ParameterGroup == DB.BuiltInParameterGroup.PG_GEOMETRY:
                    param_in_dimensions = True
                break
        
        # Load shared parameter into family
        try:
            # Get shared parameter file
            shared_param_file = app.OpenSharedParameterFile()
            if shared_param_file is None:
                output.print_md("**Warning: No shared parameter file is set. Please set it in Revit Options.**")
                family_doc.Close(False)
                families_skipped += 1
                continue
            
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
                family_doc.Close(False)
                families_skipped += 1
                continue
            
            # Process family parameter
            t_family = DB.Transaction(family_doc, "Update Unit Area Instance Parameter")
            t_family.Start()
            
            try:
                # If parameter exists but is in wrong group (Text), delete it
                if existing_param and not param_in_dimensions:
                    family_mgr.RemoveParameter(existing_param)
                    logger.debug("Removed Unit Area Instance from Text group in family: {}".format(family.Name))
                
                # If parameter doesn't exist or was in wrong group, add it to Dimensions
                if not param_in_dimensions:
                    new_param = family_mgr.AddParameter(
                        area_def,
                        DB.BuiltInParameterGroup.PG_GEOMETRY,  # Dimensions/Geometry group
                        True  # Instance parameter
                    )
                    logger.debug("Added Unit Area Instance to Dimensions group in family: {}".format(family.Name))
                
                t_family.Commit()
                
                # Save family
                save_opts = DB.SaveAsOptions()
                save_opts.OverwriteExistingFile = True
                family_doc.SaveAs(family_path, save_opts)
                
                # Reload family into project
                reload_opts = DB.FamilyLoadOptions()
                reload_opts.OnFamilyFound = DB.FamilyLoadOptions.OnFamilyFoundAction.UseExisting
                doc.LoadFamily(family_path, reload_opts)
                
                families_updated += 1
                
            except Exception as e:
                t_family.RollBack()
                logger.debug("Failed to update Unit Area Instance parameter in family {}: {}".format(family.Name, e))
                families_skipped += 1
            
            finally:
                family_doc.Close(False)
        
        except Exception as e:
            logger.debug("Error processing family {}: {}".format(family.Name, e))
            families_skipped += 1
            try:
                family_doc.Close(False)
            except:
                pass

    output.print_md(
        "**Step 1 Complete:**\n"
        "- Families processed: **{}**\n"
        "- Families updated: **{}**\n"
        "- Families skipped: **{}**\n".format(
            len(families_processed), families_updated, families_skipped
        )
    )

    forms.alert(
        "Family Parameters Updated!\n\n"
        "Families processed: {}\n"
        "Families updated: {}\n"
        "Families skipped: {}".format(
            len(families_processed), families_updated, families_skipped
        ),
        title="Load Family Parameters Complete"
    )

# Option 2: Calculate Area
elif selected_option == "2. Calculate Area":
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