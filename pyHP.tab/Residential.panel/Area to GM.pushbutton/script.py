__title__ = "Area to\n Generic Model"
__doc__ = "Step 1: Load Unit Area Instance parameter into family (General group) and reload. Step 2: Calculate area of Generic Model families."

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
    error_messages = []

    output.print_md("### Loading Unit Area Instance Parameter into Families\n")

    # First, check shared parameter file
    try:
        shared_param_file = app.OpenSharedParameterFile()
        if shared_param_file is None:
            forms.alert("No shared parameter file is set. Please set it in Revit Options.", title="Error")
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
            forms.alert("'Unit Area Instance' shared parameter not found in shared parameter file.", title="Error")
            script.exit()
        
        output.print_md("**Shared parameter '{}' found in shared parameter file.**\n".format(AREA_PARAM_NAME))
    except Exception as e:
        forms.alert("Error accessing shared parameter file: {}".format(e), title="Error")
        script.exit()

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
            output.print_md("**Skipping in-place family: {}**\n".format(family.Name))
            families_skipped += 1
            continue
        
        # Get family document path
        family_doc = None
        try:
            # Try to get family file path
            family_path = DB.FamilyPathUtils.GetFamilyPath(doc, family_id)
            if not os.path.exists(family_path):
                output.print_md("**Family file not found: {}**\n".format(family_path))
                families_skipped += 1
                continue
            
            # Open family document
            family_doc = app.OpenDocumentFile(family_path)
            output.print_md("**Processing family: {}**\n".format(family.Name))
        except Exception as e:
            error_msg = "Could not open family document for {}: {}".format(family.Name, e)
            output.print_md("**{}**\n".format(error_msg))
            logger.debug(error_msg)
            families_skipped += 1
            continue
        
        # Check for existing Unit Area Instance parameter (any group)
        family_mgr = family_doc.FamilyManager
        existing_param = None
        
        for param in family_mgr.Parameters:
            if param.Definition.Name == AREA_PARAM_NAME:
                existing_param = param
                output.print_md("  - Found existing parameter in {} group\n".format(param.Definition.ParameterGroup))
                break
        
        # Process family parameter
        try:
            t_family = DB.Transaction(family_doc, "Update Unit Area Instance Parameter")
            t_family.Start()
            
            try:
                # Always delete existing parameter if it exists (regardless of group)
                if existing_param:
                    family_mgr.RemoveParameter(existing_param)
                    output.print_md("  - Deleted existing parameter\n")
                    logger.debug("Deleted existing Unit Area Instance parameter from family: {}".format(family.Name))
                
                # Always add parameter to General group
                new_param = family_mgr.AddParameter(
                    area_def,
                    DB.BuiltInParameterGroup.PG_GENERAL,  # General group
                    True  # Instance parameter
                )
                output.print_md("  - Added parameter to General group\n")
                logger.debug("Added Unit Area Instance to General group in family: {}".format(family.Name))
                families_updated += 1
                
                t_family.Commit()
                
                # Save family
                save_opts = DB.SaveAsOptions()
                save_opts.OverwriteExistingFile = True
                family_doc.SaveAs(family_path, save_opts)
                output.print_md("  - Family saved\n")
                
                # Reload family into project
                reload_opts = DB.FamilyLoadOptions()
                reload_opts.OnFamilyFound = DB.FamilyLoadOptions.OnFamilyFoundAction.UseExisting
                doc.LoadFamily(family_path, reload_opts)
                output.print_md("  - Family reloaded into project\n")
                
            except Exception as e:
                t_family.RollBack()
                error_msg = "Failed to update Unit Area Instance parameter in family {}: {}".format(family.Name, e)
                output.print_md("**ERROR: {}**\n".format(error_msg))
                logger.debug(error_msg)
                error_messages.append(error_msg)
                families_skipped += 1
            
            finally:
                if family_doc:
                    family_doc.Close(False)
        
        except Exception as e:
            error_msg = "Error processing family {}: {}".format(family.Name, e)
            output.print_md("**ERROR: {}**\n".format(error_msg))
            logger.debug(error_msg)
            error_messages.append(error_msg)
            families_skipped += 1
            if family_doc:
                try:
                    family_doc.Close(False)
                except:
                    pass

    output.print_md(
        "\n**Step 1 Complete:**\n"
        "- Families processed: **{}**\n"
        "- Families updated: **{}**\n"
        "- Families skipped: **{}**\n".format(
            len(families_processed), families_updated, families_skipped
        )
    )
    
    if error_messages:
        output.print_md("\n**Errors encountered:**\n")
        for err in error_messages:
            output.print_md("- {}\n".format(err))

    forms.alert(
        "Family Parameters Updated!\n\n"
        "Families processed: {}\n"
        "Families updated: {}\n"
        "Families skipped: {}\n\n"
        "Check output panel for details.".format(
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