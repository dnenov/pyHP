__title__ = "Area to\n Generic Model"
__doc__ = "Step 1: Load Unit Area Instance parameter into family (General group) and reload. Step 2: Calculate area of Generic Model families."

from pyrevit import revit, DB, script, forms

doc = revit.doc
uidoc = revit.uidoc
app = revit.doc.Application
logger = script.get_logger()
output = script.get_output()

AREA_PARAM_NAME = "Unit Area Instance"   # Parameter name

# Custom FamilyLoadOptions class for reloading families
class SimpleLoadOptions(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.value = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues.value = True
        source.value = DB.FamilySource.Family
        return True

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
    families_already_correct = 0
    error_messages = []

    output.print_md("### Converting Unit Area Instance to Area type in Text group\n")

    # Collect unique families from selection
    families = set()
    for elem in elems:
        # Only Generic Models
        if elem.Category is None or elem.Category.Id.IntegerValue != int(DB.BuiltInCategory.OST_GenericModel):
            continue
        
        if isinstance(elem, DB.FamilyInstance):
            families.add(elem.Symbol.Family)
        elif isinstance(elem, DB.FamilySymbol):
            families.add(elem.Family)
    
    if not families:
        forms.alert("No Generic Model families found in selection.", title="Error", exitscript=True)
    
    families = list(families)

    # Process each family
    for family in families:
        fam_name = family.Name
        
        # Check if family is editable (not in-place)
        if family.IsInPlace:
            output.print_md("**Skipping in-place family: {}**\n".format(fam_name))
            families_skipped += 1
            continue
        
        try:
            # Open family document
            family_doc = doc.EditFamily(family)
            family_mgr = family_doc.FamilyManager
            
            output.print_md("**Processing family: {}**\n".format(fam_name))
            
            # Check for existing Unit Area Instance parameter
            existing_param = None
            existing_is_text_type = False
            existing_in_text_group = False
            existing_is_area_type = False
            existing_is_shared = False
            
            for param in family_mgr.Parameters:
                if param.Definition.Name == AREA_PARAM_NAME:
                    existing_param = param
                    fam_param_type = param.Definition.ParameterType
                    existing_is_text_type = (fam_param_type == DB.ParameterType.Text)
                    existing_is_area_type = (fam_param_type == DB.ParameterType.Area)
                    existing_in_text_group = (param.Definition.ParameterGroup == DB.BuiltInParameterGroup.PG_TEXT)
                    # Check if it's a shared parameter
                    try:
                        # Shared parameters have an ExternalDefinition
                        if hasattr(param.Definition, 'GetExternalDefinition'):
                            existing_is_shared = True
                    except:
                        pass
                    
                    output.print_md("  - Found existing: Type={}, Group={}, Shared={}\n".format(
                        fam_param_type, param.Definition.ParameterGroup, existing_is_shared))
                    break
            
            # Start transaction
            t_family = DB.Transaction(family_doc, "Update Unit Area Instance Parameter")
            t_family.Start()
            
            try:
                # If parameter exists and is already Area type in Text group, skip
                if existing_param and existing_is_area_type and existing_in_text_group:
                    output.print_md("  - Already correct (Area type in Text group), skipping\n")
                    t_family.RollBack()
                    families_already_correct += 1
                    family_doc.Close(False)
                    continue
                
                # Delete existing parameter if it exists (especially if Text type or shared)
                if existing_param:
                    family_mgr.RemoveParameter(existing_param)
                    if existing_is_text_type:
                        output.print_md("  - Deleted Text type parameter\n")
                    elif existing_is_shared:
                        output.print_md("  - Deleted shared parameter (will create Area family parameter)\n")
                    else:
                        output.print_md("  - Deleted existing parameter\n")
                
                # Create new Area-type family parameter (not shared) in Text group
                # We'll create it as a family parameter since we can't modify shared parameter file
                param_def = family_mgr.AddParameter(
                    AREA_PARAM_NAME,
                    DB.BuiltInParameterGroup.PG_TEXT,  # Text group
                    DB.ParameterType.Area,  # Area type
                    True  # Instance parameter
                )
                
                output.print_md("  - Created Area type family parameter in Text group\n")
                logger.info("Created Area type parameter in family: {}".format(fam_name))
                families_updated += 1
                
                t_family.Commit()
                
                # Reload family into project
                load_opt = SimpleLoadOptions()
                family_doc.LoadFamily(doc, load_opt)
                family_doc.Close(False)
                output.print_md("  - Family reloaded\n")
                
            except Exception as e:
                t_family.RollBack()
                error_msg = "Failed to update parameter in family {}: {}".format(fam_name, e)
                output.print_md("**ERROR: {}**\n".format(error_msg))
                logger.error(error_msg)
                error_messages.append(error_msg)
                families_skipped += 1
                try:
                    family_doc.Close(False)
                except:
                    pass
                
        except Exception as e:
            error_msg = "Error processing family {}: {}".format(fam_name, e)
            output.print_md("**ERROR: {}**\n".format(error_msg))
            logger.error(error_msg)
            error_messages.append(error_msg)
            families_skipped += 1
            try:
                family_doc.Close(False)
            except:
                pass

    output.print_md(
        "\n**Step 1 Complete:**\n"
        "- Families processed: **{}**\n"
        "- Families updated: **{}**\n"
        "- Families already correct: **{}**\n"
        "- Families skipped: **{}**\n".format(
            len(families), families_updated, families_already_correct, families_skipped
        )
    )
    
    output.print_md("\n**Note: Created as family parameters (not shared) with Area type in Text group.**\n")
    output.print_md("**This will work for calculations and appear under Text in properties panel.**\n")
    
    if error_messages:
        output.print_md("\n**Errors encountered:**\n")
        for err in error_messages:
            output.print_md("- {}\n".format(err))

    forms.alert(
        "Family Parameters Updated!\n\n"
        "Families processed: {}\n"
        "Families updated: {}\n"
        "Families already correct: {}\n"
        "Families skipped: {}\n\n"
        "Created as Area type family parameters in Text group.".format(
            len(families), families_updated, families_already_correct, families_skipped
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