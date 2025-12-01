__title__ = "Area to\n Generic Model"
__doc__ = "Calculate area of a Generic Model families. Carries over unit parameters"

# Update Unit Area Instance from Volume for SELECTED Generic Models
# Area = Volume / Height (height from bounding box)
# Does nothing if nothing is selected.

from pyrevit import revit, DB, script
from pyrevit import revit, DB, script, forms

doc = revit.doc
uidoc = revit.uidoc

@@ -18,7 +18,7 @@ AREA_PARAM_NAME = "Unit Area Instance"   # Project shared parameter (Area)
# Get selection
selection_ids = list(uidoc.Selection.GetElementIds())
if not selection_ids:
    output.print_md("**No elements selected. Select Generic Models and run again.**")
    forms.alert("No elements selected. Select Generic Models and run again.", title="Area Calculation")
    script.exit()

elems = [doc.GetElement(eid) for eid in selection_ids]


@@ -57,7 +57,7 @@ else:
        # Get shared parameter file
        shared_param_file = app.OpenSharedParameterFile()
        if shared_param_file is None:
            output.print_md("**Warning: No shared parameter file is set. Please set it in Revit Options.**")
            forms.alert("No shared parameter file is set. Please set it in Revit Options.", title="Area Calculation")
            script.exit()
        
        # Find Unit Area Instance parameter definition

@@ -71,7 +71,7 @@ else:
                break
        
        if not area_def:
            output.print_md("**Warning: 'Unit Area Instance' shared parameter not found in shared parameter file.**")
            forms.alert("'Unit Area Instance' shared parameter not found in shared parameter file.", title="Area Calculation")
            script.exit()
        
        # Create category set for Generic Model


@@ -156,6 +156,7 @@ for elem in elems:

t.Commit()

# Show completion message
output.print_md(
    "**Step 2 Complete:**\n"
    "- Selected elements: **{}**\n"

@@ -165,4 +166,24 @@ output.print_md(
    )
)

output.print_md("\n### All Steps Complete!")
output.print_md("\n### All Steps Complete!")

# Show popup confirmation
if updated > 0:
    forms.alert(
        "Area calculated for {} Generic Model(s)!\n\n"
        "Selected elements: {}\n"
        "Updated: {}\n"
        "Skipped: {}".format(updated, len(elems), updated, skipped),
        title="Area Calculation Complete"
    )
else:
    forms.alert(
        "No Generic Models were updated.\n\n"
        "Selected elements: {}\n"
        "Skipped: {}\n\n"
        "Make sure you have selected Generic Model elements with volume and Unit Area Instance parameters.".format(
            len(elems), skipped
        ),
        title="Area Calculation"
    )