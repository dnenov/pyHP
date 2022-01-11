from pyrevit import revit, DB, script, forms, HOST_APP
from rpw.ui.forms import (FlexForm, Label, ComboBox, Separator, Button)
import tempfile
import helper
import re
from pyrevit.revit.db import query

logger = script.get_logger()
output = script.get_output()

# todo: name of the element
# todo: material to type and shared parameter

sp_unit_material = helper.get_shared_param_by_name_type("Unit Material", DB.ParameterType.Material)
if not sp_unit_material:
    forms.alert(msg="No suitable parameter", \
        sub_msg="There is no suitable parameter to use for Unit Material. Please add a shared parameter 'Unit "
                "Material' of Material Type", \
        ok=True, \
        warn_icon=True, exitscript=True)

# use preselected elements, filtering rooms only
pre_selection = helper.preselection_with_filter(DB.BuiltInCategory.OST_Rooms)
# or select rooms
if pre_selection and forms.alert("You have selected {} elements. Do you want to use them?".format(len(pre_selection))):
    selection = pre_selection
else:
    selection = helper.select_rooms_filter()

if selection:
    # Create family doc from template
    fam_template_path = __revit__.Application.FamilyTemplatePath + "\Conceptual Mass\Metric Mass.rft"

    # iterate through rooms
    for room in selection:
        mass_placement_point = room.get_BoundingBox(None).Min
        # define new family doc
        try:
            new_family_doc = revit.doc.Application.NewFamilyDocument(fam_template_path)
        except NameError:
            forms.alert(msg="No Template",
                        sub_msg="Cannot find a Conceptual Mass Template in the default location.",
                        ok=True,
                        warn_icon=True, exitscript=True)

        # To name the room, collect its parameters:
        project_number = revit.doc.ProjectInformation.Number
        if not project_number:
            project_number = "Project"
        dept = room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString()  # chosen parameter for Department
        if not dept:
            dept = "Department"
        room_name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
        room_number = room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString()
        # construct family and family type names:
        fam_name = project_number + "_" + room_name + "_" + str(dept)
        # replace bad characters
        fam_name = re.sub(r'[^\w\-_\. ]', '', fam_name)
        fam_type_name = re.sub(r'[^\w\-_\. ]', '', room_name)

        # check if family already exists:
        while helper.get_fam(fam_name):
            fam_name = fam_name + "_Copy 1"

        # Save family in temp folder
        fam_path = tempfile.gettempdir() + "/" + fam_name + ".rfa"
        saveas_opt = DB.SaveAsOptions()
        saveas_opt.OverwriteExistingFile = True
        try:
            new_family_doc.SaveAs(fam_path, saveas_opt)
        except Exceptions.FileAccessException:
            fam_path = fam_path.replace(".rfa", "_Copy 1.rfa")
            new_family_doc.SaveAs(fam_path, saveas_opt)
            fam_name = fam_name + "_Copy 1"

        # Create extrusion from room boundaries
        with revit.Transaction(doc=new_family_doc, name="Create FreeForm Element"):
            room_geo = room.ClosedShell
            for geo in room_geo:
                if isinstance(geo, DB.Solid) and geo.Volume > 0.0:
                    freeform = DB.FreeFormElement.Create(new_family_doc, geo)
                    new_family_doc.Regenerate()
                    delta = DB.XYZ(0, 0, 0) - freeform.get_BoundingBox(None).Min
                    move_ff = DB.ElementTransformUtils.MoveElement(
                        new_family_doc, freeform.Id, delta
                    )
                    # create and associate a material parameter

                    ext_mat_param = freeform.get_Parameter(DB.BuiltInParameter.MATERIAL_ID_PARAM)
                    try:
                        new_mat_param = new_family_doc.FamilyManager.AddParameter(sp_unit_material,
                                                                                  DB.BuiltInParameterGroup.PG_MATERIALS,
                                                                                  False)
                        new_family_doc.FamilyManager.AssociateElementParameterToFamilyParameter(ext_mat_param,
                                                                                                new_mat_param)
                    except Exception as err:
                        logger.error(err)

        # save and close family
        save_opt = DB.SaveOptions()
        new_family_doc.Save(save_opt)
        new_family_doc.Close()

        # Reload family with extrusion and place it in the same position as the room
        with revit.Transaction("Load Family", revit.doc):
            loaded_f = revit.db.create.load_family(fam_path, doc=revit.doc)
            # find family symbol and activate
            fam_symbol = helper.get_fam(fam_name)
            if not fam_symbol.IsActive:
                fam_symbol.Activate()
                revit.doc.Regenerate()
            # place instance
            new_fam_instance = revit.doc.Create.NewFamilyInstance(room.Level.GetPlaneReference(),
                                                                  mass_placement_point,
                                                                  DB.XYZ(1, 0, 0),
                                                                  fam_symbol
                                                                  )
            print(
                "Created and placed Mass family instance : {1} - {2} {0} ".format(
                    output.linkify(new_fam_instance.Id),
                    fam_name, fam_type_name))
