__title__ = "Room to\n Generic Model"
__doc__ = "Transforms rooms into Generic Model families. Carries over unit parameters"

from pyrevit import revit, DB, script, forms, HOST_APP
from rpw.ui.forms import (FlexForm, Label, ComboBox, Separator, Button)
import tempfile
import helper
import re
import sys


logger = script.get_logger()
output = script.get_output()
# get shared parameter for the extrusion material


sp_unit_material = helper.get_shared_param_by_name_type("Unit Material", DB.ParameterType.Material)
if not sp_unit_material:
    forms.alert(msg="No suitable parameter", \
        sub_msg="There is no suitable parameter to use for Unit Material. Please add a shared parameter 'Unit Material' of Material Type", \
        ok=True, \
        warn_icon=True, exitscript=True)


# use preselected elements, filtering rooms only
pre_selection = helper.preselection_with_filter(DB.BuiltInCategory.OST_Rooms)
# or select rooms
if pre_selection and forms.alert("You have selected {} elements. Do you want to use them?".format(len(pre_selection))):
    selection = pre_selection
else:
    selection = helper.select_rooms_filter()

# test
def get_external_definition_by_name(name):
    sparam_file = HOST_APP.app.OpenSharedParameterFile()
    for def_groups in sparam_file.Groups:
        for sparam_def in def_groups.Definitions:
            if sparam_def.Name==name:
                return sparam_def




if selection:
    # Create family doc from template
    # get file template from location
    fam_template_path = __revit__.Application.FamilyTemplatePath + "\Metric Generic Model.rft"

    # format parameters for UI
    # gather and organize Room parameters: (only editable text params)
    room_parameter_set = selection[0].Parameters
    room_params_text = [p.Definition.Name for p in room_parameter_set if
                        p.StorageType.ToString() == "String" and p.IsReadOnly == False]
    # collect and organize Generic Model parameters: (only editable text type params)
    gm_parameter_set = helper.param_set_by_cat(DB.BuiltInCategory.OST_GenericModel)
    # gm_params_text = [p for p in gm_parameter_set if p.StorageType.ToString() == "String"]
    # gm_params_area = [p for p in gm_parameter_set if p.Definition.ParameterType.ToString()=="Area"]

    # if not gm_params_area:
    #     forms.alert(msg="No suitable parameter",
    #                 sub_msg="There is no suitable parameter to use for Unit Area. Please add a shared parameter 'Unit "
    #                         "Area' of Area Type. The Unit Area parameter must be a Type parameter.",
    #                 ok=True,
    #                 warn_icon=True, exitscript=True)
    #
    # gm_dict1 = {p.Definition.Name: p for p in gm_params_text}
    # gm_dict2 = {p.Definition.Name: p for p in gm_params_area}
    # construct rwp UI
    # components = [
    #     # Label("[Department] Match Room parameters:"),
    #     # ComboBox(name="room_combobox1", options=room_params_text, default="Department"),
    #     # Label("[Description] to Generic Model parameters:"),
    #     # ComboBox("gm_combobox1", gm_dict1, default="Description"),
    #     Label("[Unit Area] parameter:"),
    #     ComboBox("gm_combobox2", gm_dict2),
    #     Button("Select")]
    # form = FlexForm("Match parameters", components)
    # ok = form.show()
    # if ok:
    #     # assign chosen parameters
    #     # chosen_room_param1 = form.values["room_combobox1"]
    #     # chosen_gm_param1 = form.values["gm_combobox1"]
    #     chosen_gm_param2 = form.values["gm_combobox2"]
    # else:
    #     sys.exit()


    # iterate through rooms
    for room in selection:
        # helper: define inverted transform to translate room geometry to origin
        geo_translation = helper.inverted_transform(room)
        # collect room boundaries and translate them to origin
        room_boundaries = helper.room_bound_to_origin(room, geo_translation)

        # define new family doc
        try:
            new_family_doc = revit.doc.Application.NewFamilyDocument(fam_template_path)
        except:
            forms.alert(msg="No Template",
                        sub_msg="There is no Generic Model Template",
                        ok=True,
                        warn_icon=True, exitscript=True)

        # Name the Family ( Proj_Ten_Type_Name : Ten_Type)
        # get values of selected Room parameters and replace with default values if empty:
        # Project Number:
        project_number = revit.doc.ProjectInformation.Number
        if not project_number:
            project_number = "H&P"
        # Department:
        # dept = room.LookupParameter("Department").AsString()  # chosen parameter for Department
        dept = "Flat"  # chosen parameter for Department
        # if not dept:
        #     dept = "Unit"
        # Room name:
        room_name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()

        # Room area:
        unit_area = room.get_Parameter(DB.BuiltInParameter.ROOM_AREA).AsDouble()
        if unit_area == 0:
            continue

        # Room number (to be used as layout type differentiation)
        room_number = room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString()

        # construct family and family type names:
        fam_name = str(dept) + "_" + room_name + "_" + room_number
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
        new_family_doc.SaveAs(fam_path, saveas_opt)

        # Load Family into project
        with revit.Transaction("Load Family", revit.doc):
            loaded_f = revit.db.create.load_family(fam_path, doc=revit.doc)
            revit.doc.Regenerate()

        # Create extrusion from room boundaries
        with revit.Transaction(doc=new_family_doc, name="Create Extrusion"):
            try:
                extrusion = helper.room_to_extrusion(room, new_family_doc,output)
                helper.assign_material_param(extrusion, sp_unit_material, new_family_doc)
                placement_point = room.Location.Point
            except Exception as err:
                    logger.error(err)
                    continue


            # parameter definition
            external_parameter_definition = get_external_definition_by_name("Unit Area")
            builtin_param_group = DB.BuiltInParameterGroup.PG_TEXT
            is_instance_parameter = False

            # create new family parameter
            unit_area_parameter = new_family_doc.FamilyManager.AddParameter(external_parameter_definition,
                                                      builtin_param_group,
                                                      is_instance_parameter)

            unit_area_instance_parameter = new_family_doc.FamilyManager.AddParameter(
                get_external_definition_by_name("Unit Area Instance"),
                builtin_param_group,
                True
            )

            if unit_area_instance_parameter.CanAssignFormula:
                new_family_doc.FamilyManager.NewType("new")
                new_family_doc.FamilyManager.SetFormula(unit_area_instance_parameter, "Unit Area")

        # save and close family
        save_opt = DB.SaveOptions()
        new_family_doc.Save(save_opt)
        new_family_doc.Close()

        # Reload family with extrusion and place it in the same position as the room
        with revit.Transaction("Reload Family", revit.doc):
            try:
                loaded_f = revit.db.create.load_family(fam_path, doc=revit.doc)
                # find family symbol and activate
                fam_symbol = helper.get_fam(fam_name)
                if not fam_symbol.IsActive:
                    fam_symbol.Activate()
                    revit.doc.Regenerate()

                fam_symbol.LookupParameter("Description").Set(dept)
                fam_symbol.LookupParameter("Unit Area").Set(unit_area)

                # place family symbol at position
                new_fam_instance = revit.doc.Create.NewFamilyInstance(placement_point, fam_symbol,
                                                                      room.Level,
                                                                      DB.Structure.StructuralType.NonStructural)
                # correct level offset
                correct_lvl_offset = new_fam_instance.get_Parameter(
                    DB.BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(0)
                print(
                    "Created and placed family instance : {1} - {2} {0} ".format(
                        output.linkify(new_fam_instance.Id),
                        fam_name, fam_type_name))
            except Exception as err:
                logger.error(err)

