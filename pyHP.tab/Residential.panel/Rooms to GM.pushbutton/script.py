__title__ = "Room to\n Generic Model"
__doc__ = "Transforms rooms into Generic Model families. Carries over unit parameters"

from pyrevit import revit, DB, script, forms, HOST_APP
from rpw.ui.forms import (FlexForm, Label, ComboBox, Separator, Button)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit import Exceptions
import tempfile
import rpw
import helper

# use preselected elements, filtering rooms only
pre_selection = helper.preselection_with_filter("Rooms")
# or select rooms
if pre_selection:
    selection = pre_selection
else:
    selection = helper.select_rooms_filter()

# pick material to use
all_mat = DB.FilteredElementCollector(revit.doc).OfClass(DB.Material).ToElements()
for mat in all_mat:
    if mat.Name == "Flat 1 Bed Private":
        chosen_mat = mat
    else:
        pass

if selection:
    # Create family doc from template
    # get file template from location
    fam_template_path = "C:\ProgramData\Autodesk\RVT " + \
                        HOST_APP.version + "\Family Templates\English\Metric Generic Model.rft"

    # format parameters for UI
    # gather and organize Room parameters: (only editable text params)
    room_parameter_set = selection[0].Parameters
    room_params_text = [p.Definition.Name for p in room_parameter_set if
                        p.StorageType.ToString() == "String" and p.IsReadOnly == False]
    # collect and organize Generic Model parameters: (only editable text type params)
    gm_parameter_set = helper.param_set_by_cat(DB.BuiltInCategory.OST_GenericModel)
    gm_params_text = [p for p in gm_parameter_set if p.StorageType.ToString() == "String"]
    gm_dict1 = {p.Definition.Name: p for p in gm_params_text}

    # construct rwp UI
    components = [
        Label("[Department] Match Room parameters:"),
        ComboBox(name="room_combobox1", options=room_params_text, default="Department"),
        Label("[Department] to Generic Model parameters:"),
        ComboBox("gm_combobox1", gm_dict1, default="Description"),
        Separator(),
        Label("[Unit Type] Match Room parameters:"),
        ComboBox(name="room_combobox2", options=room_params_text, default="Room_Unit Type"),
        Label("[Unit Type] to Generic Model parameters:"),
        ComboBox("gm_combobox2", gm_dict1, default="Unit Type"),
        Separator(),
        Label("[Tenure] Match Room parameters:"),
        ComboBox(name="room_combobox3", options=room_params_text, default="Room_Unit Tenure"),
        Label("[Tenure] to Generic Model parameters:"),
        ComboBox("gm_combobox3", gm_dict1, default="Tenure"),
        Button("Select")]
    form = FlexForm("Match parameters", components)
    form.show()
    # assign chosen parameters
    chosen_room_param1 = form.values["room_combobox1"]
    chosen_gm_param1 = form.values["gm_combobox1"]
    chosen_room_param2 = form.values["room_combobox2"]
    chosen_gm_param2 = form.values["gm_combobox2"]
    chosen_room_param3 = form.values["room_combobox3"]
    chosen_gm_param3 = form.values["gm_combobox3"]

    # iterate through rooms
    for room in selection:
        # helper: define inverted transform to translate room geometry to origin
        geo_translation = helper.inverted_transform(room)
        # collect room boundaries and translate them to origin
        room_boundaries = helper.room_bound_to_origin(room, geo_translation)

        # get shared parameter for the extrusion material
        sp_unit_material = helper.get_shared_param_by_name_type("Unit Material", DB.ParameterType.Material)
        # define new family doc
        new_family_doc = revit.doc.Application.NewFamilyDocument(fam_template_path)
        # Name the Family
        project_number = revit.doc.ProjectInformation.Number

        # get values of selected Room parameters and replace with default values if empty:
        # Project Number:
        if not project_number:
            project_number = "H&P"
        # Department:
        dept = room.LookupParameter(chosen_room_param1).AsString()  # chosen parameter for Department
        if not dept:
            dept = "Unit"
        # Room name:
        try:
            room_name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
        finally:
            room_name = str(room.Id)
        # Tenure
        try:
            room_tenure = room.LookupParameter(chosen_room_param3).AsString()
        finally:
            room_tenure = "TN"
        if not room_tenure:
            room_tenure = "TN"
        # Room Unit Type (nr. bedrooms and bed spaces)
        try:
            room_type_instance = room.LookupParameter(chosen_room_param2).AsString()
        except:
            room_type_instance = "nBnP"
        if not room_type_instance:
            room_type_instance = "nBnP"
        # Room number (to be used as layout type differentiation)
        room_number = room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString()
        if not room_number:
            room_number = str(room.Id)

        # construct family and family type names:
        fam_name = project_number + "_" + str(dept) + "_" + room_name
        fam_type_name = room_tenure + "_" + room_type_instance + "_" + room_number

        # Save family in temp folder
        fam_path = tempfile.gettempdir() + "\ " + fam_name + ".rfa"
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
                extrusion_height = helper.convert_to_internal(2500)
                ref_plane = helper.get_ref_lvl_plane (new_family_doc)
                # create extrusion, assign material, associate with shared parameter
                extrusion = new_family_doc.FamilyCreate.NewExtrusion(True, room_boundaries, ref_plane[0], extrusion_height)
                ext_mat_param = extrusion.get_Parameter(DB.BuiltInParameter.MATERIAL_ID_PARAM)
                new_mat_param = new_family_doc.FamilyManager.AddParameter(sp_unit_material,DB.BuiltInParameterGroup.PG_MATERIALS, False)
                new_family_doc.FamilyManager.AssociateElementParameterToFamilyParameter(ext_mat_param, new_mat_param)
            finally:
                pass

        # save and close family
        save_opt = DB.SaveOptions()
        new_family_doc.Save(save_opt)
        new_family_doc.Close()

        # Reload family with extrusion and place it in the same position as the room
        with revit.Transaction("Reload Family", revit.doc):
            loaded_f = revit.db.create.load_family(fam_path, doc=revit.doc)
            revit.doc.Regenerate()
            str_type = DB.Structure.StructuralType.NonStructural
            # find family symbol and activate
            fam_symbol = None
            get_fam = DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol).OfCategory(
                DB.BuiltInCategory.OST_GenericModel).WhereElementIsElementType().ToElements()
            for fam in get_fam:
                type_name = fam.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                if str.strip(type_name) == fam_name:
                    fam_symbol = fam
                    fam_symbol.Name = fam_type_name
                    # set type parameters
                    fam_symbol.LookupParameter(chosen_gm_param1.Definition.Name).Set(dept)
                    fam_symbol.LookupParameter(chosen_gm_param2.Definition.Name).Set(room_type_instance)
                    fam_symbol.LookupParameter(chosen_gm_param3.Definition.Name).Set(room_tenure)
                    fam_symbol.LookupParameter("Unit Material").Set(chosen_mat.Id)
                    if not fam_symbol.IsActive:
                        fam_symbol.Activate()
                        revit.doc.Regenerate()

                    # place family symbol at postision
                    new_fam_instance = revit.doc.Create.NewFamilyInstance(room.Location.Point, fam_symbol, room.Level, str_type)
                    correct_lvl_offset = new_fam_instance.get_Parameter(DB.BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(0)

