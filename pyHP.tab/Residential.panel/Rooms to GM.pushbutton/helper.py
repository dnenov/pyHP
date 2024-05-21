"""Transform rooms to generic model families"""

from pyrevit import revit, DB, script, forms, HOST_APP
from rpw.ui.forms import (FlexForm, Label, ComboBox, Separator, Button)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit import Exceptions
import tempfile
import rpw
from pyrevit.revit.db import query


# selection filter for rooms
class RoomsFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            if elem.Category.Name == "Rooms":
                return True
            else:
                return False
        except AttributeError:
            return False

    def AllowReference(self, reference, position):
        try:
            if isinstance(revit.doc.GetElement(reference), DB.Room):
                return True
            else:
                return False
        except AttributeError:
            return False


def select_rooms_filter():
    # select elements while applying category filter
    try:
        with forms.WarningBar(title="Pick Rooms to transform"):
            selection = [revit.doc.GetElement(reference) for reference in rpw.revit.uidoc.Selection.PickObjects(
                ObjectType.Element, RoomsFilter())]
            return selection
    except Exceptions.OperationCanceledException:
        forms.alert("Cancelled", ok=True, warn_icon=False)


def convert_length_to_internal(d_units):
    # convert length units from display units to internal
    units = revit.doc.GetUnits()
    if HOST_APP.is_newer_than(2021):
        internal_units = units.GetFormatOptions(DB.SpecTypeId.Length).GetUnitTypeId()
    else:  # pre-2021
        internal_units = units.GetFormatOptions(DB.UnitType.UT_Length).DisplayUnits
    converted = DB.UnitUtils.ConvertToInternalUnits(d_units, internal_units)
    return converted


def get_shared_param_by_name_type(sp_name, sp_type):
    # query shared parameters file and return the desired parameter by name and parameter type
    # will return first result
    spf = revit.doc.Application.OpenSharedParameterFile()
    try:
        for def_group in spf.Groups:
            for sp in def_group.Definitions:
                if HOST_APP.is_newer_than(2022):
                    # print("name: {}\n sp data type: {}\n sp type: {} ".format(sp_name, sp.GetDataType().TypeId, sp_type.TypeId))
                    if sp.Name == sp_name and sp.GetDataType().TypeId == sp_type.TypeId:
                        return sp
                else:
                    if sp.Name == sp_name and sp.ParameterType == sp_type:
                        return sp
        if not sp:
            forms.alert("Shared parameter not found", ok=True, warn_icon=True)
            return None
    except:
        forms.alert("Shared parameter not found", ok=True, warn_icon=True)
        return None


def param_set_by_cat(cat):
    # get all project type parameters of a given category
    # can be used to gather parameters for UI selection
    all_gm = DB.FilteredElementCollector(revit.doc).OfCategory(cat).WhereElementIsElementType().ToElements()
    parameter_set = []
    for gm in all_gm:
        params = gm.Parameters
        for p in params:
            if p not in parameter_set and p.IsReadOnly == False:
                parameter_set.append(p)
    return parameter_set


def preselection_with_filter(bic):
    # use pre-selection of elements, but filter them by given category name
    pre_selection = []
    for id in rpw.revit.uidoc.Selection.GetElementIds():
        sel_el = revit.doc.GetElement(id)
        if sel_el.Category.Id.IntegerValue == int(bic):
            pre_selection.append(sel_el)
    return pre_selection


def inverted_transform(element):
    # get element location and return its inverted transform
    # can be used to translate geometry to 0,0,0 origin to recreate geometry inside a family
    el_location = element.Location.Point
    bb = element.get_BoundingBox(revit.doc.ActiveView)
    orig_cs = bb.Transform
    # Creates a transform that represents a translation via the specified vector.
    translated_cs = orig_cs.CreateTranslation(el_location)
    # Transform from the room location point to origin
    return translated_cs.Inverse


def room_bound_to_origin(room, translation):
    room_boundaries = DB.CurveArrArray()
    # get room boundary segments
    room_segments = room.GetBoundarySegments(DB.SpatialElementBoundaryOptions())
    # iterate through loops of segments and add them to the array
    for seg_loop in room_segments:
        curve_array = DB.CurveArray()
        for s in seg_loop:
            old_curve = s.GetCurve()
            new_curve = old_curve.CreateTransformed(translation)  # move curves to origin
            curve_array.Append(new_curve)
        room_boundaries.Append(curve_array)
    return room_boundaries


def get_ref_lvl_plane(family_doc):
    # from given family doc, return Ref. Level reference plane
    find_planes = DB.FilteredElementCollector(family_doc).OfClass(DB.SketchPlane)
    return [plane for plane in find_planes if plane.Name == "Ref. Level"]


def get_fam(family_name):
    fam_bip_id = DB.ElementId(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
    fam_bip_provider = DB.ParameterValueProvider(fam_bip_id)
    if HOST_APP.is_newer_than(2022):
        fam_filter_rule = DB.FilterStringRule(fam_bip_provider, DB.FilterStringEquals(), family_name)
    else:
        fam_filter_rule = DB.FilterStringRule(fam_bip_provider, DB.FilterStringEquals(), family_name, True)
    fam_filter = DB.ElementParameterFilter(fam_filter_rule)

    collector = DB.FilteredElementCollector(revit.doc) \
        .WherePasses(fam_filter) \
        .WhereElementIsElementType() \
        .FirstElement()

    return collector


def get_family_slow_way(name):
    # look for loaded families, workaround to deal with extra space
    get_loaded = DB.FilteredElementCollector(revit.doc) \
        .OfCategory(DB.BuiltInCategory.OST_GenericModel) \
        .WhereElementIsNotElementType() \
        .ToElements()
    if get_loaded:
        for el in get_loaded:
            el_f_name = el.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
            if el_f_name.strip(" ") == name:
                return el


def room_to_extrusion(r, family_doc,output):
    # room_height = r.get_Parameter(DB.BuiltInParameter.ROOM_HEIGHT).AsDouble()
    room_height = convert_length_to_internal(2500)
    # helper: define inverted transform to translate room geometry to origin
    geo_translation = inverted_transform(r)
    # collect room boundaries and translate them to origin
    room_boundaries = room_bound_to_origin(r, geo_translation)
    # skip if the boundaries are not a closed loop (can happen with misaligned boundaries)
    if not room_boundaries:
        print("Extrusion failed for room {}. Try fixing room boundaries".format(output.linkify(r.Id)))
        return

    ref_plane = get_ref_lvl_plane(family_doc)
    # create extrusion, assign material, associate with shared parameter
    try:
        extrusion = family_doc.FamilyCreate.NewExtrusion(True, room_boundaries, ref_plane[0],
                                                         room_height)

        return extrusion

    except Exceptions.InternalException:
        print("Extrusion failed for room {}. Try fixing room boundaries".format(output.linkify(r.Id)))
        return


def assign_material_param(extrusion, mat_param, fam_doc):
    ext_mat_param = extrusion.get_Parameter(DB.BuiltInParameter.MATERIAL_ID_PARAM)
    # create and associate a material parameter
    new_mat_param = fam_doc.FamilyManager.AddParameter(mat_param,
                                                              DB.BuiltInParameterGroup.PG_MATERIALS, False)
    fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(ext_mat_param,
                                                                        new_mat_param)
    return extrusion
