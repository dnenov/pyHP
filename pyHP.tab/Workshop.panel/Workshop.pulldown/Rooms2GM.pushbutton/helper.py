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

def convert_length_to_internal(from_units):
    # convert length units from project  to internal
    d_units = DB.Document.GetUnits(revit.doc).GetFormatOptions(DB.UnitType.UT_Length).DisplayUnits
    converted = DB.UnitUtils.ConvertToInternalUnits(from_units, d_units)
    return converted


def get_shared_param_by_name_type(sp_name, sp_type):
    # query shared parameters file and return the desired parameter by name and parameter type
    # will return first result
    spf = revit.doc.Application.OpenSharedParameterFile()
    try:
        for def_group in spf.Groups:
            for sp in def_group.Definitions:
                if sp.Name == sp_name and sp.ParameterType == sp_type:
                    return sp
        if not sp:
            forms.alert("Shared parameter not found", ok=True, warn_icon=True)
            return None
    except:
        forms.alert("Shared parameter not found", ok=True, warn_icon=True )
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

def preselection_with_filter (cat_name):
    # use pre-selection of elements, but filter them by given category name
    pre_selection = []
    for id in rpw.revit.uidoc.Selection.GetElementIds():
        sel_el = revit.doc.GetElement(id)
        if sel_el.Category.Name == cat_name:
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


def point_equal_list(pt, lst):
    for el in list(lst):
        if pt.IsAlmostEqualTo(el):
            return el
    else:
        return None


def get_open_ends(curves_list):
    endpoints = []
    for curve in curves_list:
        for i in range(2):
            duplicate = point_equal_list(curve.GetEndPoint(i), endpoints)
            if duplicate:
                endpoints.remove(duplicate)
            else:
                endpoints.append(curve.GetEndPoint(i))
    if endpoints:
        return endpoints
    else:
        return None


def room_bound_to_origin (room, translation):
    # iterate through room boundaries and translate them close to the origin
    # also query open ends and return none if the loop is open
    room_boundaries = DB.CurveArrArray()
    # get room boundary segments
    room_segments = room.GetBoundarySegments(DB.SpatialElementBoundaryOptions())
    # iterate through loops of segments and add them to the array
    for seg_loop in room_segments:
        curve_loop = [s.GetCurve() for s in seg_loop]
        open_ends = get_open_ends(curve_loop)
        if open_ends:
            return None
        curve_array = DB.CurveArray()
        for old_curve in curve_loop:
            #old_curve = s.GetCurve()
            new_curve = old_curve.CreateTransformed(translation)  # move curves to origin
            curve_array.Append(new_curve)
        room_boundaries.Append(curve_array)
    return room_boundaries


def get_ref_lvl_plane (family_doc):
    # from given family doc, return Ref. Level reference plane
    find_planes = DB.FilteredElementCollector(family_doc).OfClass(DB.SketchPlane)
    return [plane for plane in find_planes if plane.Name == "Ref. Level"]



def get_fam(some_name):
    fam_name_filter = query.get_biparam_stringequals_filter({DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM: some_name})
    found_fam = DB.FilteredElementCollector(revit.doc) \
        .OfCategory(DB.BuiltInCategory.OST_GenericModel) \
        .WherePasses(fam_name_filter) \
        .WhereElementIsNotElementType().ToElements()

    return found_fam


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


