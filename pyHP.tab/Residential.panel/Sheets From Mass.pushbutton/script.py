from pyrevit import revit, DB, script, forms
from rpw.ui.forms import FlexForm, Label, TextBox, Button, ComboBox, CheckBox, Separator
from pyHP import select, database, geo, units
import sys
import ui, locator
from pyrevit.framework import List

# prerequisites

ui = ui.UI(script)
output = script.get_output()

# first collect all params of Mass category, instance, String
mass_param_dict = database.param_dict_by_cat(DB.BuiltInCategory.OST_Mass, is_instance_param=True, storage_type = "String")
ui.mass_param_dict = mass_param_dict
ui.set_massparam()
# ask which parameter is used for Layout Type (ex. "1B2P - Type A" or "Flat Type D")


components1 = [
    Label("Which parameter stores the Unit Type Name?"),
    ComboBox(name="param", options=sorted(mass_param_dict.values()), default=ui.massparam),
    Separator(),
    Button("Select")
]
form1 = FlexForm("Which parameter to use?", components1)
ok1 = form1.show()

if ok1:
    chosen_massparam_name = form1.values["param"]
else:
    sys.exit()

ui.set_config("massparam", chosen_massparam_name)

# using the param name (dict value), get the parameter itself
for k, v in mass_param_dict.items():
    if v == chosen_massparam_name:
        chosen_massparam = k

# prep and fill in default values for the next UI
viewplans = DB.FilteredElementCollector(revit.doc).OfClass(DB.ViewPlan)  # collect view plans
ui.vt_layout_dict \
    = {v.Name: v for v in viewplans if v.IsTemplate}  # only fetch IsTemplate plans
ui.vt_layout_dict["<None>"] = None # add None as a choice
# use a special collector w viewport param
ui.viewport_dict = {database.get_name(v): v for v in
                    database.get_viewport_types(revit.doc)}
# collect titleblocks (there may be a better method in pyrevit lib)
titleblocks = DB.FilteredElementCollector(revit.doc).OfCategory(
    DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
if not titleblocks:
    forms.alert("There are no Titleblocks loaded in the model.", exitscript=True)
ui.titleblock_dict = {'{} : {}'.format(tb.FamilyName, revit.query.get_name(tb)): tb for tb in titleblocks}

# TODO: remember last choice
all_schedules = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_Schedules).WhereElementIsNotElementType()
schedules = [s for s in all_schedules if "<Revision Schedule>" not in str(s.Title)]
sh_dict = {revit.query.get_name(sh): sh for sh in schedules}
ui.schedule_dict = sh_dict
ui.set_titleblocks()
ui.set_vp_types()
ui.set_viewtemplates()
ui.set_schedules()
cat = DB.BuiltInCategory.OST_Mass
fl_plan_type = database.get_view_family_types(DB.ViewFamily.FloorPlan, revit.doc)[0]



# select masses
selection = select.select_with_cat_filter(cat, "Select a Mass")

unique_types = {}
for m in selection:
    v = m.get_Parameter(chosen_massparam).AsString()

    if v and v not in unique_types.keys():
        unique_types[v] = m

def_tb = database.tb_name_match(ui.titleblock)
print (def_tb)
components2 = [
    Label("Select Titleblock"),
    ComboBox(name="tb", options=sorted(ui.titleblock_dict), default=database.tb_name_match(ui.titleblock)),
    Label("Sheet Number"),
    TextBox("sheet_number", Text=ui.sheet_number),
    Label("Crop offset [mm]"),
    TextBox("crop_offset", Text=str(ui.crop_offset)),
    Separator(),
    Label("View Template for Layout"),
    ComboBox(name="vt_layout", options=sorted(ui.vt_layout_dict),
             default=database.vt_name_match(ui.viewplan)),
    Label("View Template for Key Plans"),
    ComboBox(name="vt_keyplan", options=sorted(ui.vt_layout_dict),
             default=database.vt_name_match(ui.viewkeyplan)),
    Label("Viewport Type"),
    ComboBox(name="vp_types", options=sorted(ui.viewport_dict), default=database.vp_name_match(ui.viewport)),
    Separator(),
    Label("Area Schedule Template"),
    ComboBox(name="area_sh", options=sorted(sh_dict), default=database.sh_name_match(ui.schedule)),

    Separator(),
    Button("Select"),
]

form2 = FlexForm("Settings", components2)
ok2 = form2.show()

if ok2:
    # match the variables with user input
    chosen_sheet_nr = form2.values["sheet_number"]
    chosen_vt_layout = ui.vt_layout_dict[form2.values["vt_layout"]]
    chosen_vt_keyplan = ui.vt_layout_dict[form2.values["vt_keyplan"]]
    chosen_tb = ui.titleblock_dict[form2.values["tb"]]
    chosen_vp_type = ui.viewport_dict[form2.values["vp_types"]]
    chosen_crop_offset = units.correct_input_units(form2.values["crop_offset"], revit.doc)
    chosen_area_sh = ui.schedule_dict[form2.values["area_sh"]]
else:
    sys.exit()

ui.set_config("sheet_number", chosen_sheet_nr)
ui.set_config("crop_offset", form2.values["crop_offset"])
ui.set_config("titleblock", form2.values["tb"])
ui.set_config("viewplan", form2.values["vt_layout"])
ui.set_config("viewkeyplan", form2.values["vt_keyplan"])
ui.set_config("schedule", form2.values["area_sh"])

sys.exit()
all_mass = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_Mass).WhereElementIsNotElementType().ToElements()

all_view_filters = DB.FilteredElementCollector(revit.doc).OfClass(DB.FilterElement).ToElements()
overrides = DB.OverrideGraphicSettings()
overrides.SetSurfaceTransparency(1)
filter_cats = List[DB.ElementId](DB.ElementId(cat) for cat in [DB.BuiltInCategory.OST_Mass])

def create_filter_from_rules(rules):
    elem_filters = List[DB.ElementFilter]()
    for rule in rules:
        elem_param_filter = DB.ElementParameterFilter(rule)
        elem_filters.Add(elem_param_filter)
    el_filter = DB.LogicalAndFilter(elem_filters)
    return el_filter

with revit.Transaction("Create Flat Type Sheets", revit.doc):
    for layout_type_name in unique_types:
        fam_instance = unique_types[layout_type_name]
        level = fam_instance.Host
        layout_plan = DB.ViewPlan.Create(revit.doc, fl_plan_type.Id, level.Id)

        # find all the levels with this element
        all_levels_with_same_layout = []
        for mass in all_mass:
            if mass.get_Parameter(chosen_massparam).AsString() == layout_type_name:
                host = mass.Host
                if isinstance(host, DB.Level):
                    all_levels_with_same_layout.append(host.Id)

        layout_filter_id = None
        filter_name = "Mass_"+ layout_type_name
        for f in all_view_filters:
            if str(f.Name) == str(filter_name):
                layout_filter_id = f.Id

        key_plans = {}
        for lvl_id in set(all_levels_with_same_layout):
            kp = DB.ViewPlan.Create(revit.doc, fl_plan_type.Id, lvl_id)
            key_plans[kp] = lvl_id

        if not layout_filter_id:
            # if not found, create a new filter

            new_filter = DB.ParameterFilterElement.Create(revit.doc, "Mass - " + layout_type_name, filter_cats)
            # sp = get_shared_param(LAYOUT_PARAM_NAME)
            p = fam_instance.get_Parameter(chosen_massparam)
            equals_rule = DB.ParameterFilterRuleFactory.CreateEqualsRule(p.Id, layout_type_name, False)
            f_rules = List[DB.FilterRule]([equals_rule])

            filt = create_filter_from_rules(f_rules)

            new_filter.SetElementFilter(filt)

            layout_filter_id = new_filter.Id
        for kp in key_plans:
            kp.AddFilter(layout_filter_id)
            kp.SetFilterOverrides(layout_filter_id, overrides)

        layout_plan.CropBoxActive = True
        bb = fam_instance.get_BoundingBox(None)
        bb_outline = geo.get_bb_outline(bb)
        bb_loop = DB.CurveLoop()
        for line in bb_outline:
            bb_loop.Append(line)
        crsm_plan = layout_plan.GetCropRegionShapeManager()
        crsm_plan.SetCropShape(bb_loop)

        # rename the view
        layout_name = "Layout - " + layout_type_name
        layout_plan.Name = database.unique_view_name(layout_name, suffix=" Plan")

        for k in key_plans:
            keyplan_name = "Key Plan - " + layout_type_name + " - " + revit.doc.GetElement(key_plans[k]).Name
            k.Name = database.unique_view_name(keyplan_name, suffix=" Key Plan")
            database.apply_vt(k, chosen_vt_keyplan)

        # apply view template
        database.apply_vt(layout_plan, chosen_vt_layout)


        # duplicate template schedule
        area_schedule_id = chosen_area_sh.Duplicate(DB.ViewDuplicateOption.Duplicate)
        area_schedule = revit.doc.GetElement(area_schedule_id)


        def find_field(sch, param):
            definition = sch.Definition
            param_id = param.Id
            for field_id in definition.GetFieldOrder():
                found_field = definition.GetField(field_id)
                if found_field.ParameterId == param_id:
                    return found_field

        layout_name_param = fam_instance.get_Parameter(chosen_massparam)
        typename_field = find_field(area_schedule, layout_name_param)
        if not typename_field:
            typename_field = area_schedule.Definition.AddField(DB.ScheduleFieldType.Instance,
                                                               (layout_name_param.Id))

        field_id = typename_field.FieldId
        filter_type = DB.ScheduleFilterType.Contains.Equal
        # this needs to change to a more stable, Type parameter
        string_value_type = layout_name_param.AsString()
        flat_type_filter = DB.ScheduleFilter(field_id, filter_type, string_value_type)
        area_schedule.Definition.AddFilter(flat_type_filter)

        # rename schedule
        td = area_schedule.GetTableData()
        td = area_schedule.GetTableData()
        tds = td.GetSectionData(DB.SectionType.Header)
        text = tds.GetCellText(0, 0)
        tds.SetCellText(0, 0, database.unique_view_name(string_value_type, suffix=" Area Schedule"))
        area_schedule.Name = database.unique_schedule_name(string_value_type, suffix=" Mass Schedule")
        # create sheet
        sheet = database.create_sheet(chosen_sheet_nr, layout_name, chosen_tb.Id)

        # get positions on sheet
        loc = locator.Locator(sheet, 5.1, 'Vertical', 'Tiles')
        layout_position = loc.plan
        keyplan_position = loc.rcp

        # place view on sheet
        place_layout = DB.Viewport.Create(revit.doc, sheet.Id, layout_plan.Id, layout_position)
        place_area_sh = DB.ScheduleSheetInstance.Create(revit.doc, sheet.Id, area_schedule.Id, layout_position)
        place_keyplan = [DB.Viewport.Create(revit.doc, sheet.Id, kp.Id, keyplan_position) for kp in key_plans]
        # for kp in key_plans:
        #     place_keyplan = DB.Viewport.Create(revit.doc, sheet.Id, kp.Id, keyplan_position)
        #     kp.ChangeTypeId(chosen_vp_type.Id)

        # new: change viewport types
        for vp in [place_layout] + place_keyplan:
            vp.ChangeTypeId(chosen_vp_type.Id)

        revit.doc.Regenerate()

        # realign the viewports to their desired positions
        loc.realign_pos(revit.doc, [place_layout], [layout_position])
        loc.realign_pos(revit.doc, place_keyplan, [keyplan_position])
        # loc.realign_pos(revit.doc, [place_area_sh], [keyplan_position])


        print("Sheet : {0} \t Layout {1} ".format(output.linkify(sheet.Id), layout_type_name))
