from pyrevit import revit, DB, script, forms
from rpw.ui.forms import FlexForm, Label, TextBox, Button, ComboBox, CheckBox, Separator
from pyHP import select, database, geo, units
import sys
import ui, locator


# prerequisites
ui = ui.UI(script)
output = script.get_output()
ui.viewport_dict = {database.get_name(v): v for v in
                    database.get_viewport_types(revit.doc)}  # use a special collector w viewport param

# add none as an option
ui.viewsection_dict["<None>"] = None
ui.vt_layout_dict["<None>"] = None
ui.set_viewtemplates()
viewplans = DB.FilteredElementCollector(revit.doc).OfClass(DB.ViewPlan)  # collect plans
ui.vt_layout_dict \
    = {v.Name: v for v in viewplans if v.IsTemplate}  # only fetch IsTemplate plans

cat = DB.BuiltInCategory.OST_Mass
fl_plan_type = database.get_view_family_types(DB.ViewFamily.FloorPlan, revit.doc)[0]
# collect titleblocks in a dictionary
titleblocks = DB.FilteredElementCollector(revit.doc).OfCategory(
    DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType()
ui.titleblock_dict = {'{}: {}'.format(tb.FamilyName, revit.query.get_name(tb)): tb for tb in titleblocks}
ui.set_titleblocks()

# TODO: remember last choice
all_schedules = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_Schedules).WhereElementIsNotElementType()
schedules = [s for s in all_schedules if "<Revision Schedule>" not in str(s.Title)]
sh_dict = {revit.query.get_name(sh): sh for sh in schedules}

# select masses
selection = select.select_with_cat_filter(cat, "Select a Mass")

LAYOUT_PARAM_NAME = "Unit Layout Name"
unique_types = {m.LookupParameter(LAYOUT_PARAM_NAME).AsString(): m for m in selection}

components = [
    Label("Select Titleblock"),
    ComboBox(name="tb", options=sorted(ui.titleblock_dict), default=database.tb_name_match(ui.titleblock, revit.doc)),
    Label("Sheet Number"),
    TextBox("sheet_number", Text=ui.sheet_number),
    Label("Crop offset [mm]"),
    TextBox("crop_offset", Text=str(ui.crop_offset)),
    Separator(),
    Label("View Template for Layout"),
    ComboBox(name="vt_layout", options=sorted(ui.vt_layout_dict),
             default=database.vt_name_match(ui.viewplan, revit.doc)),
    Label("View Template for Key Plans"),
    ComboBox(name="vt_keyplan", options=sorted(ui.vt_layout_dict),
             default=database.vt_name_match(ui.viewkeyplan, revit.doc)),
    Label("Viewport Type"),
    ComboBox(name="vp_types", options=sorted(ui.viewport_dict)),
    Separator(),
    Label("Area Schedule Template"),
    ComboBox(name="area_sh", options=sorted(sh_dict)),

    Separator(),
    Button("Select"),
]

form = FlexForm("Fill in parameters", components)
ok = form.show()

if ok:
    # match the variables with user input
    chosen_sheet_nr = form.values["sheet_number"]
    chosen_vt_layout = ui.vt_layout_dict[form.values["vt_layout"]]
    chosen_vt_keyplan = ui.vt_layout_dict[form.values["vt_keyplan"]]
    chosen_tb = ui.titleblock_dict[form.values["tb"]]
    chosen_vp_type = ui.viewport_dict[form.values["vp_types"]]
    chosen_crop_offset = units.correct_input_units(form.values["crop_offset"], revit.doc)
    chosen_area_sh = sh_dict[form.values["area_sh"]]
else:
    sys.exit()

ui.set_config("sheet_number", chosen_sheet_nr)
ui.set_config("crop_offset", form.values["crop_offset"])
ui.set_config("titleblock", form.values["tb"])
ui.set_config("viewplan", form.values["vt_layout"])
ui.set_config("viewkeyplan", form.values["vt_keyplan"])

all_mass = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_Mass).WhereElementIsNotElementType().ToElements()



with revit.Transaction("Create Flat Type Sheets", revit.doc):
    for layout_type_name in unique_types:
        fam_instance = unique_types[layout_type_name]
        level = fam_instance.Host
        layout_plan = DB.ViewPlan.Create(revit.doc, fl_plan_type.Id, level.Id)

        # find all the levels with this element
        all_levels_with_same_layout = []
        for mass in all_mass:
            if mass.LookupParameter(LAYOUT_PARAM_NAME).AsString() == layout_type_name:
                host = mass.Host
                if isinstance(host, DB.Level):
                    all_levels_with_same_layout.append(host)

        key_plans = {}
        for lvl in set(all_levels_with_same_layout):
            kp= DB.ViewPlan.Create(revit.doc, fl_plan_type.Id, lvl.Id)
            key_plans[kp] = lvl


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
            keyplan_name = "Key Plan - " + layout_type_name + " - " + key_plans[k].Name
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

        layout_name_param = fam_instance.LookupParameter(LAYOUT_PARAM_NAME)
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
