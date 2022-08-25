class UI:
    def __init__(self, script):
        self.script = script
        self.config = script.get_config()
        self.titleblock_dict = {}
        self.vt_layout_dict = {}
        # self.viewsection_dict = {}
        self.viewport_dict = {}
        self.massparam_dict = {}
        self.schedule_dict = {}
        self.sheet_number = self.config.get_option('sheet_number', '1000')
        self.crop_offset = self.config.get_option('crop_offset', '350')
        # self.massparam = None
        # self.massparam = self.config.get_option('massparam', self.massparam_dict)
        # self.viewport = self.config.get_option('viewport', self.viewport_dict)
        # self.titleblock = self.config.get_option('titleblock', self.titleblock_dict)
        # self.schedule_dict = self.config.get_option('schedule', self.schedule_dict)

    def set_massparam(self):
        self.massparam = self.config.get_option('massparam', None)

    def set_viewtemplates(self):
        self.viewplan = self.config.get_option('viewplan', "<None>")
        self.viewkeyplan = self.config.get_option('viewkeyplan', "<None>")

    def set_titleblocks(self):
        # set default
        self.titleblock = self.config.get_option('titleblock', None)

    def set_vp_types(self):
        # set default
        self.viewport = self.config.get_option('viewport', list(self.viewport_dict.keys())[0])

    def set_schedules(self):
        # set default
        self.schedule = self.config.get_option('schedule', list(self.schedule_dict.keys())[0])

    def set_config(self, var, val):
        if var == "sheet_number":
            self.config.sheet_number = val
        if var == "crop_offset":
            self.config.crop_offset = val
        if var == "titleblock":
            self.config.titleblock = val
        if var == "viewplan":
            self.config.viewplan = val
        if var == "viewkeyplan":
            self.config.viewkeyplan = val
        if var == "viewsection":
            self.config.viewsection = val
        if var == "viewport":
            self.config.viewport = val
        if var == "massparam":
            self.config.massparam = val
        if var == "schedule":
            self.config.schedule = val

        self.script.save_config()
