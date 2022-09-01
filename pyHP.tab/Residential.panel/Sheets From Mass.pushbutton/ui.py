class UI:
    def __init__(self, script):
        self.script = script
        self.config = script.get_config()
        self.titleblock_dict = {}
        self.view_temp_dict = {}
        self.viewport_dict = {}
        self.massparam_dict = {}
        self.schedule_dict = {}
        self.sheet_number = self.config.get_option('sheet_number', '1000')
        self.crop_offset = self.config.get_option('crop_offset', '350')
        self.massparam = None

    def set_massparam(self):
        self.massparam = self.config.get_option('massparam', list(self.massparam_dict.keys())[0])
        # self.massparam = {0:1}

    def set_viewtemplates(self):
        self.viewplan = self.config.get_option('viewplan', list(self.view_temp_dict.keys())[0])
        self.viewkeyplan = self.config.get_option('viewkeyplan', list(self.view_temp_dict.keys())[0])

    def set_titleblocks(self):
        # set default
        self.titleblock = self.config.get_option('titleblock', list(self.titleblock_dict.keys())[0])

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
        if var == "viewport":
            self.config.viewport = val
        if var == "massparam":
            self.config.massparam = val
        if var == "schedule":
            self.config.schedule = val

        self.script.save_config()
