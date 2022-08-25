class UI:
    def __init__(self, script):
        self.script = script
        self.config = script.get_config()
        self.titleblock_dict = {}
        self.vt_layout_dict = {}
        self.viewsection_dict = {}
        self.viewport_dict = {}
        self.massparam_dict = {}
        self.titleblock = None

        # self.tblock_orientation = ['Vertical', 'Horizontal']
        # self.layout_orientation = ['Tiles', 'Cross']
        self.sheet_number = self.config.get_option('sheet_number', '1000')
        self.crop_offset = self.config.get_option('crop_offset', '350')
        self.massparam = self.config.get_option('massparam', self.massparam_dict)
        # self.titleblock_offset = self.config.get_option('titleblock_offset', '165')
        # self.titleblock_orientation = self.config.get_option('titleblock_orientation', self.tblock_orientation[0])
        # self.layout_ori = self.config.get_option('layout_ori', self.layout_orientation[0])
        # self.rotated_elevations = self.config.get_option('rotated_elevations', False)
        # self.el_as_sec = self.config.get_option('el_as_sec', False)

    def set_titleblocks(self):        
        self.titleblock = self.config.get_option('titleblock', list(self.titleblock_dict.keys())[0])

    def set_viewtemplates(self):
        self.viewplan = self.config.get_option('viewplan', "<None>")
        self.viewkeyplan = self.config.get_option('viewkeyplan', "<None>")
        self.viewsection = self.config.get_option('viewsection', "<None>")

    def set_vp_types(self):
        self.viewport = self.config.get_option('viewport', list(self.viewport_dict.keys())[0])

    def set_config(self, var, val):
        if var == "sheet_number":
            self.config.sheet_number = val
        if var == "crop_offset":
            self.config.crop_offset = val
        # if var == "titleblock_offset":
        #     self.config.titleblock_offset = val
        # if var == "titleblock_orientation":
        #     self.config.titleblock_orientation = val
        # if var == "layout_orientation":
        #     self.config.layout_ori = val
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
            

        self.script.save_config()