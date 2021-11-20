
class ppt_color:
    def __init__(self, r=255, g=255, b=255):
        self.r = r
        self.g = g
        self.b = b
    def __str__(self):
        return "%3d,%3d,%3d"%(self.r, self.g, self.b)

        
_default_slide_back_col = ppt_color(0,32,96) #ppt_color(255,255,255)
_default_font_col  = ppt_color(255,255,255)

_default_solid_fill_color     = ppt_color(100, 100, 100)
_default_gradient_fill_color1 = ppt_color(0, 0, 0)
_default_gradient_fill_color2 = ppt_color(155, 155, 155)
_color_black = ppt_color(0, 0, 0)
_color_white = ppt_color(255, 255, 255)
_color_red   = ppt_color(255, 0, 0)
_color_green = ppt_color(0, 255, 0)
_color_blue  = ppt_color(0, 0, 255)
_color_yellow= ppt_color(255, 255, 0)
