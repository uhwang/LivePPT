import re
import livepptfunc as func
import livepptcolor as color
import livepptconst as const
from livepptfunc import get_slide_size, get_textalign
import livepptconst as const
import msoLine, msoDash, msoShadow

class ppt_outlinetext_info:
    def __init__(self, col=color._color_black,
                    style=msoLine.msoLineSingle,
                    weight=const._default_outline_weight):
        self.show = True
        self.col = col
        self.style = style
        self.dash = msoDash.msoLineSolid
        self.transprancy = 0.0
        self.weight = weight

    def __str__(self):
        return "Show  : %s\nColor : %s\nStyle : %s\nDash  : %s\nTransp: %f\nWeight: %f"%\
                (str(self.show), str(self.col), 
                msoLine.get_linestyle_name(self.style),
                msoDash.get_dashstyle_name(self.dash),
                self.transprancy,
                self.weight)

# <a:effectLst>
#  <a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0">
#   <a:prstClr val="black">
#    <a:alpha val="40000"/>
#   </a:prstClr>
#  </a:outerShdw>
# </a:effectLst>

class ppt_shadow_info():
    def __init__(self):
        self.Visible = True
        self.Style = msoShadow._default_shadow_type # Outer: 2
        self.OffsetX = 2
        self.OffsetY = 2
        self.Blur = 2
        self.Transparency = 0.7

    def __str__(self):
        return "Style : %s\nOff(x): %d\nOff(y): %d\nBlur  : %d\nTransp: %f"%(
        msoShadow.get_shadowstyle_name(self.Style),
        self.OffsetX,self.OffsetY,self.Blur,self.Transparency)

    def get_style_index(self):
        return self.Style-1

    def set_style_index(self, s):
        self.Style = s+1

class ppt_fx_info():
    def __init__(self):
        self.show_outline = True
        self.show_shadow = True
        self.shadow = ppt_shadow_info()
        self.outline = ppt_outlinetext_info()

    def __str__(self):
        return "Outline: %s\n%s\nShadow: %s\n%s"%(
        str(self.show_outline),
        str(self.outline),
        str(self.show_shadow),
        str(self.shadow))

class ppt_textbox_fill_info:
    def __init__(self):
        self.show = True
        self.type = const._TEXTBOX_SOLID_FILL
        self.solid_col = color._default_slide_back_col
        self.gradient_col1 = color._default_gradient_fill_color1
        self.gradient_col2 = color._default_gradient_fill_color2

    def ftype(self): 
        return "Solid" if self.type is const._TEXTBOX_SOLID_FILL else "Gradient"

    def __str__(self):
        return 'Fill  : %s\nType  : %s\nS.Col : %s\nG.Col1: %s\nG.Col2: %s'%(
                str(self.show),
                self.ftype(), 
                str(self.solid_col), 
                str(self.gradient_col1),
                str(self.gradient_col2))

class ppt_textbox_info:
    def __init__(self, sx=const._default_txt_sx,
                       sy=const._default_txt_sy,
                       wid=const._default_txt_wid,
                       hgt=const._default_txt_hgt,
                       font_size = const._default_font_size):
        self.sx            = float(sx)
        self.sy            = float(sy)
        self.wid           = float(wid)
        self.hgt           = float(hgt)
        self.font_name     = const._default_font_name
        self.font_col      = color.ppt_color()
        self.font_bold     = True
        self.font_size     = font_size
        self.word_wrap     = False
        self.left_margin   = const._default_textbox_left_margin
        self.top_margin    = const._default_textbox_top_margin
        self.right_margin  = const._default_textbox_right_margin
        self.bottom_margin = const._default_textbox_bottom_margin
        self.fill          = ppt_textbox_fill_info()
        self.vanchor       = const._default_textframe_vanchor_index # middle : 1
        self.pp_align      = const._default_pragraph_align_index # center : 0
        self.fx            = ppt_fx_info()

    def __str__(self):
        return 'Sx   : %2.4f\nSy   : %2.4f\nWid  : %2.4f\nHgt  : %2.4f\n\n'\
            'Font : %s\nColor: %s\nSize : %2.4f\nWrap : %s\nVAlign: %s\n\n'\
            'HAlign: %s\n\nLeft  : %2.2f\nTop   : %2.2f\nRight : %2.2f\nBottom: %2.2f\n\n%s'%(
            self.sx, self.sy, self.wid, self.hgt, self.font_name, 
            str(self.font_col), self.font_size, str(self.word_wrap), 
            get_textalign(self.vanchor), get_textalign(self.pp_align), 
            self.left_margin, self.top_margin, self.right_margin, 
            self.bottom_margin, str(self.fill))

class ppt_slide_info:
    def __init__(self, slide_size_index):

        w, h = get_slide_size(const.get_slide_size_string(slide_size_index))
        self.back_col = color._default_slide_back_col
        self.wid = w
        self.hgt = h
        self.size_index = slide_size_index
        self.skip_image = True
        self.deep_copy = True

    def __str__(self):
        return "Wid  : %2.4f\nHgt  : %2.4f\nColor: %s\nDCopy: %s"%(\
            self.wid,self.hgt,str(self.back_col),str(self.deep_copy))

class ppt_slide_data():
    def __init__(self, slide_size_index):
        w, h = func.get_slide_size(const.get_slide_size_string(slide_size_index))
        self.slide = ppt_slide_info(slide_size_index)
        self.textbox = ppt_textbox_info(0,0,w,h,const._default_font_size)
        
    def __str__(self):
        return "%s\n%s\n"%(str(self.slide), str(self.textbox))
        
# use max size: sx=0, sy=0, wid=max wid, hgt=max hgt
#class ppt_hymal_info(ppt_slide_info, ppt_textbox_info):
class ppt_hymal_data():
    def __init__(self, slide_size_index):
        w, h = get_slide_size(const.get_slide_size_string(slide_size_index)) 

        self.slide = ppt_slide_info(slide_size_index)
        self.textbox = ppt_textbox_info(0,0,w,h,const._default_hymal_font_size)
        
        self.chap = 10
        self.chap_font_size = const._default_hymal_chap_font_size
        self.back_col = color.ppt_color(0,32,96)
