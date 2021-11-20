import re
from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_ANCHOR

_TEXTBOX_SOLID_FILL = 0
_TEXTBOX_GRADIENT_FILL = 1

_slide_size_type = [
    "[ 4:3],[10:7.5  ]",
    "[16:9],[10:5.625]",
    "[16:9],[13.3:7.5]",
    "[Custom]"
]

_custom_slide_size_index = 3

def get_slide_sizeindex_4x3 (): return 0
def get_slide_sizeindex_16x9(): return 1
def get_slide_size_string(index): return _slide_size_type[index]
def get_custom_slide_size_string(): return _slide_size_type[_custom_slide_size_index]

#def get_slide_sizeindex_16x9(): return 2

#_skip_hymal_info = re.compile('[\(\d\)]')

_wide_slide_size_index_1 = 1
_wide_slide_size_index_2 = 2

_default_txt_sx = 0
_default_txt_sy = 4.48
_default_txt_wid = 10.0
_default_txt_hgt = 1.15
_default_font_size = 20.0 # point
_default_font_name = "맑은고딕"
_default_txt_nparagraph = 1
_default_slide_size_index = 1

_default_hymal_font_size = 40.0
_default_hymal_slide_size_index = 0
_default_hymal_chap_font_size = 22.0

_default_outline_weight = 0.25
_default_textframe_vanchor_index = 1 # middle
_default_pragraph_align_index    = 0 # center


_default_textbox_left_margin = 0.01
_default_textbox_top_margin = 0.05
_default_textbox_right_margin = 0.01
_default_textbox_bottom_margin = 0.05

_liveppt_postfix = '-sub'
_message_separator= '-'*15


_ppttab_text     = "PPT"
_slidetab_text   = "Slide"
_hymaltab_text   = "Hymal"
_txtppt_text     = "TxtPPT"
_fxtab_text      = "Fx"
_respread_text   = "RespRd"
_bibppt_text     = "BibPPT"
_messagetab_text = "Message"
_default_pptx_template = "default.pptx"

_worship_type = ["주일예배", "수요예배", "금요성령", "새벽기도", 
                 "부흥회", "부활절", "추수감사", "송구영신", "특별예베", "직접입력"]

_textframe_vanchor = {
    "0": MSO_ANCHOR.TOP, 
    "1": MSO_ANCHOR.MIDDLE,
    "2": MSO_ANCHOR.BOTTOM,
    "3": MSO_ANCHOR.MIXED
    }

_textframe_vanchor_string = [
    "TOP", 
    "MIDDLE",
    "BOTTOM", 
    "MIXED"
    ]
     
def get_all_textframe_vanchortype_string(): return _textframe_vanchor_string
def get_textframe_vanchortype_string(index): return _textframe_vanchor_string[index]
def get_textframe_vanchortype(index): return _textframe_vanchor[str(index)]
                     
_pp_align = {"0": PP_ALIGN.CENTER, 
            "1": PP_ALIGN.DISTRIBUTE,
            "2": PP_ALIGN.JUSTIFY,
            "3": PP_ALIGN.JUSTIFY_LOW,
            "4": PP_ALIGN.LEFT,
            "5": PP_ALIGN.RIGHT,
            "6": PP_ALIGN.THAI_DISTRIBUTE,
            "7": PP_ALIGN.MIXED
            }

_pp_align_string = ["CENTER", "DISTRIBUTE","JUSTIFY", 
                    "JUSTIFY_LOW","LEFT","RIGHT","THAI_DISTRIBUTE","MIXED"]            

def get_all_pragraph_aligntype_string(): return _pp_align_string
def get_pagraph_aligntype_string(index): return _pp_align_string[index]
def get_paragraph_aligntype(index): return _pp_align[str(index)]
            
_responsive_reading_file = 'respread.fmt'

_responsive_reading_format_key = [
'Slide', 'TextBox', 'Text', 'NewLine'
]
_responsive_reading_default_format = [
        'Slide   | 10, 7.5',
        'TextBox | 0.43, 0.92, 6.59, 9.57, 0:32:96',
        'Text    | 맑은고딕, 30, 1, 255:255:255, Left',
        'NewLine | 맑은고딕, 20',
        'Text    | 맑은고딕, 20, 0, 255:255:255, Left',
        'Text    | 맑은고딕, 28, 1, 255:255:255, Left',
        'NewLine | 맑은고딕, 30',
        'Text    | 맑은고딕, 20, 0, 255:255:255, Left',
        'Text    | 맑은고딕, 44, 1, 255:255:0  , Left'
    ]

_responsive_reading_text_align = {
    'Left'  : PP_ALIGN.LEFT,
    'Right' : PP_ALIGN.RIGHT,
    'Center': PP_ALIGN.CENTER,
}

#print(get_aligntype(1))