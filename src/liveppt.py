'''
	Author: Uisang Hwang
	Email : uhwangtx@gmail.com
	
	LivePPT Ver 0.1
	
	08/07/20  Custom Image Resolution @ PPP to Image
	          Custom Shadow effect
	08/08/20  Separation of outline and shadow effect 
	          Each presentation opens a same file.
			  Add Text align
	08/13/20  Merge ppt files
	08/15/20  Add dash, transprancy in Fx outline
	08/21/20  Add Gradient fill in a text frame
		      Add margin (L,T,R,B) of a text frame 
	08/23/20  Rewrite the code
	
	Convert praise ppt to subtitle ppt for live streaming

	Requirements
	----------------------------------------------
	* Read multiple praise ppt
	* Get slide size (4:3 or 16:9)
	* Get text box size: sx, sy, wid, hgt
	* Get text properties: typeface, size(point), color (custom, rgb)
	* Insert blank slide after a praise ppt
	* Can change the order of source ppt files
	* Can delete a source ppt file
	* Remove all ppt files from list box
	
	Dependencies
	----------------------------------------------
	* PyQt4
	* Python pptx
	* win32com
	
	Outline text VBA code
	----------------------------------------------
	Sub outline()
	Dim sld As Slide
	Dim shp As Shape
	
	For Each sld In ActivePresentation.Slides
		sld.Select
		ActiveWindow.ViewType = ppViewSlide
		ActiveWindow.Activate
		
		sld.Shapes(1).Select
		With ActiveWindow.Selection.TextRange2.Font
			.Line.Visible = msoCTrue
			.Line.Pattern = msoPattern10Percent
			.Line.ForeColor.RGB = RGB(255, 0, 0)
			.Line.Weight = 2
			.Line.Style = msoLineSingle
			.Line.DashStyle = msoLineDash
			.Line.Transparency = 1
		End With
	Next sld
	End Sub

	Shadow text VBA code
	----------------------------------------------
	Sub shadow()
	Dim sld As Slide
	Dim shp As Shape
	For Each sld In ActivePresentation.Slides
		For Each shp In sld.Shapes
			With shp.TextFrame2.TextRange.Font
			.shadow.Visible = msoCTrue
			.shadow.Style = msoShadowStyleOuterShadow
			.shadow.OffsetX = 2
			.shadow.OffsetY = 2
			' !!!! Do not use size option !!!!
			'.shadow.Size = 1
			.shadow.Blur = 2
			.shadow.Transparency = 0.7
			End With
		Next shp
	Next sld
	End Sub

'''

import re
import os
import datetime
import pptx
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
import sys
from PyQt4 import QtCore, QtGui
import win32com.client
import msoLine
import msoDash
import msoShadow

import icon_file_add
import icon_folder_open
import icon_arrow_down    
import icon_arrow_up     
#import icon_delete_all    
import icon_delete        
import icon_table_sort_asc  
import icon_table_sort_desc 
import icon_trash
#import icon_play
import icon_convert
import icon_color_picker
import icon_font_picker
import icon_ppt_pptx
import icon_ppt_image
import icon_merge
import icon_shadow
import icon_outline
import icon_liveppt

class ppt_color:
	def __init__(self, r=255,g=255,b=255):
		self.r = r
		self.g = g
		self.b = b
	def __str__(self):
		return "%3d,%3d,%3d"%(self.r,self.g,self.b)

_slide_size_type = [
	"[ 4:3],[10:7.5  ]",
	"[16:9],[10:5.625]",
	"[16:9],[13.3:7.5]"
]

#_skip_hymal_info = re.compile('[\(\d\)]')
_skip_hymal_info = re.compile('찬송가')
_find_lyric_number = re.compile('\d\.')
_find_rgb = re.compile("(\d{1,3}),\s*(\d{1,3}),\s*(\d{1,3})")

_default_txt_sx = 0
_default_txt_sy = 4.48
_default_txt_wid = 10.0
_default_txt_hgt = 1.15
_default_slide_back_col = ppt_color(255,255,255)
_default_font_col  = ppt_color(255,255,255)
_default_font_size = 20.0 # point
_default_font_name = "맑은고딕"
_default_txt_nparagraph = 1
_default_slide_size_index = 1

_default_hymal_font_size        = 40.0
_default_hymal_slide_size_index = 0
_default_hymal_chap_font_size   = 22.0

_default_outline_weight   = 0.25
_default_text_align_index = 4

_default_textbox_left_margin   = 0.7
_default_textbox_top_margin    = 0.05
_default_textbox_right_margin  = 0.1
_default_textbox_bottom_margin = 0.05

_default_solid_fill_color     = ppt_color(100,100,100)
_default_gradient_fill_color1 = ppt_color(0,0,0)
_default_gradient_fill_color2 = ppt_color(155,155,155)

_liveppt_postfix = '-sub'
_message_separator= '-'*15

_color_black = ppt_color(0,0,0)
_color_white = ppt_color(255,255,255)
_color_red   = ppt_color(255,0,0)
_color_green = ppt_color(0,255,0)
_color_blue  = ppt_color(0,0,255)
_color_yellow= ppt_color(255,255,0)

_ppttab_text     = "PPT"
_slidetab_text   = "Slide"
_hymaltab_text   = "Hymal"
_txtppt_text     = "TxtPPT"
_fxtab_text      = "Fx"
_messagetab_text = "Message"

_worship_type = ["주일예배", "수요예배", "새벽기도", 
                 "부흥회"  , "특별예베", "직접입력"]
				 
_pp_align = {"0": PP_ALIGN.CENTER, 
             "1": PP_ALIGN.DISTRIBUTE,
			 "2": PP_ALIGN.JUSTIFY,
			 "3": PP_ALIGN.JUSTIFY_LOW,
			 "4": PP_ALIGN.LEFT,
			 "5": PP_ALIGN.RIGHT,
			 "6": PP_ALIGN.THAI_DISTRIBUTE,
			 "7": PP_ALIGN.MIXED
			 }
			 
def get_textalign(idx):
	return _pp_align[str(idx)]
	
def RGB(red, green, blue):
    assert 0 <= red <=255    
    assert 0 <= green <=255
    assert 0 <= blue <=255
    return red + (green << 8) + (blue << 16)
	
def get_rgb(c):
	c1 = _find_rgb.search(c)
	r = int(c1.group(1))
	g = int(c1.group(2))
	b = int(c1.group(3))
	assert 0 <= r <=255    
	assert 0 <= g <=255
	assert 0 <= b <=255
	return ppt_color(r, g, b)
	
def get_slide_size(t):
	t1 = t.split(',')
	t2 = t1[1][1:-1].split(':')
	return float(t2[0]), float(t2[1])

class QShadowInfo(QtGui.QDialog):
	def __init__(self, shd):
		super(QShadowInfo, self).__init__()
		self.initUI(shd)
		
	def initUI(self, shd):
		layout = QtGui.QFormLayout()
		button_layout = QtGui.QHBoxLayout()
		item_layout = QtGui.QVBoxLayout()
		
		self.style = QtGui.QComboBox(self)
		shadow_style = ["Inner", "Outer", "Mixed"]
		self.style.addItems(shadow_style)
		self.style.setCurrentIndex(shd.Style-1)
		item_layout.addWidget(self.style)
		item_layout.addWidget(QtGui.QLabel("Offset(X)"))
		self.offset_x = QtGui.QLineEdit("%d"%shd.OffsetX)
		item_layout.addWidget(self.offset_x)
		item_layout.addWidget(QtGui.QLabel("Offset(Y)"))
		self.offset_y = QtGui.QLineEdit("%d"%shd.OffsetY)
		item_layout.addWidget(self.offset_y)
		item_layout.addWidget(QtGui.QLabel("Blur"))
		self.blur = QtGui.QLineEdit("%d"%shd.Blur)
		item_layout.addWidget(self.blur)
		item_layout.addWidget(QtGui.QLabel("Transparency"))
		self.trans = QtGui.QLineEdit("%f"%shd.Transparency)
		item_layout.addWidget(self.trans)
		
		self.ok = QtGui.QPushButton('OK')
		self.ok.clicked.connect(self.accept)
		button_layout.addWidget(self.ok)

		self.no = QtGui.QPushButton('Cancel')
		self.no.clicked.connect(self.reject)
		button_layout.addWidget(self.no)
		
		layout.addRow(item_layout)
		layout.addRow(button_layout)
		self.setLayout(layout)
		
	def get_shadow_info(self):
		st = int(self.style.currentIndex())
		ox = int(self.offset_x.text())
		oy = int(self.offset_y.text())
		bl = int(self.blur.text())
		tr = float(self.trans.text())
		return st, ox, oy, bl, tr
		
class QImageResolution(QtGui.QDialog):
	def __init__(self):
		super(QImageResolution, self).__init__()
		self.initUI()
	
	def initUI(self):
		layout = QtGui.QFormLayout()
		button_layout = QtGui.QHBoxLayout()
		item_layout = QtGui.QHBoxLayout()
		
		self.res = QtGui.QComboBox(self)
		self.res_list = ["HD:1280x720", "Full HD:1920x1080", "Quad HD:2560x1440", "Ultra HD:3840x2160"]
		self.res.addItems(self.res_list)
		#for x in encopt._pcm_acodec:
		#	self.pcm.addItem(x)
		item_layout.addWidget(self.res)
		self.res.setCurrentIndex(0)
			
		self.ok = QtGui.QPushButton('OK')
		self.ok.clicked.connect(self.accept)
		button_layout.addWidget(self.ok)

		self.no = QtGui.QPushButton('Cancel')
		self.no.clicked.connect(self.reject)
		button_layout.addWidget(self.no)
		
		layout.addRow(item_layout)
		layout.addRow(button_layout)
		self.setLayout(layout)
		
	def get_resolution(self):
		r = self.res.currentText().split(':')[1].split('x')
		return int(r[0]), int(r[1])
		
class QChooseFxSource(QtGui.QDialog):
	def __init__(self):
		super(QChooseFxSource, self).__init__()
		self.initUI()
		
	def initUI(self):
		layout = QtGui.QFormLayout()
		# Create an array of radio buttons
		moods = [QtGui.QRadioButton("Current"), QtGui.QRadioButton("Fx")]
		
		# Set a radio button to be checked by default
		moods[1].setChecked(True)   
		
		# Radio buttons usually are in a vertical layout   
		source_layout = QtGui.QHBoxLayout()
		
		# Create a button group for radio buttons
		self.mood_button_group = QtGui.QButtonGroup()
		
		for i in range(len(moods)):
			# Add each radio button to the button layout
			source_layout.addWidget(moods[i])
			# Add each radio button to the button group & give it an ID of i
			self.mood_button_group.addButton(moods[i], i)
			# Connect each radio button to a method to run when it's clicked
			#self.connect(moods[i], QtCore.SIGNAL("clicked()"), self.radio_button_clicked)
	
		self.radio_button_clicked()
		button_layout = QtGui.QHBoxLayout()
		self.ok = QtGui.QPushButton('OK')
		self.ok.clicked.connect(self.accept)
		button_layout.addWidget(self.ok)

		self.no = QtGui.QPushButton('Cancel')
		self.no.clicked.connect(self.reject)
		button_layout.addWidget(self.no)
		
		layout.addRow(source_layout)
		layout.addRow(button_layout)
		
		self.setLayout(layout)

	def radio_button_clicked(self):
		return
		
	def get_source(self):
		return self.mood_button_group.checkedId()
		
class ppt_outlinetext_info:
	def __init__(self, col=_color_black,
					   style=msoLine.msoLineSingle,
					   weight=_default_outline_weight):
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
		
_TEXTBOX_SOLID_FILL = 0
_TEXTBOX_GRADIENT_FILL = 1

class ppt_textbox_fill_info:
	def __init__(self):
		self.show = True
		self.fill_type = _TEXTBOX_SOLID_FILL
		self.solid_col = _default_solid_fill_color
		self.gradient_col1 = _default_gradient_fill_color1
		self.gradient_col2 = _default_gradient_fill_color2
		
	def ftype(self): return "Solid" if self.fill_type is _TEXTBOX_SOLID_FILL else "Gradient"
	def __str__(self):
		return 'Fill  : %s\nType  : %s\nS.Col : %s\nG.Col1: %s\nG.Col2: %s'%(
				str(self.show),
				self.ftype(), 
				str(self.solid_col), 
				str(self.gradient_col1),
				str(self.gradient_col2))
		
class ppt_textbox_info:
	def __init__(self, sx=_default_txt_sx,
	                   sy=_default_txt_sy,
					   wid=_default_txt_wid,
					   hgt=_default_txt_hgt,
					   font_size = _default_font_size):
		self.sx            = float(sx)
		self.sy            = float(sy)
		self.wid           = float(wid)
		self.hgt           = float(hgt)
		self.font_name     = _default_font_name
		self.font_col      = ppt_color()
		self.font_bold     = True
		self.font_size     = font_size
		self.word_wrap     = False
		self.left_margin   = _default_textbox_left_margin
		self.top_margin    = _default_textbox_top_margin
		self.right_margin  = _default_textbox_right_margin
		self.bottom_margin = _default_textbox_bottom_margin
		self.fill          = ppt_textbox_fill_info()
		self.align         = _default_text_align_index # center(0), LEFT(4)
		self.fx            = ppt_fx_info()
		
	def __str__(self):
		return 'Sx   : %2.4f\nSy   : %2.4f\nWid  : %2.4f\nHgt  : %2.4f\n\nFont : %s\nColor: %s\nSize : %2.4f\nWrap : %s\nAlign: %s\n\nLeft  : %2.2f\nTop   : %2.2f\nRight : %2.2f\nBottom: %2.2f\n\n%s'%(
			   self.sx, self.sy, self.wid, self.hgt, self.font_name, 
			   str(self.font_col), self.font_size, str(self.word_wrap), 
			   get_textalign(self.align), self.left_margin, self.top_margin, 
			   self.right_margin, self.bottom_margin, str(self.fill))
		
class ppt_slide_info:
	def __init__(self, w=None,h=None):
		if not w and not h:
			w, h = get_slide_size(_slide_size_type[_default_slide_size_index])
		self.back_col = _default_slide_back_col
		self.wid = w
		self.hgt = h
		self.skip_image = True
		self.size_index = _default_slide_size_index
		self.deep_copy = True
	
	def __str__(self):
		return "Wid  : %2.4f\nHgt  : %2.4f\nColor: %s\nDCopy: %s"%(\
		       self.wid,self.hgt,str(self.back_col),str(self.deep_copy))

# use max size: sx=0, sy=0, wid=max wid, hgt=max hgt
class ppt_hymal_info(ppt_slide_info, ppt_textbox_info):
	def __init__(self):
		w, h = get_slide_size(_slide_size_type[_default_hymal_slide_size_index])
		ppt_slide_info.__init__(self,w,h)
		ppt_textbox_info.__init__(self, 0,0,w,h,_default_hymal_font_size)
		self.chap = 10
		self.chap_font_size = _default_hymal_chap_font_size
		self.back_col = ppt_color(0,32,96)
		
class QLivePPT(QtGui.QWidget):
	def __init__(self):
		super(QLivePPT, self).__init__()
		self.initUI()

	def initUI(self):
		self.set_common_var()
		self.form_layout  = QtGui.QFormLayout()
		tab_layout = QtGui.QVBoxLayout()
		self.tabs = QtGui.QTabWidget()
		policy = self.tabs.sizePolicy()
		policy.setVerticalStretch(1)
		self.tabs.setSizePolicy(policy)
		self.tabs.setEnabled(True)
		self.tabs.setTabPosition(QtGui.QTabWidget.West)
		self.tabs.setObjectName('Media List')

		self.ppt_tab = QtGui.QWidget()
		self.slide_tab = QtGui.QWidget()
		self.hymal_tab = QtGui.QWidget()
		self.fx_tab = QtGui.QWidget()
		self.message_tab = QtGui.QWidget()
		self.txtppt_tab = QtGui.QWidget()
		self.tabs.addTab(self.ppt_tab, _ppttab_text)
		self.tabs.addTab(self.slide_tab, _slidetab_text)
		self.tabs.addTab(self.fx_tab, _fxtab_text)
		self.tabs.addTab(self.hymal_tab, _hymaltab_text)
		self.tabs.addTab(self.txtppt_tab, _txtppt_text)
		self.tabs.addTab(self.message_tab, _messagetab_text)
		
		self.message_tab_UI()
		self.ppt_tab_UI()
		self.slide_tab_UI()
		self.hymal_tab_UI()
		self.fx_tab_UI()
		self.txtppt_tab_UI()
		tab_layout.addWidget(self.tabs)
		self.form_layout.addRow(tab_layout)
		self.setLayout(self.form_layout)
		
		self.setWindowTitle("PPT")
		self.setWindowIcon(QtGui.QIcon(QtGui.QPixmap(icon_liveppt.table)))
		self.show()
		
	def fx_tab_UI(self):
		import msoDash
		layout = QtGui.QFormLayout()
		ohlyout = QtGui.QHBoxLayout()
		ohlyout.addWidget(QtGui.QLabel("OUTLINE EFFECT"))
		self.fx_show_outline = QtGui.QCheckBox()
		self.fx_show_outline.setChecked(self.ppt_textbox.fx.show_outline)
		ohlyout.addWidget(self.fx_show_outline)

		self.fx_outline_tbl = QtGui.QTableWidget()
		font = QtGui.QFont("Arial",9,True)
		self.fx_outline_tbl.setFont(font)
		self.fx_outline_tbl.horizontalHeader().hide()
		self.fx_outline_tbl.verticalHeader().hide()
		self.fx_outline_tbl.setColumnCount(3)
		self.fx_outline_tbl.setHorizontalHeaderItem(0, QtGui.QTableWidgetItem("ddd"))
		self.fx_outline_tbl.setHorizontalHeaderItem(1, QtGui.QTableWidgetItem("Data"))
		self.fx_outline_tbl.setHorizontalHeaderItem(2, QtGui.QTableWidgetItem("aaa"))

		header = self.fx_outline_tbl.horizontalHeader()
		header.setResizeMode(0, QtGui.QHeaderView.ResizeToContents)
		header.setResizeMode(1, QtGui.QHeaderView.Stretch)
		header.setResizeMode(2, QtGui.QHeaderView.ResizeToContents)
		
		info = ["Show", "Color", "Line", "Dash", "Transp", "weight"]
		self.fx_outline_tbl.setRowCount(len(info))
		for ii, jj in enumerate(info):
			item = QtGui.QTableWidgetItem(jj)
			item.setFlags(QtCore.Qt.ItemIsEnabled)
			self.fx_outline_tbl.setItem(ii, 0, item)
			self.fx_outline_tbl.setItem(ii, 1, QtGui.QTableWidgetItem(""))

			
		self.fx_outline_show = QtGui.QCheckBox()
		self.fx_outline_show.setChecked(True)
		self.fx_outline_tbl.setCellWidget(0, 1, self.fx_outline_show)
		
		self.fx_outline_col_picker = QtGui.QPushButton('', self)
		self.fx_outline_col_picker.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
		self.fx_outline_col_picker.setIconSize(QtCore.QSize(16,16))
		self.connect(self.fx_outline_col_picker, QtCore.SIGNAL('clicked()'), self.pick_text_outline_color)
		self.fx_outline_tbl.setCellWidget(1,2, self.fx_outline_col_picker)
		self.fx_outline_tbl.item(1,1).setText(str(self.ppt_textbox.fx.outline.col))
		
		style = msoLine.get_linestyle_list()
		self.fx_outline_style = QtGui.QComboBox()
		self.fx_outline_style.addItems(style)
		self.fx_outline_tbl.setCellWidget(2,1, self.fx_outline_style)

		dash = msoDash.get_dashstyle_list()
		dash.insert(0, "N/A")
		self.fx_outline_dash_style = QtGui.QComboBox()
		self.fx_outline_dash_style.addItems(dash)
		self.fx_outline_dash_style.setCurrentIndex(1) # solid 
		self.fx_outline_tbl.setCellWidget(3,1, self.fx_outline_dash_style)
		
		self.fx_outline_tbl.item(4,1).setText("%f"%self.ppt_textbox.fx.outline.transprancy)
		self.fx_outline_tbl.item(5,1).setText("%f"%self.ppt_textbox.fx.outline.weight)
				
		self.fx_outline_tbl.resizeRowsToContents()

		shlyout = QtGui.QHBoxLayout()
		shlyout.addWidget(QtGui.QLabel("SHADOW EFFECT"))
		self.fx_show_shadow = QtGui.QCheckBox()
		self.fx_show_shadow.setChecked(self.ppt_textbox.fx.show_shadow)
		shlyout.addWidget(self.fx_show_shadow)
        
		self.fx_shadow_tbl = QtGui.QTableWidget()
		self.fx_shadow_tbl.resizeRowsToContents()	
		self.fx_shadow_tbl.setFont(font)
		self.fx_shadow_tbl.horizontalHeader().hide()
		self.fx_shadow_tbl.verticalHeader().hide()
		self.fx_shadow_tbl.setColumnCount(3)
		self.fx_shadow_tbl.setHorizontalHeaderItem(0, QtGui.QTableWidgetItem("ddd"))
		self.fx_shadow_tbl.setHorizontalHeaderItem(1, QtGui.QTableWidgetItem("Data"))
		self.fx_shadow_tbl.setHorizontalHeaderItem(2, QtGui.QTableWidgetItem("aaa"))
        
		info = ["Style", "Offset(X)", "Offset(Y)", "Blur", "Transp"]
		self.fx_shadow_tbl.setRowCount(len(info))
		for ii, jj in enumerate(info):
			item = QtGui.QTableWidgetItem(jj)
			item.setFlags(QtCore.Qt.ItemIsEnabled)
			self.fx_shadow_tbl.setItem(ii, 0, item)
			self.fx_shadow_tbl.setItem(ii, 1, QtGui.QTableWidgetItem(""))
        
		self.fx_shadow_style = QtGui.QComboBox(self)
		s_style = ["Inner", "Outer", "Mixed"]
		self.fx_shadow_style.addItems(s_style)
		self.fx_shadow_style.setCurrentIndex(self.ppt_textbox.fx.shadow.Style-1)
		self.fx_shadow_tbl.setCellWidget(0,1,self.fx_shadow_style)
		
		self.fx_shadow_tbl.item(1,1).setText("%d"%self.ppt_textbox.fx.shadow.OffsetX)
		self.fx_shadow_tbl.item(2,1).setText("%d"%self.ppt_textbox.fx.shadow.OffsetY)
		self.fx_shadow_tbl.item(3,1).setText("%d"%self.ppt_textbox.fx.shadow.Blur)
		self.fx_shadow_tbl.item(4,1).setText("%f"%self.ppt_textbox.fx.shadow.Transparency)
		
		self.fx_shadow_tbl.resizeRowsToContents()
		
		layout.addRow(ohlyout)
		layout.addRow(self.fx_outline_tbl)
		layout.addRow(shlyout)
		layout.addRow(self.fx_shadow_tbl)
		self.fx_tab.setLayout(layout)
		self.global_message.appendPlainText("... Fx Tab UI created\n")

	def clear_global_message(self):
		self.global_message.clear()
		
	def message_tab_UI(self):
		layout = QtGui.QVBoxLayout()
		
		clear_btn = QtGui.QPushButton('Clear', self)
		clear_btn.clicked.connect(self.clear_global_message)
		
		self.global_message = QtGui.QPlainTextEdit()
		self.global_message.setFont(QtGui.QFont("Courier",9,True))
		policy = self.sizePolicy()
		policy.setVerticalStretch(1)
		self.global_message.setSizePolicy(policy)
		self.global_message.setFont(QtGui.QFont( "Courier,15,-1,2,50,0,0,0,1,0"))
		layout.addWidget(clear_btn)
		layout.addWidget(self.global_message)
		self.message_tab.setLayout(layout)
		self.global_message.appendPlainText("... Message Tab UI Created")
		
	def clear_txtppt(self):
		self.txtppt_edit.clear()
		
	def change_txtppt_save_path(self):
		startingDir = os.getcwd() 
		path = QtGui.QFileDialog.getExistingDirectory(None, 'Save folder', startingDir, QtGui.QFileDialog.ShowDirsOnly)
		if not path: return
		self.txtppt_save_path.setText(path)
		
	def create_txtppt(self):
		text = self.txtppt_edit.toPlainText().split('\n')
		
		for i, t in enumerate(text):
			text[i] = t.strip()

		size = len(text) 
		if size == 1:
			line_text = text
		else:
			#https://www.geeksforgeeks.org/python-split-list-into-lists-by-particular-value/
			idx_list = [idx + 1 for idx, val in enumerate(text) if val == ''] 
			line_text = [text[i:j] 
			                for i, j in zip([0] + idx_list, 
							                      idx_list + (
												               [size] if idx_list[-1] != size else []
															 )
					                       )
						] 
			
		fn = self.txtppt_fname.text()
		if fn.endswith('.pptx'):
			sfn = os.path.join(self.txtppt_save_path.text(), fn)
		else:
			sfn = os.path.join(self.txtppt_save_path.text(), '%s.pptx'%fn)
		
		if os.path.isfile(sfn):
			ans = QtGui.QMessageBox.question(self, 'Continue?', 
					'%s already exist!'%sfn, QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
			if ans == QtGui.QMessageBox.No: return
			else: os.remove(sfn)
			
		self.global_message.appendPlainText('... Create TxtPPT')
		dest_ppt = pptx.Presentation()
		dest_ppt.slide_width = pptx.util.Inches(self.ppt_hymal.wid)
		dest_ppt.slide_height = pptx.util.Inches(self.ppt_hymal.hgt)
		blank_slide_layout = dest_ppt.slide_layouts[6]

		self.global_message.appendPlainText('Total slide: %d'%len(line_text))
		self.global_message.appendPlainText('%s'%str(self.ppt_hymal))
		
		sx  = float(self.hymal_info_table.item(1,1).text())
		sy  = float(self.hymal_info_table.item(2,1).text())
		wid = float(self.hymal_info_table.item(3,1).text())
		hgt = float(self.hymal_info_table.item(4,1).text())
		font_name = self.hymal_info_table.item(5,1).text()

		back_col = get_rgb(self.hymal_info_table.item(8,1).text())
		font_col  = get_rgb(self.hymal_info_table.item(6,1).text())
		font_size = float(self.hymal_info_table.item(7,1).text())

		for in_text in line_text:
			dest_slide = self.add_empty_slide(dest_ppt, blank_slide_layout, back_col)
			txt_box = self.add_textbox(dest_slide, sx, sy, wid, hgt)
			txt_f = txt_box.text_frame
			self.set_textbox(txt_f, MSO_AUTO_SIZE.NONE, 
			MSO_ANCHOR.MIDDLE, MSO_ANCHOR.MIDDLE, 0, 0, 0, 0)
			
			for l2 in in_text:
				if l2 == '': continue
				p = txt_f.add_paragraph()
				p.text = l2
				self.set_paragraph(p, PP_ALIGN.CENTER, font_name, font_size, font_col, True)
		try:
			dest_ppt.save(sfn)
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', str(e), 
			QtGui.QMessageBox.Yes)
			self.global_message.appendPlainText('... Error: %s'%e)
			dest_ppt = None
			return
			
		self.global_message.appendPlainText('... Create TxtPPT: success\n')
		#QtGui.QMessageBox.question(QtGui.QWidget(), 'Completed!', sfn, QtGui.QMessageBox.Yes)
				
	def txtppt_tab_UI(self):
		layout = QtGui.QFormLayout()
		laybtn = QtGui.QHBoxLayout()
		
		clear_btn = QtGui.QPushButton('Clear', self)
		clear_btn.clicked.connect(self.clear_txtppt)
		
		run_btn = QtGui.QPushButton('Create', self)
		run_btn.clicked.connect(self.create_txtppt)
		laybtn.addWidget(clear_btn)
		laybtn.addWidget(run_btn)
		
		laypub = QtGui.QGridLayout()
		#layfolder = QtGui.QHBoxLayout()
		#layfolder.addWidget(QtGui.QLabel('Dest'))
		laypub.addWidget(QtGui.QLabel('Dest'), 0, 0)
		self.txtppt_save_path  = QtGui.QLineEdit(os.getcwd())
		self.txtppt_save_path_btn = QtGui.QPushButton('', self)
		self.txtppt_save_path_btn.clicked.connect(self.change_txtppt_save_path)
		self.txtppt_save_path_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_folder_open.table)))
		self.txtppt_save_path_btn.setIconSize(QtCore.QSize(16,16))
		#layfolder.addWidget(self.txtppt_save_path)
		#layfolder.addWidget(self.txtppt_save_path_btn)
		laypub.addWidget(self.txtppt_save_path, 0, 1)
		laypub.addWidget(self.txtppt_save_path_btn, 0, 2)
		
		#layfile = QtGui.QHBoxLayout()
		#layfile.addWidget(QtGui.QLabel('File'))
		laypub.addWidget(QtGui.QLabel('File'), 1, 0)
		self.txtppt_fname = QtGui.QLineEdit('txtppt.pptx')
		#layfile.addWidget(self.txtppt_fname)
		laypub.addWidget(self.txtppt_fname, 1, 1)
		
		self.txtppt_edit = QtGui.QPlainTextEdit()
		self.txtppt_edit.setFont(QtGui.QFont("Courier",9,True))
		policy = self.sizePolicy()
		policy.setVerticalStretch(1)
		self.txtppt_edit.setSizePolicy(policy)
		self.txtppt_edit.setFont(QtGui.QFont( "Courier,15,-1,2,50,0,0,0,1,0"))
		
		layout.addRow(laybtn)
		#layout.addRow(layfolder)
		#layout.addRow(layfile)
		layout.addRow(laypub)
		layout.addWidget(self.txtppt_edit)
		self.txtppt_tab.setLayout(layout)
		
	def hymal_tab_UI(self):
		layout = QtGui.QFormLayout()
		self.hymal_info_table = QtGui.QTableWidget()
		font = QtGui.QFont("Arial",9,True)
		self.hymal_info_table.setFont(font)
		self.hymal_info_table.horizontalHeader().hide()
		self.hymal_info_table.verticalHeader().hide()
		self.hymal_info_table.setColumnCount(3)
		self.hymal_info_table.setHorizontalHeaderItem(0, QtGui.QTableWidgetItem("ddd"))
		self.hymal_info_table.setHorizontalHeaderItem(1, QtGui.QTableWidgetItem("Data"))
		self.hymal_info_table.setHorizontalHeaderItem(2, QtGui.QTableWidgetItem("aaa"))
		
		header = self.hymal_info_table.horizontalHeader()
		header.setResizeMode(0, QtGui.QHeaderView.ResizeToContents)
		header.setResizeMode(1, QtGui.QHeaderView.Stretch)
		header.setResizeMode(2, QtGui.QHeaderView.ResizeToContents)
		
		info = ["Chapter", "Textbox(sx)", "Textbox(sy)", "Textbox(wid)",
		        "Textbox(hgt)", "Font", "Font(RGB)", "Size(pt)", "Back(RGB)"]

		self.hymal_info_table.setRowCount(len(info))
				
		for ii, jj in enumerate(info):
			item = QtGui.QTableWidgetItem(jj)
			item.setFlags(QtCore.Qt.ItemIsEnabled)
			self.hymal_info_table.setItem(ii, 0, item)
			self.hymal_info_table.setItem(ii, 1, QtGui.QTableWidgetItem(""))
		
		self.hymal_textbox_font_picker = QtGui.QPushButton('', self)
		self.hymal_textbox_font_picker.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_font_picker.table)))
		self.hymal_textbox_font_picker.setIconSize(QtCore.QSize(16,16))
		self.connect(self.hymal_textbox_font_picker, QtCore.SIGNAL('clicked()'), self.pick_hymal_font)
		self.hymal_info_table.setCellWidget(5,2, self.hymal_textbox_font_picker)
		
		self.hymal_font_col_picker = QtGui.QPushButton('', self)
		self.hymal_font_col_picker.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
		self.hymal_font_col_picker.setIconSize(QtCore.QSize(16,16))
		self.connect(self.hymal_font_col_picker, QtCore.SIGNAL('clicked()'), self.pick_hymal_font_color)
		self.hymal_info_table.setCellWidget(6,2, self.hymal_font_col_picker)

		self.hymal_back_col_picker = QtGui.QPushButton('', self)
		self.hymal_back_col_picker.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
		self.hymal_back_col_picker.setIconSize(QtCore.QSize(16,16))
		self.connect(self.hymal_back_col_picker, QtCore.SIGNAL('clicked()'), self.pick_hymal_bk_color)
		self.hymal_info_table.setCellWidget(8,2, self.hymal_back_col_picker)
		
		self.hymal_info_table.item(0,1).setText("%d"%self.ppt_hymal.chap)
		self.hymal_info_table.item(1,1).setText("%f"%float(self.ppt_hymal.sx ))
		self.hymal_info_table.item(2,1).setText("%f"%float(self.ppt_hymal.sy ))
		self.hymal_info_table.item(3,1).setText("%f"%float(self.ppt_hymal.wid))
		self.hymal_info_table.item(4,1).setText("%f"%float(self.ppt_hymal.hgt))
		self.hymal_info_table.item(5,1).setText(self.ppt_hymal.font_name)
		c = self.ppt_hymal.font_col
		c1="%03d,%03d,%03d"%(c.r,c.g,c.b)
		self.hymal_info_table.item(6,1).setText(c1)
		self.hymal_info_table.item(7,1).setText("%d"%self.ppt_hymal.font_size)
		c = self.ppt_hymal.back_col
		c1="%03d,%03d,%03d"%(c.r,c.g,c.b)
		self.hymal_info_table.item(8,1).setText(c1)
		self.hymal_info_table.resizeRowsToContents()			

		publish_layout = QtGui.QHBoxLayout()
		publish_layout.addWidget(QtGui.QLabel('Dest'))
		self.ppt_hymal.save_path = os.getcwd()
		self.hymal_save_path  = QtGui.QLineEdit(self.ppt_hymal.save_path)
		self.hymal_save_path_btn = QtGui.QPushButton('', self)
		self.hymal_save_path_btn.clicked.connect(self.change_hymal_save_path)
		self.hymal_save_path_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_folder_open.table)))
		self.hymal_save_path_btn.setIconSize(QtCore.QSize(16,16))
		publish_layout.addWidget(self.hymal_save_path)
		publish_layout.addWidget(self.hymal_save_path_btn)
		
		run_layout = QtGui.QHBoxLayout()
		self.hymal_convert_bth = QtGui.QPushButton('', self)
		self.hymal_convert_bth.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_convert.table)))
		self.hymal_convert_bth.setIconSize(QtCore.QSize(24,24))
		self.hymal_convert_bth.clicked.connect(self.create_hymal_ppt)
		run_layout.addWidget(self.hymal_convert_bth)

		size_layout = QtGui.QHBoxLayout()
		size_layout.addWidget(QtGui.QLabel('Slide Size'))
		self.choose_hymal_slide_size = QtGui.QComboBox(self)
		self.choose_hymal_slide_size.addItems(_slide_size_type)
		self.choose_hymal_slide_size.setCurrentIndex(0)
		self.choose_hymal_slide_size.currentIndexChanged.connect(self.set_hymal_slide_size)
		size_layout.addWidget(self.choose_hymal_slide_size)
		
		layout.addRow(self.hymal_info_table)
		layout.addRow(size_layout)
		layout.addRow(publish_layout)
		layout.addRow(run_layout)
		self.hymal_tab.setLayout(layout)
		self.global_message.appendPlainText('... Hymal tab UI created')
		
	def set_hymal_slide_size(self):
		w,h = get_slide_size(self.choose_hymal_slide_size.currentText())
		self.ppt_hymal.wid = w
		self.ppt_hymal.hgt = h
		self.hymal_info_table.item(3,1).setText("%f"%w)
		self.hymal_info_table.item(4,1).setText("%f"%h)
		
	def change_hymal_save_path(self):
		startingDir = os.getcwd() 
		path = QtGui.QFileDialog.getExistingDirectory(None, 'Save folder', startingDir, QtGui.QFileDialog.ShowDirsOnly)
		if not path: return
		self.hymal_save_path.setText(path)
	
	def pick_hymal_bk_color(self):
		col = QtGui.QColorDialog.getColor()
		if col.isValid():
			r,g,b,a = col.getRgb()
			self.ppt_hymal.bk_col = ppt_color(r,g,b)
			self.hymal_info_table.item(8,1).setText("%03d,%03d,%03d"%(r, g, b))
		
	def pick_hymal_font_color(self):
		col = QtGui.QColorDialog.getColor()
		if col.isValid():
			r,g,b,a = col.getRgb()
			self.ppt_hymal.font_col = ppt_color(r,g,b)
			self.hymal_info_table.item(6,1).setText("%03d,%03d,%03d"%(r, g, b))
		
	def pick_hymal_font(self):
		font, valid = QtGui.QFontDialog.getFont()
		if valid: 
			self.ppt_hymal.font_name = font.family()
			self.hymal_info_table.item(5,1).setText(font.family())	
        
	def create_hymal_ppt(self):
		import hymal
		import titnum
		
		self.global_message.appendPlainText("... Create Hymal")
		chap = int(self.hymal_info_table.item(0,1).text())
		if chap < 1 or chap > hymal.max_hymal:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', 
			"Invalid hymal chapter: {}".format(chap), QtGui.QMessageBox.Yes)
			return
			
		col_str = self.hymal_info_table.item(8,1).text()
		bk_col = get_rgb(col_str)
		if not bk_col:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Invalid Back color', 
			"Comma separated {}".format(col_str), QtGui.QMessageBox.Yes)
			return

		col_str = self.hymal_info_table.item(6,1).text()
		ft_col = get_rgb(col_str)
		if not ft_col:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Invalid Font color', 
			"Comma separated {}".format(col_str), QtGui.QMessageBox.Yes)
			return
			
		lyric = hymal.get_hymal_by_chapter(chap)
		for key, value in titnum.title_chap.items():
			if value == chap:
				title = key
		
		fname = "%s-%d.pptx"%(title,chap)
		self.global_message.appendPlainText("Title: %s\nChap: %d"%(title, chap))
		sfn = os.path.join(self.hymal_save_path.text(), fname)
		
		if os.path.isfile(sfn):
			ans = QtGui.QMessageBox.question(self, 'Continue?', 
					'%s already exist!'%sfn, QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
			if ans == QtGui.QMessageBox.No: return
			else: os.remove(sfn)
		
		self.ppt_hymal.sx  = float(self.hymal_info_table.item(1,1).text())
		self.ppt_hymal.sy  = float(self.hymal_info_table.item(2,1).text())
		self.ppt_hymal.wid = float(self.hymal_info_table.item(3,1).text())
		self.ppt_hymal.hgt = float(self.hymal_info_table.item(4,1).text())
		self.ppt_hymal.font_name = self.hymal_info_table.item(5,1).text()
		
		self.ppt_hymal.back_col  = bk_col
		self.ppt_hymal.font_size = float(self.hymal_info_table.item(7,1).text())
		self.ppt_hymal.font_col  = ft_col
		
		dest_ppt = pptx.Presentation()
		dest_ppt.slide_width = pptx.util.Inches(self.ppt_hymal.wid)
		dest_ppt.slide_height = pptx.util.Inches(self.ppt_hymal.hgt)
		blank_slide_layout = dest_ppt.slide_layouts[6]
		
		for l1, l in enumerate(lyric):
			dest_slide = self.add_empty_slide(dest_ppt, blank_slide_layout, 
			self.ppt_hymal.back_col)
			txt_box = self.add_textbox(dest_slide, 
			                           self.ppt_hymal.sx,
									   self.ppt_hymal.sy,
									   self.ppt_hymal.wid,
									   self.ppt_hymal.hgt)
			txt_f = txt_box.text_frame
			self.set_textbox(txt_f, MSO_AUTO_SIZE.NONE, MSO_ANCHOR.MIDDLE, 
			MSO_ANCHOR.MIDDLE, 0, 0, 0, 0)
			
			for l2 in l:
				p = txt_f.add_paragraph()
				p.text = l2
				self.set_paragraph(p, PP_ALIGN.CENTER, 
			                   self.ppt_hymal.font_name, 
							   self.ppt_hymal.font_size,
							   self.ppt_hymal.font_col, 
							   True)
			p = txt_f.add_paragraph()
			p.text ='\n'
			self.set_paragraph(p, PP_ALIGN.CENTER, 
			                   self.ppt_hymal.font_name, 
							   10,
							   self.ppt_hymal.font_col, 
							   True)
							   
			p = txt_f.add_paragraph()
			p.text = '(찬송가 %d장 %d절)'%(chap, l1+1)
			self.set_paragraph(p, PP_ALIGN.CENTER, 
			                   self.ppt_hymal.font_name, 
							   20,
							   self.ppt_hymal.font_col, 
							   True)
				
		try:
			dest_ppt.save(sfn)
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', str(e), 
			QtGui.QMessageBox.Yes)
			self.global_message.appendPlainText('... Error: %s'%str(e))
			dest_ppt = None
			return
			
		self.global_message.appendPlainText('... Create Hymal: success\n')
			
	def slide_tab_UI(self):
		layout = QtGui.QFormLayout()

		background_layout  = QtGui.QHBoxLayout()
		background_layout.addWidget(QtGui.QLabel('Back(RGB)')) 
		self.slide_back_col = QtGui.QLineEdit(str(self.ppt_slide.back_col))
		#self.slide_back_col.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		self.slide_back_col.setFixedWidth(100)

		self.back_col_picker = QtGui.QPushButton('', self)
		self.back_col_picker.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
		self.back_col_picker.setIconSize(QtCore.QSize(16,16))
		#self.back_col_picker.setFixedSize(20,20)
		self.connect(self.back_col_picker, QtCore.SIGNAL('clicked()'), self.pick_slide_back_color)
		background_layout.addWidget(self.slide_back_col)
		background_layout.addWidget(self.back_col_picker)
		
		slide_layout1 = QtGui.QHBoxLayout()
		slide_layout1.addWidget(QtGui.QLabel("Custom Slide Size"))
		self.custom_slide_size = QtGui.QCheckBox()
		self.custom_slide_size.stateChanged.connect(self.custom_slide_state_changed)
		self.custom_slide_size.setChecked(False)
		slide_layout1.addWidget(self.custom_slide_size)
		
		slide_layout2 = QtGui.QHBoxLayout()
		self.choose_slide_size = QtGui.QComboBox(self)
		self.choose_slide_size.addItems(_slide_size_type)
		self.choose_slide_size.setCurrentIndex(self.ppt_slide.size_index)
		self.choose_slide_size.currentIndexChanged.connect(self.set_custom_slide_size)
		#self.choose_slide_size.setFixedWidth(50)
		
		self.custom_slide_wid = QtGui.QLineEdit()
		self.custom_slide_hgt = QtGui.QLineEdit()
		self.custom_slide_wid.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		self.custom_slide_hgt.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		#self.custom_slide_wid.setFixedWidth(50)
		#self.custom_slide_hgt.setFixedWidth(50)
		self.custom_slide_wid.setEnabled(False)
		self.custom_slide_hgt.setEnabled(False)
		self.custom_slide_wid.setText("%f"%self.ppt_slide.wid)
		self.custom_slide_hgt.setText("%f"%self.ppt_slide.hgt)
		
		slide_layout2.addWidget(self.choose_slide_size)
		slide_layout2.addWidget(self.custom_slide_wid)
		slide_layout2.addWidget(self.custom_slide_hgt)
		
		text_layout = QtGui.QGridLayout()
		text_layout.addWidget(QtGui.QLabel("Textbox(x)"), 0,0)
		self.textbox_sx = QtGui.QLineEdit("%f"%self.ppt_textbox.sx)
		self.textbox_sx.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		#self.textbox_sx.setFixedWidth(100)
		text_layout.addWidget(self.textbox_sx, 0, 1)
		text_layout.addWidget(QtGui.QLabel("inch"), 0,2)

		text_layout.addWidget(QtGui.QLabel("Textbox(y)"), 1,0)
		self.textbox_sy = QtGui.QLineEdit("%f"%self.ppt_textbox.sy)
		self.textbox_sy.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		text_layout.addWidget(self.textbox_sy, 1, 1)
		text_layout.addWidget(QtGui.QLabel("inch"), 1,2)
		
		text_layout.addWidget(QtGui.QLabel("Textbox(wid)"), 2,0)
		self.textbox_wid = QtGui.QLineEdit("%f"%self.ppt_textbox.wid)
		self.textbox_wid.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		text_layout.addWidget(self.textbox_wid, 2, 1)
		text_layout.addWidget(QtGui.QLabel("inch"), 2,2)

		text_layout.addWidget(QtGui.QLabel("Textbox(hgt)"), 3,0)
		self.textbox_hgt = QtGui.QLineEdit("%f"%self.ppt_textbox.hgt)
		self.textbox_hgt.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		text_layout.addWidget(self.textbox_hgt, 3, 1)
		text_layout.addWidget(QtGui.QLabel("inch"), 3,2)

		t_align = ["CENTER", "DISTRIBUTE","JUSTIFY",
		           "JUSTIFY_LOW","LEFT","RIGHT","THAI_DISTRIBUTE","MIXED"]
		text_layout.addWidget(QtGui.QLabel("Align"))
		self.choose_text_align = QtGui.QComboBox()
		self.choose_text_align.addItems(t_align)
		self.choose_text_align.setCurrentIndex(_default_text_align_index) # center
		text_layout.addWidget(self.choose_text_align)
		
		font_layout = QtGui.QGridLayout()
		font_layout.addWidget(QtGui.QLabel("Font(RGB)"), 0, 0)
		self.textbox_font_col = QtGui.QLineEdit(str(self.ppt_textbox.font_col))
		#self.textbox_font_col.setFixedWidth(100)
		self.textbox_font_col_picker = QtGui.QPushButton('', self)
		self.textbox_font_col_picker.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
		self.textbox_font_col_picker.setIconSize(QtCore.QSize(16,16))
		self.connect(self.textbox_font_col_picker, QtCore.SIGNAL('clicked()'), self.pick_font_color)
		font_layout.addWidget(self.textbox_font_col, 0, 1)
		font_layout.addWidget(self.textbox_font_col_picker,0,2)

		font_layout.addWidget(QtGui.QLabel("Name"),1,0)
		self.textbox_font_name = QtGui.QLineEdit(self.ppt_textbox.font_name)
		font_layout.addWidget(self.textbox_font_name, 1, 1)
		self.textbox_font_picker = QtGui.QPushButton('', self)
		self.textbox_font_picker.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_font_picker.table)))
		self.textbox_font_picker.setIconSize(QtCore.QSize(16,16))
		self.connect(self.textbox_font_picker, QtCore.SIGNAL('clicked()'), self.pick_font)
		font_layout.addWidget(self.textbox_font_picker, 1, 2)

		font_layout.addWidget(QtGui.QLabel("Size"), 2,0)
		self.textbox_font_size = QtGui.QLineEdit("%f"%self.ppt_textbox.font_size)
		font_layout.addWidget(self.textbox_font_size, 2,1)

		option_layout = QtGui.QHBoxLayout()

		self.textbox_deep_copy = QtGui.QCheckBox("D.Copy")
		self.textbox_deep_copy.setChecked(self.ppt_slide.deep_copy)
		self.textbox_deep_copy.stateChanged.connect(self.textbox_deep_copy_state_changed)
		option_layout.addWidget(self.textbox_deep_copy)
		#font_layout.addWidget(self.textbox_deep_copy,3, 0)
		
		self.textbox_font_bold = QtGui.QCheckBox("Bold")
		self.textbox_font_bold.setChecked(True)
		#font_layout.addWidget(self.textbox_font_bold, 3,1)
		option_layout.addWidget(self.textbox_font_bold)
	
		#font_layout.addWidget(QtGui.QLabel("Word wrap"),4, 0)
		#self.textbox_word_wrap = QtGui.QCheckBox("Word wrap")
		self.textbox_word_wrap = QtGui.QCheckBox("Wrap")
		self.textbox_word_wrap.setChecked(self.ppt_textbox.word_wrap)
		self.textbox_word_wrap.stateChanged.connect(self.textbox_word_wrap_state_changed)
		#font_layout.addWidget(self.textbox_word_wrap,3, 2)
		option_layout.addWidget(self.textbox_word_wrap)
		
		self.textbox_fill = QtGui.QCheckBox("Fill")
		self.textbox_fill.setChecked(True)
		#font_layout.addWidget(self.textbox_fill,4,0)
		option_layout.addWidget(self.textbox_fill)
		
		self.textbox_fill_type = QtGui.QComboBox()
		self.textbox_fill_type.addItems(["Solid", "Grad"])
		self.textbox_fill_type.setFixedWidth(50)
		self.connect(self.textbox_fill_type, QtCore.SIGNAL("currentIndexChanged(int)"), self.textbox_filltype_changed)

		self.textbox_solid_color = QtGui.QLineEdit(str(_default_solid_fill_color))
		self.textbox_solid_color.setFixedWidth(100)
		
		self.textbox_solid_color_btn = QtGui.QPushButton('', self)
		self.textbox_solid_color_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
		self.textbox_solid_color_btn.setIconSize(QtCore.QSize(16,16))
		self.connect(self.textbox_solid_color_btn, QtCore.SIGNAL('clicked()'), self.pick_textbox_fill_solid_color)

		self.textbox_gradient_color1 = QtGui.QLineEdit(str(_default_gradient_fill_color1))
		self.textbox_gradient_color1.setFixedWidth(72)
		
		self.textbox_gradient_color2 = QtGui.QLineEdit(str(_default_gradient_fill_color2))
		self.textbox_gradient_color2.setFixedWidth(72)
		
		self.textbox_gradient_color1_btn = QtGui.QPushButton('', self)
		self.textbox_gradient_color1_btn.setFixedSize(20,20)
		self.textbox_gradient_color1_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
		self.textbox_gradient_color1_btn.setIconSize(QtCore.QSize(16,16))
		self.connect(self.textbox_gradient_color1_btn, QtCore.SIGNAL('clicked()'), self.pick_textbox_fill_gradient_color1)
				
		self.textbox_gradient_color2_btn = QtGui.QPushButton('', self)
		self.textbox_gradient_color2_btn.setFixedSize(20,20)
		self.textbox_gradient_color2_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
		self.textbox_gradient_color2_btn.setIconSize(QtCore.QSize(16,16))
		self.connect(self.textbox_gradient_color2_btn, QtCore.SIGNAL('clicked()'), self.pick_textbox_fill_gradient_color1)

		margin_layout  = QtGui.QVBoxLayout()
		margin_layout0 = QtGui.QHBoxLayout()
		margin_layout1 = QtGui.QHBoxLayout()
		margin_layout2 = QtGui.QHBoxLayout()
		
		margin_layout0.addWidget(self.textbox_fill_type)
		margin_layout0.addWidget(self.textbox_solid_color)
		margin_layout0.addWidget(self.textbox_solid_color_btn)
		margin_layout0.addWidget(self.textbox_gradient_color1)
		margin_layout0.addWidget(self.textbox_gradient_color1_btn)
		margin_layout0.addWidget(self.textbox_gradient_color2)
		margin_layout0.addWidget(self.textbox_gradient_color2_btn)
		
		self.textbox_gradient_color1.hide()
		self.textbox_gradient_color1_btn.hide()
		self.textbox_gradient_color2.hide()
		self.textbox_gradient_color2_btn.hide()
		
		margin_layout1.addWidget(QtGui.QLabel('L/R Margin'))
		self.text_box_lmargin = QtGui.QLineEdit('%f'%self.ppt_textbox.left_margin)
		self.text_box_rmargin = QtGui.QLineEdit('%f'%self.ppt_textbox.right_margin)
		self.text_box_lmargin.setFixedWidth(90)
		self.text_box_rmargin.setFixedWidth(90)
		margin_layout1.addWidget(self.text_box_lmargin)
		margin_layout1.addWidget(self.text_box_rmargin)
		
		margin_layout2.addWidget(QtGui.QLabel('T/B Margin'))
		self.text_box_tmargin = QtGui.QLineEdit('%f'%self.ppt_textbox.top_margin)
		self.text_box_bmargin = QtGui.QLineEdit('%f'%self.ppt_textbox.bottom_margin)
		self.text_box_tmargin.setFixedWidth(90)
		self.text_box_bmargin.setFixedWidth(90)
		margin_layout2.addWidget(self.text_box_tmargin)
		margin_layout2.addWidget(self.text_box_bmargin)
		
		margin_layout.addLayout(margin_layout0)
		margin_layout.addLayout(margin_layout1)
		margin_layout.addLayout(margin_layout2)
	
		menu_layout = QtGui.QHBoxLayout()
		self.slide_menu_init_btn = QtGui.QPushButton('Init')
		self.slide_menu_apply_btn = QtGui.QPushButton('Apply')
		self.slide_menu_save_btn = QtGui.QPushButton('Save')
		self.connect(self.slide_menu_init_btn , QtCore.SIGNAL('clicked()'), self.set_init_slide_info)
		self.connect(self.slide_menu_apply_btn, QtCore.SIGNAL('clicked()'), self.apply_current_slide_info)
		self.connect(self.slide_menu_save_btn , QtCore.SIGNAL('clicked()'), self.save_current_slide_info)
		menu_layout.addWidget(self.slide_menu_init_btn)
		menu_layout.addWidget(self.slide_menu_apply_btn)
		menu_layout.addWidget(self.slide_menu_save_btn)
				
		layout.addRow(background_layout)
		layout.addRow(slide_layout1)
		layout.addRow(slide_layout2)
		layout.addRow(text_layout)
		layout.addRow(font_layout)
		layout.addRow(option_layout)
		layout.addRow(margin_layout)
		layout.addRow(menu_layout)
		self.slide_tab.setLayout(layout)
		self.global_message.appendPlainText('... Slide Tab UI created')
	
	def textbox_filltype_changed(self):
		if self.textbox_fill_type.currentIndex() == 0:
			# solid fill
			self.textbox_gradient_color1.hide()
			self.textbox_gradient_color1_btn.hide()
			self.textbox_gradient_color2.hide()
			self.textbox_gradient_color2_btn.hide()
			
			self.textbox_solid_color.show()
			self.textbox_solid_color_btn.show()
			
		else:
			# gradient fill
			self.textbox_solid_color.hide()
			self.textbox_solid_color_btn.hide()

			self.textbox_gradient_color1.show()
			self.textbox_gradient_color1_btn.show()
			self.textbox_gradient_color2.show()
			self.textbox_gradient_color2_btn.show()
			
	def pick_textbox_fill_solid_color(self):
		col = QtGui.QColorDialog.getColor()
		if col.isValid():
			r,g,b,a = col.getRgb()
			self.textbox_solid_color.setText("%03d,%03d,%03d"%(r, g, b))
			self.ppt_textbox.fill.solid_col = ppt_color(r,g,b)

	def pick_textbox_fill_gradient_color1(self):
		col = QtGui.QColorDialog.getColor()
		if col.isValid():
			r,g,b,a = col.getRgb()
			self.textbox_gradient_color1.setText("%03d,%03d,%03d"%(r, g, b))
			self.ppt_textbox.fill.gradient_col1 = ppt_color(r,g,b)

	def pick_textbox_fill_gradient_color2(self):
		col = QtGui.QColorDialog.getColor()
		if col.isValid():
			r,g,b,a = col.getRgb()
			self.textbox_gradient_color2.setText("%03d,%03d,%03d"%(r, g, b))
			self.ppt_textbox.fill.gradient_col2 = ppt_color(r,g,b)

	def text_outline_state_changed(self):
		if self.text_outline.isChecked():
			self.ppt_textbox.fx.show_outline = True
		else:
			self.ppt_textbox.fx.show_outline = False
			
	def textbox_word_wrap_state_changed(self):
		if self.textbox_deep_copy.isChecked():
			self.ppt_textbox.word_wrap = True
		else:
			self.ppt_textbox.word_wrap = False
	
	def textbox_deep_copy_state_changed(self):
		if self.textbox_deep_copy.isChecked():
			self.ppt_slide.deep_copy = True
		else:
			self.ppt_slide.deep_copy = False
			
	def pick_text_outline_color(self):
		col = QtGui.QColorDialog.getColor()
		if col.isValid():
			r,g,b,a = col.getRgb()
			self.ppt_textbox.fx.outline.col = ppt_color(r,g,b)
			self.fx_outline_tbl.item(1,1).setText("%03d,%03d,%03d"%(r, g, b))
			
	def set_init_slide_info(self):
		self.global_message.appendPlainText('... Init Slide Info\n')

		w, h = get_slide_size(_slide_size_type[_default_slide_size_index])
		self.ppt_slide.back_col    = _default_slide_back_col
		self.ppt_slide.size_index  = _default_slide_size_index
		self.ppt_slide.wid         = w
		self.ppt_slide.hgt         = h
		self.ppt_textbox.sx        = _default_txt_sx
		self.ppt_textbox.sy        = _default_txt_sy
		self.ppt_textbox.wid       = _default_txt_wid
		self.ppt_textbox.hgt       = _default_txt_hgt
		self.ppt_textbox.font_name = _default_font_name
		self.ppt_textbox.font_col  = _default_font_col
		self.ppt_textbox.font_size = _default_font_size
		self.ppt_textbox.font_bold = True
		self.ppt_textbox.align     = _default_text_align_index # center(0), left(4)
		self.ppt_textbox.word_wrap = False
		
		self.ppt_textbox.left_margin  = _default_textbox_left_margin
		self.ppt_textbox.top_margin   = _default_textbox_top_margin
		self.ppt_textbox.right_margin = _default_textbox_right_margin
		self.ppt_textbox.bottom_margin= _default_textbox_bottom_margin
		
		self.ppt_textbox.fill.show = True
		self.ppt_textbox.fill.type = _TEXTBOX_SOLID_FILL
		self.ppt_textbox.fill.solid_col     = _default_solid_fill_color
		self.ppt_textbox.fill.gradient_col1 = _default_gradient_fill_color1
		self.ppt_textbox.fill.gradient_col2 = _default_gradient_fill_color2
		
		self.slide_back_col.setText("%s"%(_default_slide_back_col))
		self.custom_slide_size.setChecked(False)
		self.choose_slide_size.setCurrentIndex(_default_slide_size_index)
		self.custom_slide_wid.setEnabled(False)
		self.custom_slide_hgt.setEnabled(False)
		self.custom_slide_wid.setText("%f"%w)
		self.custom_slide_hgt.setText("%f"%h)
		self.custom_slide_size.setChecked(False)
		self.textbox_sx .setText("%f"%_default_txt_sx)
		self.textbox_sy .setText("%f"%_default_txt_sy)
		self.textbox_wid.setText("%f"%_default_txt_wid)
		self.textbox_hgt.setText("%f"%_default_txt_hgt)
		self.choose_text_align.setCurrentIndex(self.ppt_textbox.align) 
		self.textbox_font_name.setText(_default_font_name)
		self.textbox_font_col.setText("%s"%(str(_default_font_col)))
		self.textbox_font_size.setText("%f"%_default_font_size)
		self.textbox_font_bold.setChecked(True)
		self.textbox_word_wrap.setChecked(False)
		self.textbox_fill.setChecked(True)
		self.textbox_fill_type.setCurrentIndex(0)
		
	def apply_current_slide_info(self):
		self.global_message.appendPlainText("... Apply Slide Info")
		self.ppt_slide.back_col = get_rgb(self.slide_back_col.text())
		
		if self.custom_slide_size.isChecked():
			self.ppt_slide.wid = float(self.custom_slide_wid.text())
			self.ppt_slide.hgt = float(self.custom_slide_hgt.text())
		else:
			w,h = get_slide_size(str(self.choose_slide_size.currentText()))
			self.ppt_slide.wid = w
			self.ppt_slide.hgt = h
			
		self.ppt_textbox.sx  = float(self.textbox_sx .text())
		self.ppt_textbox.sy  = float(self.textbox_sy .text())
		self.ppt_textbox.wid = float(self.textbox_wid.text())
		self.ppt_textbox.hgt = float(self.textbox_hgt.text())

		self.ppt_textbox.font_name = self.textbox_font_name.text()
		self.ppt_textbox.font_col  = get_rgb(self.textbox_font_col.text())
		self.ppt_textbox.font_size = float(self.textbox_font_size.text())
		self.ppt_textbox.font_bold = self.textbox_font_bold.isChecked()
		self.ppt_textbox.fill.show = self.textbox_fill.isChecked()
		self.ppt_textbox.fill.type = self.textbox_fill_type.currentIndex()
		self.ppt_textbox.fill.solid_col     = get_rgb(self.textbox_solid_color.text())
		self.ppt_textbox.fill.gradient_col1 = get_rgb(self.textbox_gradient_color1.text())
		self.ppt_textbox.fill.gradient_col2 = get_rgb(self.textbox_gradient_color2.text())
		self.ppt_textbox.word_wrap = self.textbox_word_wrap.isChecked()
		self.ppt_textbox.align     = self.choose_text_align.currentIndex()
		
		self.ppt_textbox.left_margin  = float(self.text_box_lmargin.text())
		self.ppt_textbox.top_margin   = float(self.text_box_tmargin.text())
		self.ppt_textbox.right_margin = float(self.text_box_rmargin.text())
		self.ppt_textbox.bottom_margin= float(self.text_box_bmargin.text())

		self.global_message.appendPlainText('Slide\n%s\n%s\n%s\nTextbox\n%s\n%s\n%s'%(
		_message_separator,	str(self.ppt_slide)  , _message_separator,
		_message_separator,	str(self.ppt_textbox), _message_separator))
		
	def save_current_slide_info(self):
		return
	
	def pick_slide_back_color(self):
		col = QtGui.QColorDialog.getColor()
		if col.isValid():
			r,g,b,a = col.getRgb()
			self.ppt_slide.slide_back_col = ppt_color(r,g,b)
			self.slide_back_col.setText("%03d,%03d,%03d"%(r, g, b))

	def pick_font_color(self):
		col = QtGui.QColorDialog.getColor()
		if col.isValid():
			r,g,b,a = col.getRgb()
			self.ppt_textbox.font_col = ppt_color(r,g,b)
			self.textbox_font_col.setText("%03d,%03d,%03d"%(r, g, b))
		
	def pick_font(self):
		font, valid = QtGui.QFontDialog.getFont()
		if valid:
			self.ppt_textbox.font_name = font.family()
			self.textbox_font_name.setText(font.family())
		
	def custom_slide_state_changed(self):
		if self.custom_slide_size.isChecked():
			self.choose_slide_size.setEnabled(False)
			self.custom_slide_wid.setEnabled(True)
			self.custom_slide_hgt.setEnabled(True)
		else:
			self.choose_slide_size.setEnabled(True)
			self.custom_slide_wid.setEnabled(False)
			self.custom_slide_hgt.setEnabled(False)
		
	def set_custom_slide_size(self):
		w,h = get_slide_size(str(self.choose_slide_size.currentText()))
		self.custom_slide_wid.setText("%f"%w)
		self.custom_slide_hgt.setText("%f"%h)
		
		# textbox sx : not change
		# textbox hgt: not change
		self.textbox_sy.setText('%f'%(h-float(self.textbox_hgt.text())))
		self.textbox_wid.setText('%f'%w)
		
	def ppt_tab_UI(self):
		layout = QtGui.QFormLayout()
		open_layout  = QtGui.QHBoxLayout()
		self.add_btn = QtGui.QPushButton('Add', self)
		self.add_btn.clicked.connect(self.add_ppt_file)
		self.add_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_file_add.table)))
		self.add_btn.setIconSize(QtCore.QSize(16,16))
		self.open_btn = QtGui.QPushButton('Open', self)
		self.open_btn.clicked.connect(self.open_ppt_file)
		self.open_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_folder_open.table)))
		self.open_btn.setIconSize(QtCore.QSize(16,16))
		open_layout.addWidget(self.open_btn)
		open_layout.addWidget(self.add_btn)

		self.ppt_list_table = QtGui.QTableWidget()
		self.ppt_list_table.setColumnCount(4)
		self.ppt_list_table.setHorizontalHeaderItem(0, QtGui.QTableWidgetItem("Name"))
		self.ppt_list_table.setHorizontalHeaderItem(1, QtGui.QTableWidgetItem("Slide"))
		self.ppt_list_table.setHorizontalHeaderItem(2, QtGui.QTableWidgetItem("Path"))
		self.ppt_list_table.setHorizontalHeaderItem(3, QtGui.QTableWidgetItem("Parg"))
	
		header = self.ppt_list_table.horizontalHeader()
		header.setResizeMode(1, QtGui.QHeaderView.ResizeToContents)
		header.setResizeMode(3, QtGui.QHeaderView.ResizeToContents)
		
		self.move_up_btn = QtGui.QPushButton('', self)
		self.move_up_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_arrow_up.table)))
		self.move_up_btn.setIconSize(QtCore.QSize(16,16))
		self.connect(self.move_up_btn, QtCore.SIGNAL('clicked()'), self.move_item_up)
		 
		self.move_down_btn = QtGui.QPushButton('', self)
		self.move_down_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_arrow_down.table)))
		self.move_down_btn.setIconSize(QtCore.QSize(16,16))
		self.connect(self.move_down_btn, QtCore.SIGNAL('clicked()'), self.move_itme_down)
		
		self.sort_asnd_btn = QtGui.QPushButton('', self)
		self.sort_asnd_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_table_sort_asc.table)))
		self.sort_asnd_btn.setIconSize(QtCore.QSize(16,16))

		self.sort_dsnd_btn = QtGui.QPushButton('', self)
		self.sort_dsnd_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_table_sort_desc.table)))
		self.sort_dsnd_btn.setIconSize(QtCore.QSize(16,16))

		self.delete_btn = QtGui.QPushButton('', self)
		self.delete_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_delete.table)))
		self.delete_btn.setIconSize(QtCore.QSize(16,16))
		self.connect(self.delete_btn, QtCore.SIGNAL('clicked()'), self.delete_item)
		
		self.delete_all_btn = QtGui.QPushButton('', self)
		self.delete_all_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_trash.table)))
		self.delete_all_btn.setIconSize(QtCore.QSize(16,16))
		self.connect(self.delete_all_btn, QtCore.SIGNAL('clicked()'), self.delete_all_item)
		
		#self.play = QtGui.QPushButton('', self)
		#self.play.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_play.table)))
		#self.play.setIconSize(QtCore.QSize(16,16))
		
		btn_layout = QtGui.QHBoxLayout()
		btn_layout.addWidget(self.move_up_btn)
		btn_layout.addWidget(self.move_down_btn)
		btn_layout.addWidget(self.sort_asnd_btn)
		btn_layout.addWidget(self.sort_dsnd_btn)
		btn_layout.addWidget(self.delete_btn)
		btn_layout.addWidget(self.delete_all_btn)
		#btn_layout.addWidget(self.play)
		
		publish_layout = QtGui.QGridLayout()
		publish_layout.addWidget(QtGui.QLabel('Date'), 1, 0) 
		self.publish_date  = QtGui.QLineEdit(self)
		publish_layout.addWidget(self.publish_date, 1, 1)
		
		self.add_publish_date = QtGui.QCheckBox()
		self.add_publish_date.stateChanged.connect(self.add_publish_date_state_changed)
		publish_layout.addWidget(self.add_publish_date, 1, 2)
		self.publish_date.setText(datetime.datetime.now().strftime("%m%d%Y"))
		self.add_publish_date.setChecked(True)
		
		publish_layout.addWidget(QtGui.QLabel('Name'), 2, 0) 
		self.publish_title = QtGui.QComboBox(self)
		self.publish_title.addItems(_worship_type)
		self.publish_title.currentIndexChanged.connect(self.custom_worship_type)
		publish_layout.addWidget(self.publish_title, 2, 1)

		publish_layout.addWidget(QtGui.QLabel('Sorc'), 3, 0)
		self.src_directory_path  = QtGui.QLineEdit()
		publish_layout.addWidget(self.src_directory_path, 3, 1)		

		publish_layout.addWidget(QtGui.QLabel('Dest'), 4, 0)
		self.save_directory_path  = QtGui.QLineEdit(os.getcwd())
		self.save_directory_button = QtGui.QPushButton('', self)
		self.save_directory_button.clicked.connect(self.get_save_directory_path)
		self.save_directory_button.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_folder_open.table)))
		self.save_directory_button.setIconSize(QtCore.QSize(16,16))
		publish_layout.addWidget(self.save_directory_path, 4, 1)
		publish_layout.addWidget(self.save_directory_button, 4, 2)
		
		run_layout = QtGui.QGridLayout()
		isz = 32
		self.run_convert = QtGui.QPushButton('', self)
		self.run_convert.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_convert.table)))
		self.run_convert.setIconSize(QtCore.QSize(isz,isz))
		self.run_convert.clicked.connect(self.create_liveppt)

		self.outline_btn = QtGui.QPushButton('', self)
		self.outline_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_outline.table)))
		self.outline_btn.setIconSize(QtCore.QSize(isz,isz))
		self.outline_btn.clicked.connect(self.create_outline_text)

		self.shadow_btn = QtGui.QPushButton('', self)
		self.shadow_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_shadow.table)))
		self.shadow_btn.setIconSize(QtCore.QSize(isz,isz))
		self.shadow_btn.clicked.connect(self.create_shadow_text)
				
		self.merge_btn = QtGui.QPushButton('', self)
		self.merge_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_merge.table)))
		self.merge_btn.setIconSize(QtCore.QSize(isz,isz))
		self.merge_btn.clicked.connect(self.run_merge_ppt)
		
		self.ppt_pptx_btn = QtGui.QPushButton('', self)
		self.ppt_pptx_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_ppt_pptx.table)))
		self.ppt_pptx_btn.setIconSize(QtCore.QSize(isz,isz))
		self.ppt_pptx_btn.clicked.connect(self.convert_ppt_to_pptx)

		self.ppt_img_btn = QtGui.QPushButton('', self)
		self.ppt_img_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_ppt_image.table)))
		self.ppt_img_btn.setIconSize(QtCore.QSize(isz,isz))
		self.ppt_img_btn.clicked.connect(self.convert_ppt_to_image)
		
		run_layout.addWidget(self.run_convert, 0, 0)
		run_layout.addWidget(self.outline_btn, 0, 1)
		run_layout.addWidget(self.shadow_btn, 0, 2)
		run_layout.addWidget(self.merge_btn, 1, 0)
		run_layout.addWidget(self.ppt_pptx_btn, 1, 1)
		run_layout.addWidget(self.ppt_img_btn, 1, 2)
		
		layout.addRow(open_layout)
		layout.addRow(self.ppt_list_table)
		layout.addRow(btn_layout)
		layout.addRow(publish_layout)
		layout.addRow(run_layout)
		self.ppt_tab.setLayout(layout)
		self.global_message.appendPlainText('... PPT Tab UI created')
				
	def run_merge_ppt(self):
		
		nppt = self.ppt_list_table.rowCount()
		if nppt == 0: return

		try:
			self.global_message.appendPlainText('... Merge PPT: open PowerPoint')
			Application = win32com.client.Dispatch("PowerPoint.Application")
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), 
			QtGui.QMessageBox.Yes)
			self.global_message.appendPlainText('... Fail: %s'%str(e))
			return

		save_file = self.get_save_file(pfix='')
		sfn = os.path.join(self.save_directory_path.text(),save_file)
		
		if os.path.isfile(sfn):
			ans = QtGui.QMessageBox.question(self, 'Continue?', 
                 '%s already exist!'%sfn, QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
			if ans == QtGui.QMessageBox.No: return
			else:
				os.remove(sfn)
				
		try:
			self.global_message.appendPlainText('... Open Presentation')
			# create new presentation object
			# presumes if the file exist, it must be deleted
			dest_ppt = Application.Presentations.Add()
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), 
			QtGui.QMessageBox.Yes)
			self.global_message.appendPlainText('... Fail: %s'%str(e))
			return
			
		dest_ppt.Slides.Add(1, 6)
		index = 1
		self.global_message.appendPlainText('... Start InsertFromFile')
		for npr in range(nppt):
			sp = self.ppt_list_table.item(npr, 2).text()
			sf = self.ppt_list_table.item(npr,0).text()
			src = os.path.join(sp, sf)
			src_ppt = Application.Presentations.Open(src)
			dest_ppt.Slides.InsertFromFile(src, index)
			index += len(src_ppt.Slides)+1
			dest_ppt.Slides.Add(index, 6)
			src_ppt.Close()
		self.global_message.appendPlainText('... End')
		
		self.global_message.appendPlainText('... Start change background color')
		for i, sld in enumerate(dest_ppt.Slides):
			sld.Select()
			# ppViewSlide : 1
			Application.ActiveWindow.ViewType = 1
			Application.ActiveWindow.Activate()
			sr = Application.ActiveWindow.Selection.SlideRange
			sr.FollowMasterBackground = False
			#sr.Background.Fill.Solid = True
			c = self.ppt_slide.back_col
			sr.Background.Fill.ForeColor.RGB = RGB(c.r, c.g, c.b)
		self.global_message.appendPlainText('... End')
		
		dest_ppt.SaveAs(sfn)
		Application.Quit()
		self.global_message.appendPlainText('... Merge PPt: success\n')
		
	def create_shadow_text(self, sorce):
		nppt = self.ppt_list_table.rowCount()
		if nppt == 0: return
		
		res = QChooseFxSource()
		if res.exec_() == 1:
			src = res.get_source()
		else: return
		
		if src == 0: # current file on the table
			sfn = os.path.join(self.ppt_list_table.item(0,2).text(),
                               self.ppt_list_table.item(0,0).text())
		else: # fx source
			save_file = self.get_save_file()
			sfn = os.path.join(self.save_directory_path.text(),save_file)

		if os.path.isfile(sfn):
			ans = QtGui.QMessageBox.question(self, 'Continue?', 
                 '%s already exist!'%sfn, QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
			if ans == QtGui.QMessageBox.No: return
			
		try:
			self.global_message.appendPlainText('... Shadow Text: open PowerPoint')
			Application = win32com.client.Dispatch("PowerPoint.Application")
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), 
			QtGui.QMessageBox.Yes)
			self.global_message.appendPlainText('... Fail: %s'%str(e))
			return

		self.ppt_textbox.fx.shadow.Style   = self.fx_shadow_style.currentIndex()+1
		self.ppt_textbox.fx.shadow.OffsetX = int(self.fx_shadow_tbl.item(1,1).text())
		self.ppt_textbox.fx.shadow.OffsetY = int(self.fx_shadow_tbl.item(2,1).text())
		self.ppt_textbox.fx.shadow.Blur    = int(self.fx_shadow_tbl.item(3,1).text())
		self.ppt_textbox.fx.shadow.Transparency = float(self.fx_shadow_tbl.item(4,1).text())
		
		self.global_message.appendPlainText(str(self.ppt_textbox.fx.shadow))
		
		try:
			Presentation = Application.Presentations.Open(sfn)
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), 
			QtGui.QMessageBox.Yes)
			self.global_message.appendPlainText('... Fail: %s'%str(e))
			return
			
		for i, sld in enumerate(Presentation.Slides):
			for shp in sld.Shapes:
				sdw = shp.TextFrame2.TextRange.Font.Shadow
				sdw.Visible = 1 # msoTrue
				sdw.Style   = self.ppt_textbox.fx.shadow.Style
				sdw.OffsetX = self.ppt_textbox.fx.shadow.OffsetX
				sdw.OffsetY = self.ppt_textbox.fx.shadow.OffsetY
				sdw.Blur    = self.ppt_textbox.fx.shadow.Blur
				sdw.Transparency = self.ppt_textbox.fx.shadow.Transparency
			
		Presentation.Save()
		self.global_message.appendPlainText('... Shadow Text: success\n')
		Application.Quit()
		
	def create_outline_text(self, sorce):

		nppt = self.ppt_list_table.rowCount()
		if nppt == 0: return
		
		res = QChooseFxSource()
		if res.exec_() == 1:
			src = res.get_source()
		else: return
		
		if src == 0: # current file on the table
			sfn = os.path.join(self.ppt_list_table.item(0,2).text(),
			                   self.ppt_list_table.item(0,0).text())
		else: # fx source
			save_file = self.get_save_file()
			sfn = os.path.join(self.save_directory_path.text(),save_file)
			
		if os.path.isfile(sfn):
			ans = QtGui.QMessageBox.question(self, 'Continue?', 
                 '%s already exist!'%sfn, QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
			if ans == QtGui.QMessageBox.No: return
		
		try:
			self.global_message.appendPlainText('... Outline Text: open PowerPoint')
			Application = win32com.client.Dispatch("PowerPoint.Application")
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), 
			QtGui.QMessageBox.Yes)
			self.global_message.appendPlainText('... Fail: %s'%str(e))
			return

		try:
			Presentation = Application.Presentations.Open(sfn)
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), 
			QtGui.QMessageBox.Yes)
			self.global_message.appendPlainText('... Fail: %s'%str(e))
			return
		
		
		h = self.fx_outline_show.isChecked()
		c = get_rgb(self.fx_outline_tbl.item(1,1).text())
		s = msoLine.index_to_style(self.fx_outline_style.currentIndex()) # style
		d = self.fx_outline_dash_style.currentIndex()
		t = float(self.fx_outline_tbl.item(4,1).text()) # transprancy
		w = float(self.fx_outline_tbl.item(5,1).text()) # weight
		
		self.ppt_textbox.fx.outline.show = h
		self.ppt_textbox.fx.outline.col = c
		self.ppt_textbox.fx.outline.style = s
		self.ppt_textbox.fx.outline.dash = d
		self.ppt_textbox.fx.outline.transprancy = t
		self.ppt_textbox.fx.outline.weight = w
			
		self.global_message.appendPlainText(str(self.ppt_textbox.fx.outline))
		
		for i, sld in enumerate(Presentation.Slides):
			sld.Select()
			# ppViewSlide : 1
			Application.ActiveWindow.ViewType = 1
			Application.ActiveWindow.Activate()
        
			try:
				sld.Shapes[0].Select()
			except Exception: # in case of an empty slide
				continue 
				
			fnt = Application.ActiveWindow.Selection.TextRange2.Font
			# msoCTrue: 1, msoTrue: -1, msoFalse: 0
			fnt.Line.Visible = h if h is True else False 
			fnt.Line.ForeColor.RGB = RGB(c.r, c.g, c.b)
			fnt.Line.Style = s
			fnt.Line.DashStyle = d
			fnt.Line.Transparency = t
			fnt.Line.Weight = w
		
		Presentation.Save()
		Application.Quit()
		self.global_message.appendPlainText("... Outline text: success\n")
			
	def convert_ppt_to_image(self):
		nppt = self.ppt_list_table.rowCount()
		if nppt == 0: return

		res = QImageResolution()
		if res.exec_() == 1:
			w, h = res.get_resolution()
		else: return
			
		h = int(float(w * self.ppt_slide.hgt) / self.ppt_slide.wid)
				
		try:
			self.global_message.appendPlainText('... Convert PPT to Img: open PowerPoint')
			Application = win32com.client.Dispatch("PowerPoint.Application")
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), 
			QtGui.QMessageBox.Yes)
			self.global_message.appendPlainText("... Fail: %s"%str(e))
			return

		for npr in range(nppt):
			path = self.ppt_list_table.item(npr, 2).text()
			sorc = os.path.join(path, self.ppt_list_table.item(npr,0).text())
			ff = os.path.splitext(self.ppt_list_table.item(npr,0).text())
			img_folder = os.path.join(path, ff[0])
			try:
				os.mkdir(img_folder)
			except FileExistsError:
				pass

			Presentation = Application.Presentations.Open(sorc)
			for i, sld in enumerate(Presentation.Slides):
				sld.Select()
				sld.Export(os.path.join(img_folder,"%s%03d.jpg"%(ff[0],i)), "JPG", w, h)
		Application.Quit()
		self.global_message.appendPlainText("... Convert PPT to Img: success\n")
		
		QtGui.QMessageBox.question(QtGui.QWidget(), 'completed!', "%s"%img_folder, 
		QtGui.QMessageBox.Yes)
		
	def convert_ppt_to_pptx(self):
		nppt = self.ppt_list_table.rowCount()
		if nppt == 0: return

		try:
			self.global_message.appendPlainText('... PPT to PPTX: open PowerPoint')
			Application = win32com.client.Dispatch("PowerPoint.Application")
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), QtGui.QMessageBox.Yes)
			self.global_message.appendPlainText('... Open Powerpoint Fail: %s'%str(e))
			return
			
		for npr in range(nppt):
			path = self.ppt_list_table.item(npr, 2).text()
			sorc = os.path.join(path, self.ppt_list_table.item(npr,0).text())
			try:
				Presentation = Application.Presentations.Open(sorc)
			except Exception as e:
				res = QtGui.QMessageBox.question(QtGui.QWidget(), '', "Error:%s"%str(e),
				QtGui.QMessageBox.Cancel)
				self.global_message.appendPlainText('... Open Fiel Fail: %s'%str(e))
				Application.Quit()
				return
				
			fname = os.path.splitext(self.ppt_list_table.item(npr,0).text())
			dest = os.path.join(self.save_directory_path.text(), "%s.pptx"%(fname[0]))
			try:
				Presentation.Saveas(dest)
			except Exception as e:
				res = QtGui.QMessageBox.question(QtGui.QWidget(), 'Continue?', "Error:%s"%str(e),\
				QtGui.QMessageBox.Yes|QtGui.QMessageBox.Cancel)
				if res is QtGui.QMessageBox.Cancel:
					Presentation.Close()
					Application.Quit()
					return
			Presentation.Close()
		Application.Quit()
		self.global_message.appendPlainText("PPTX: %s"%dest)
		self.global_message.appendPlainText("... PPT to PPTX: success\n")
		QtGui.QMessageBox.question(QtGui.QWidget(), 'completed!', "%s"%dest, QtGui.QMessageBox.Yes)
		
	def set_common_var(self):
		self.ppt_slide = ppt_slide_info()
		self.ppt_textbox = ppt_textbox_info()
		self.ppt_hymal = ppt_hymal_info()
		self.ppt_shadow = ppt_shadow_info() 
		self.ppt_fx = ppt_fx_info()
		
	def custom_worship_type(self):
		cid = self.publish_title.currentIndex()
		nwt = len(_worship_type)-1
		
		if cid == nwt:
			txt, ok = QtGui.QInputDialog.getText(self, 'Custom Worship Type', "Enter")
			if ok:
				self.publish_title.setEditable(True)
				self.publish_title.setItemText(cid, txt)
				self.publish_title.setEditable(False)
		
	def add_publish_date_state_changed(self):
		if self.add_publish_date.isChecked():
			self.publish_date.setEnabled(True)
		else:
			self.publish_date.setEnabled(False)
		
	def get_save_directory_path(self):
		startingDir = os.getcwd() 
		path = QtGui.QFileDialog.getExistingDirectory(None, 'Save folder', startingDir, 
		QtGui.QFileDialog.ShowDirsOnly)
		if not path: return
		self.save_directory_path.setText(path)

	def open_ppt_file(self):
		title = self.open_btn.text()
		self.ppt_filenames = QtGui.QFileDialog.getOpenFileNames(self, title, 
		directory=self.src_directory_path.text(), 
		filter="PPTX (*.pptx);;PPT (*.ppt);;All files (*.*)")
		nppt = len(self.ppt_filenames)

		if nppt: 
			cur_tab = self.tabs.currentIndex()
			if self.tabs.tabText(cur_tab) == _ppttab_text:
				self.clear_pptlist_table()
				self.ppt_list_table.setRowCount(nppt)
				for k in range(nppt):
					fpath, fname = os.path.split(self.ppt_filenames[k])
					try:
						prs = pptx.Presentation(self.ppt_filenames[k])
						nslide = len(prs.slides)
					except Exception:
						nslide = 0
						pass
							
					self.ppt_list_table.setItem(k, 0, QtGui.QTableWidgetItem(fname))
					self.ppt_list_table.setItem(k, 1, QtGui.QTableWidgetItem("%d"%nslide))
					self.ppt_list_table.setItem(k, 2, QtGui.QTableWidgetItem(fpath))
					self.ppt_list_table.setItem(k, 3, QtGui.QTableWidgetItem("%d"%_default_txt_nparagraph))
				self.src_directory_path.setText(fpath)
		
	def add_ppt_file(self):
		title = self.add_btn.text()
		files = QtGui.QFileDialog.getOpenFileNames(self, title, 
		directory=self.src_directory_path.text(), 
		filter="PPTX (*.pptx);;PPT (*.ppt);;All files (*.*)")
		
		nppt = len(files)
		if nppt:
			cur_row = self.ppt_list_table.rowCount()
			for k in range(nppt):
				j = k + cur_row
				fpath, fname = os.path.split(files[k])
				try:
					prs = pptx.Presentation(self.ppt_filenames[k])
				except Exception as e:
					res = QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', str(e), QtGui.QMessageBox.Yes|QtGui.QMessageBox.Cancel)
					if res is QtGui.QMessageBox.Cancel:
						return
						
				self.ppt_list_table.insertRow(j)
				self.ppt_list_table.setItem(j, 0, QtGui.QTableWidgetItem(fname))
				self.ppt_list_table.setItem(j, 1, QtGui.QTableWidgetItem("%s"%len(prs.slides)))
				self.ppt_list_table.setItem(j, 2, QtGui.QTableWidgetItem(fpath))
				self.ppt_list_table.setItem(j, 3, QtGui.QTableWidgetItem("%d"%_default_txt_nparagraph))
			self.src_directory_path.setText(fpath)

	#http://stackoverflow.com/questions/9166087/move-row-up-and-down-in-pyqt4
	def move_itme_down(self):
		row = self.ppt_list_table.currentRow()
		column = self.ppt_list_table.currentColumn()
		ncolumn = self.ppt_list_table.columnCount()
		new_pos = row+2

		if row < self.ppt_list_table.rowCount()-1:
			self.ppt_list_table.insertRow(new_pos)
			for i in range(ncolumn):
				self.ppt_list_table.setItem(new_pos,i,self.ppt_list_table.takeItem(row,i))
				self.ppt_list_table.setCurrentCell(new_pos,column)
			
			self.ppt_list_table.removeRow(row)        

	def move_item_up(self):    
		row = self.ppt_list_table.currentRow()
		column = self.ppt_list_table.currentColumn();
		ncolumn = self.ppt_list_table.columnCount()
		if row > 0:
			self.ppt_list_table.insertRow(row-1)
			for i in range(ncolumn):
				self.ppt_list_table.setItem(row-1,i,self.ppt_list_table.takeItem(row+1,i))
				self.ppt_list_table.setCurrentCell(row-1,column)
			self.ppt_list_table.removeRow(row+1)        
	
	def clear_pptlist_table(self):
		for i in reversed(range(self.ppt_list_table.rowCount())):
			self.ppt_list_table.removeRow(i)
		self.ppt_list_table.setRowCount(0)

	def delete_all_item(self):
		self.clear_pptlist_table()
		
	def delete_item(self): 
		row_count = self.ppt_list_table.rowCount()
		if row_count == 1: self.delete_all_item()
		if row_count > 1:
			column = self.ppt_list_table.currentColumn();
			row = self.ppt_list_table.currentRow()
			for i in range(self.ppt_list_table.columnCount()):
				self.ppt_list_table.setItem(row,i,self.ppt_list_table.takeItem(row+1,i))
				self.ppt_list_table.setCurrentCell(row,column)
			self.ppt_list_table.removeRow(row+1)
			self.ppt_list_table.setRowCount(row_count-1)
	
	def add_empty_slide(self, dest_ppt, layout, bk_col):
		dest_slide = dest_ppt.slides.add_slide(layout)
		bk = dest_slide.background
		bk.fill.solid()
		bk.fill.fore_color.rgb = pptx.dml.color.RGBColor(bk_col.r, bk_col.g, bk_col.b)
		return dest_slide

	def slide_has_text(self, slide):
		p_list = []
		for shape in slide.shapes:
			if shape.has_text_frame: 
				for paragraph in shape.text_frame.paragraphs:
					run_list = []
					for run in paragraph.runs:
						run_list.append(run.text)
					line_text = ''.join(run_list)
					p_list.append(line_text)
			else: return ''
		txt = ''.join(p_list)
		return txt

	def get_save_file(self, pfix = _liveppt_postfix):
		if self.add_publish_date.isChecked():
			save_file = "%s %s%s.pptx"%(
			self.publish_date.text(), 
			self.publish_title.currentText(),pfix)
		else:
			save_file = "%s%s.pptx"%(
			self.publish_title.currentText(), pfix)
		return save_file	
		
	# dp  : dest_ppt
	# bsl : blank_slide_layout
	# bc  : self.ppt_slide.back_col
	
	def insert_empty_slide(self, dp, bsl, bc):
		pbx = self.ppt_textbox
		dest_slide = self.add_empty_slide(dp, bsl, bc)
		txt_box = self.add_textbox(dest_slide, pbx.sx, pbx.sy, pbx.wid,	pbx.hgt)
		txt_f = txt_box.text_frame
		self.set_textbox(txt_f, MSO_AUTO_SIZE.NONE, MSO_ANCHOR.MIDDLE, MSO_ANCHOR.MIDDLE,
		pbx.left_margin, pbx.top_margin, pbx.right_margin, pbx.bottom_margin)
	
	def create_liveppt(self):
		#import copy
		nppt = self.ppt_list_table.rowCount()
		if nppt == 0: return

		save_file = self.get_save_file()
		sfn = os.path.join(self.save_directory_path.text(),save_file)
		
		if os.path.isfile(sfn):
			ans = QtGui.QMessageBox.question(self, 'Continue?', 
					'%s already exist!'%sfn, QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
			if ans == QtGui.QMessageBox.No: return
			else: os.remove(sfn)
			
		self.global_message.appendPlainText('... Create Subtitle PPT')
		dest_ppt = pptx.Presentation()
		dest_ppt.slide_width = pptx.util.Inches(self.ppt_slide.wid)
		dest_ppt.slide_height = pptx.util.Inches(self.ppt_slide.hgt)
		blank_slide_layout = dest_ppt.slide_layouts[6]

		for npr in range(nppt):
			sp = self.ppt_list_table.item(npr, 2).text()
			sf = self.ppt_list_table.item(npr,0).text()
			npf = os.path.join(sp, sf)
			self.global_message.appendPlainText("%02d: %s"%(npr+1, sf))
			npg = int(self.ppt_list_table.item(npr, 3).text())
			try:
				src = pptx.Presentation(npf)
			except Exception as e:
				QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "{}".format(e), QtGui.QMessageBox.Yes)
				self.global_message.appendPlainText('... Fail: %s\n'%e)
				dest_ppt = None
				return

			pbx = self.ppt_textbox
			
			for slide in src.slides:
				in_text = self.slide_has_text(slide)
				if in_text=='' and self.ppt_slide.deep_copy:
					self.insert_empty_slide(dest_ppt, blank_slide_layout, self.ppt_slide.back_col)
					continue
					
				for shape in slide.shapes:
					if not shape.has_text_frame:
						continue
					p_list = []
					for paragraph in shape.text_frame.paragraphs:
						run_list = []
						for run in paragraph.runs:
							run_list.append(run.text)
						line_text = ''.join(run_list)
						
						# Delete 1. 2. 3. ...
						match = _find_lyric_number.search(line_text)
						if match:
							line_text = _find_lyric_number.sub('', line_text)
						
						# skip hymal chapter info. ex: (찬송기 123장)
						line_text = line_text.strip()
						match = _skip_hymal_info.search(line_text)
						if match or not line_text: 
							continue
						p_list.append(line_text)
					
					np_list = len(p_list)

					for j in range(0, np_list, npg):
						if npg > 1 and (np_list-j) < npg: break
						dest_slide = self.add_empty_slide(dest_ppt, blank_slide_layout, self.ppt_slide.back_col)
						txt_box = self.add_textbox(dest_slide, pbx.sx, pbx.sy, pbx.wid, pbx.hgt)
						txt_f = txt_box.text_frame
						self.set_textbox(txt_f, 
						                 MSO_AUTO_SIZE.NONE, 
										 MSO_ANCHOR.MIDDLE, 
										 MSO_ANCHOR.MIDDLE,
										 pbx.left_margin, pbx.top_margin,
										 pbx.right_margin, pbx.bottom_margin)
		
						k_list = []
						for k in range(npg):
							k_list.append(p_list[j+k])
							
						if pbx.word_wrap:
							self.add_paragraph(txt_f, k_list[0], True)
							for w in k_list[1:]:
								self.add_paragraph(txt_f, w)
						else:
							self.add_paragraph(txt_f, ' '.join(k_list), True)

					if npg == 1: continue
					left_over = np_list%npg
					if left_over:
						dest_slide = self.add_empty_slide(dest_ppt, blank_slide_layout, self.ppt_slide.back_col)
						txt_box = self.add_textbox(dest_slide, pbx.sx, pbx.sy, pbx.wid, pbx.hgt)
						txt_f = txt_box.text_frame
						self.set_textbox(txt_f,MSO_AUTO_SIZE.NONE, 
						                       MSO_ANCHOR.MIDDLE, 
											   MSO_ANCHOR.MIDDLE,
											   pbx.left_margin, pbx.top_margin,
											   pbx.right_margin, pbx.bottom_margi)
						k_list = []
						for k in range(left_over):
							k_list.append(p_list[j+k])
							
						if pbx.word_wrap:
							self.add_paragraph(txt_f, k_list[0], True)
							for w in k_list[1:]:
								self.add_paragraph(txt_f, w)
						else:
							self.add_paragraph(txt_f, ' '.join(k_list), True)
			self.insert_empty_slide(dest_ppt, blank_slide_layout, self.ppt_slide.back_col)
			
		try:
			dest_ppt.save(sfn)
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "{}".format(e), QtGui.QMessageBox.Yes)
			self.global_message.appendPlainText('... Fail: %s'%str(e))
			return
		QtGui.QMessageBox.question(QtGui.QWidget(), 'Completed!', sfn, QtGui.QMessageBox.Yes)
		self.global_message.appendPlainText('Dest: %s\n%s\n... Create Subtitle PPT: success\n'%(
		save_file, str(pbx)))
		
		if self.textbox_fill.isChecked():
			ft = self.textbox_fill_type.currentIndex()
			if ft is _TEXTBOX_GRADIENT_FILL:
				self.fill_textbox(sfn, bool(ft))
			else:
				self.fill_textbox(sfn, bool(ft))

	def add_textbox(self, ds, sx, sy, wid, hgt):
		return ds.shapes.add_textbox(\
			pptx.util.Inches(sx),
			pptx.util.Inches(sy),
			pptx.util.Inches(wid),
			pptx.util.Inches(hgt)
		)
		
	def add_paragraph(self, tf, txt, first=False):
		if first:
			p = tf.paragraphs[0]
		else:
			p = tf.add_paragraph()
			
		p.text = txt
		self.set_paragraph(p, get_textalign(self.ppt_textbox.align),
							self.ppt_textbox.font_name,
							self.ppt_textbox.font_size,
							self.ppt_textbox.font_col,
							self.ppt_textbox.font_bold)
		
	# slide(az,va,ha): MSO_AUTO_SIZE.NONE, MSO_ANCHOR.BOTTOM, MSO_ANCHOR.MIDDLE
	# hymal(az,va,ha): MSO_AUTO_SIZE.NONE, MSO_ANCHOR.BOTTOM, MSO_ANCHOR.LEFT
	def set_textbox(self, tf, az, va, ha, lm, tm, rm, bm):
		tf.auto_size = az
		tf.vertical_anchor = va
		tf.horizontal_anchor= ha
		tf.margin_left   = pptx.util.Inches(lm)
		tf.margin_top    = pptx.util.Inches(tm)
		tf.margin_right  = pptx.util.Inches(rm)
		tf.margin_bottom = pptx.util.Inches(bm)
	
	# slide(al,fn,fz,fc,bl) = PP_ALIGN.CENTER, 
	#                         self.ppt_textbox.font_name,
	#						  self.ppt_textbox.font_size,
	#                         self.ppt_textbox.font_col,
	#                         self.ppt_textbox.font_bold
	# hymal(al,fn,fz,fc,bl) = PP_ALIGN.LEFT, 
	def set_paragraph(self, p, al, fn, fz, fc, bl=True):
		p.alignment = al
		p.font.name = fn
		p.font.size = pptx.util.Pt(fz)
		p.font.color.rgb = pptx.dml.color.RGBColor(fc.r, fc.g, fc.b)
		p.font.bold = bl

	'''
	Public Sub gradient_fill()
	Dim sld As Slide
	Dim shp As Shape
	
	For Each sld In ActivePresentation.Slides
		For Each shp In sld.Shapes
			shp.Fill.Visible = msoCTrue
			shp.Fill.GradientAngle = 0
			shp.Fill.GradientStops.Insert RGB(0, 0, 0), 0, 0
			shp.Fill.GradientStops.Insert RGB(64, 64, 64), 0.46, 0
			shp.Fill.GradientStops.Insert RGB(217, 217, 217), 0.77, 0.7
			shp.Fill.GradientStops.Insert RGB(255, 255, 255), 1, 1
		Next shp
	Next sld
	End Sub
	'''
	# https://docs.microsoft.com/en-us/office/vba/api/powerpoint.fillformat.twocolorgradient
	# https://docs.microsoft.com/en-us/office/vba/api/office.msogradientstyle
	
	def fill_textbox(self, sfn, gradient=True):
	
		try:
			self.global_message.appendPlainText('... Fill textbox: open PowerPoint')
			Application = win32com.client.Dispatch("PowerPoint.Application")
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), 
			QtGui.QMessageBox.Yes)
			self.global_message.appendPlainText('... Fail: %s'%str(e))
			return

		try:
			Presentation = Application.Presentations.Open(sfn)
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), 
			QtGui.QMessageBox.Yes)
			self.global_message.appendPlainText('... Fail: %s'%str(e))
			return
			
		self.global_message.appendPlainText('Type: %s'%('Gradient' if gradient else 'Solid'))
		sc = get_rgb(self.textbox_solid_color.text())
		g1 = get_rgb(self.textbox_gradient_color1.text())
		g2 = get_rgb(self.textbox_gradient_color2.text())
		
		self.global_message.appendPlainText('S.Col : %s\nG.Col1: %s\nG.Col2: %s'%(str(sc), str(g1), str(g2)))
		for i, sld in enumerate(Presentation.Slides):
			for shp in sld.Shapes:
				shp.Fill.Visible = True
				if gradient:
					shp.Fill.TwoColorGradient(2, 1)
					shp.Fill.GradientStops[0].Color.RGB = RGB(g1.r, g1.g ,g1.b )
					shp.Fill.GradientStops[1].Color.RGB = RGB(g2.r, g2.g ,g2.b )
				else:
					shp.Fill.ForeColor.RGB = RGB(sc.r,sc.g,sc.b)
				
				# work on Powerpoint VBS but not win32com
				#shp.Fill.GradientAngle = 0
				#shp.Fill.GradientStops.Insert(RGB(0  , 0  , 0  ), 0   , 0  )
				#shp.Fill.GradientStops.Insert(RGB(64 , 64 , 64 ), 0.46, 0  )
				#shp.Fill.GradientStops.Insert(RGB(217, 217, 217), 0.77, 0.7)
				#shp.Fill.GradientStops.Insert(RGB(255, 255, 255), 1   , 1  )
			
		Presentation.Save()
		Application.Quit()
		self.global_message.appendPlainText('... Gradient Fill: success\n')

def main():
	app = QtGui.QApplication(sys.argv)
	#QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(u'Motif'))
	#QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(u'CDE'))
	QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(u'Plastique'))
	#QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(u'Cleanlooks'))
	lppt = QLivePPT()
	sys.exit(app.exec_())
	
if __name__ == '__main__':
	main()	