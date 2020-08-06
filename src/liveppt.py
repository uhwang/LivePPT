'''
	convert praise ppt to subtitle ppt for live streaming

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
	
'''

import re
import os
import datetime
import pptx
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.dml import MSO_LINE
import sys
from PyQt4 import QtCore, QtGui, Qt
import win32com.client
import msoLine

import icon_file_add
import icon_folder_open
import icon_arrow_down    
import icon_arrow_up     
import icon_delete_all    
import icon_delete        
import icon_table_sort_asc  
import icon_table_sort_desc 
import icon_trash
import icon_play
import icon_convert
import icon_color_picker
import icon_font_picker
import icon_ppt_pptx
import icon_ppt_image
import icon_liveppt

_slide_size_type = [
	"[ 4:3],[10:7.5  ]",
	"[16:9],[10:5.625]",
	"[16:9],[13.3:7.5]"
]

_skip_hymal_info = re.compile('[\(\d\)]')
_find_lyric_number = re.compile('\d\.')
_find_rgb = re.compile("(\d{1,3}),\s*(\d{1,3}),\s*(\d{1,3})")

_default_txt_sx = 2.25
_default_txt_sy = 4.66
_default_txt_wid = 5.51
_default_txt_hgt = 0.4
_default_slide_bk_col = (0,32,96)
_default_font_col = (255,255,255)
_default_font_size = 20.0 # point
_default_font_name = "맑은고딕"
_default_txt_nparagraph = 1
_default_slide_size_index = 1
_default_hymal_slide_size_index = 0
_default_hymal_font_size = 44.0
_default_hymal_chap_font_size = 22.0
_default_outline_weight = 1

_color_black = (0,0,0)
_color_white = (255,255,255)
_color_red   = (255,0,0)
_color_green = (0,255,0)
_color_blue  = (0,0,255)
_color_yellow= (255,255,0)

_ppttab_text   = "PPT"
_slidetab_text = "Slide"
_hymaltab_text = "Hymal"
_txttoppt_text = "TxtPPT"
_messagetab_text = "Message"

_SAVE_FOLDER_SLIDE = 0
_SAVE_FOLDER_HYMAL = 1

_worship_type = ["주일예배", "수요예배", "새벽기도", 
                 "부흥회"  , "특별예베", "직접입력"]
def RGB(red, green, blue):
    assert 0 <= red <=255    
    assert 0 <= green <=255
    assert 0 <= blue <=255
    return red + (green << 8) + (blue << 16)
	
def get_rgb(c, sp=':'):
	c1 = c.split(sp)
	return ppt_color(int(c1[0]),int(c1[1]),int(c1[2]))
	
def get_slide_size(t):
	t1 = t.split(',')
	t2 = t1[1][1:-1].split(':')
	return float(t2[0]), float(t2[1])

class ppt_color:
	def __init__(self, r=255,g=255,b=255):
		self.r = r
		self.g = g
		self.b = b
	def __str__(self):
		return "(%3d,%3d,%3d)"%(self.r,self.g,self.b)


class ppt_outlinetext_info:
	def __init__(self, col=_color_black,
					   style=msoLine.msoLineSingle,
					   weight=_default_outline_weight):
		self.col = ppt_color(col[0], col[1], col[2])
		self.style = style
		self.weight = weight
		
	def __str__(self):
		return "Line color: %s\nStyle: %s\nWeight: %d"%\
		        (str(self.col), 
		        msoLine.get_linestyle_name(self.style),
				self.weight)
		
class ppt_textbox_info:
	def __init__(self, sx=_default_txt_sx,
	                   sy=_default_txt_sy,
					   wid=_default_txt_wid,
					   hgt=_default_txt_hgt,
					   font_size = _default_font_size):
		self.sx  = float(sx)
		self.sy  = float(sy)
		self.wid = float(wid)
		self.hgt = float(hgt)
		self.font_name = _default_font_name
		self.font_col  = ppt_color()
		self.font_bold = True
		self.font_size = font_size
		self.word_wrap = False
		self.text_outline = True
		self.outline = ppt_outlinetext_info()
		#self.nparagraph = _default_txt_nparagraph
		#self.paragraph_wrap = False
	def __str__(self):
		return "Sx: %2.4f\nSy: %2.4f\nWid: %2.4f\n\
		        Hgt: %2.4f\nFn: %s\nFc: %s\nFz: %2.4f"%(\
				self.sx, self.sy, self.wid, self.hgt, 
				self.font_name, str(self.font_col), self.font_size)
		
class ppt_slide_info:
	def __init__(self, w=None,h=None):
		if not w and not h:
			w, h = get_slide_size(_slide_size_type[_default_slide_size_index])
		c = _default_slide_bk_col
		self.back_col = ppt_color(c[0], c[1], c[2])
		self.wid = w
		self.hgt = h
		self.skip_image = True
		self.size_index = _default_slide_size_index
		self.deep_copy = True
	
	def __str__(self):
		return ""

# use max size: sx=0, sy=0, wid=max wid, hgt=max hgt
class ppt_hymal_info(ppt_slide_info, ppt_textbox_info):
	def __init__(self):
		w, h = get_slide_size(_slide_size_type[_default_hymal_slide_size_index])
		ppt_slide_info.__init__(self,w,h)
		ppt_textbox_info.__init__(self, 0,0,w,h,_default_hymal_font_size)
		self.chap = 10
		self.chap_font_size = _default_hymal_chap_font_size
		self.save_path = ""
		
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
		self.message_tab = QtGui.QWidget()
		self.txttoppt_tab = QtGui.QWidget()
		self.tabs.addTab(self.ppt_tab, _ppttab_text)
		self.tabs.addTab(self.slide_tab, _slidetab_text)
		self.tabs.addTab(self.hymal_tab, _hymaltab_text)
		self.tabs.addTab(self.txttoppt_tab, _txttoppt_text)
		self.tabs.addTab(self.message_tab, _messagetab_text)
		
		self.ppt_tab_UI()
		self.slide_tab_UI()
		self.hymal_tab_UI()
		self.message_tab_UI()
		tab_layout.addWidget(self.tabs)
		self.form_layout.addRow(tab_layout)
		self.setLayout(self.form_layout)
		
		self.setWindowTitle("PPT")
		self.setWindowIcon(QtGui.QIcon(QtGui.QPixmap(icon_liveppt.table)))
		self.show()

		
	def txttoppt_tab_UI(self):
		return
		
	def clear_global_message(self):
		self.global_message.clear()
		
	def message_tab_UI(self):
		layout = QtGui.QVBoxLayout()
		
		clear_btn = QtGui.QPushButton('Clear', self)
		clear_btn.clicked.connect(self.clear_global_message)
		
		self.global_message = QtGui.QPlainTextEdit()
		policy = self.sizePolicy()
		policy.setVerticalStretch(1)
		self.global_message.setSizePolicy(policy)
		layout.addWidget(clear_btn)
		layout.addWidget(self.global_message)
		self.message_tab.setLayout(layout)
		
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
		self.hymal_save_path_btn.clicked.connect(self.get_save_directory_path)
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

		layout.addRow(self.hymal_info_table)
		layout.addRow(publish_layout)
		layout.addRow(run_layout)
		self.hymal_tab.setLayout(layout)
		
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
		
		chap = int(self.hymal_info_table.item(0,1).text())
		if chap < 1 or chap > hymal.max_hymal:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', 
			"Invalid hymal chapter: {}".format(chap), QtGui.QMessageBox.Yes)
			return
			
		col_str = self.hymal_info_table.item(8,1).text()
		bk_col = _find_rgb.search(col_str)
		if not bk_col:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Invalid Back color', 
			"Comma separated {}".format(col_str), QtGui.QMessageBox.Yes)
			return

		col_str = self.hymal_info_table.item(6,1).text()
		ft_col = _find_rgb.search(col_str)
		if not ft_col:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Invalid Font color', 
			"Comma separated {}".format(col_str), QtGui.QMessageBox.Yes)
			return
			
		lyric = hymal.get_hymal_by_chapter(chap)
		for key, value in titnum.title_chap.items():
			if value == chap:
				title = key
		
		fname = "%s-%d.pptx"%(title,chap)
		save_file = os.path.join(self.ppt_hymal.save_path, fname)
		
		self.ppt_hymal.sx  = float(self.hymal_info_table.item(1,1).text())
		self.ppt_hymal.sy  = float(self.hymal_info_table.item(2,1).text())
		self.ppt_hymal.wid = float(self.hymal_info_table.item(3,1).text())
		self.ppt_hymal.hgt = float(self.hymal_info_table.item(4,1).text())
		self.ppt_hymal.font_name = self.hymal_info_table.item(5,1).text()
		
		self.ppt_hymal.back_col = ppt_color(int(bk_col.group(1)), 
		                                        int(bk_col.group(2)), 
												int(bk_col.group(3)))
		self.ppt_hymal.font_size = float(self.hymal_info_table.item(7,1).text())
		self.ppt_hymal.font_col = ppt_color(int(ft_col.group(1)), 
		                                    int(ft_col.group(2)), 
											int(ft_col.group(3)))
		
		dest_ppt = pptx.Presentation()
		dest_ppt.slide_width = pptx.util.Inches(self.ppt_hymal.wid)
		dest_ppt.slide_height = pptx.util.Inches(self.ppt_hymal.hgt)
		blank_slide_layout = dest_ppt.slide_layouts[6]
		
		for l in lyric:
			dest_slide = self.add_empty_slide(dest_ppt, blank_slide_layout, 
			self.ppt_hymal.back_col)
			txt_box = self.add_textbox(dest_slide, 
			                           self.ppt_hymal.sx,
									   self.ppt_hymal.sy,
									   self.ppt_hymal.wid,
									   self.ppt_hymal.hgt)
			txt_f = txt_box.text_frame
			self.set_textbox(txt_f, MSO_AUTO_SIZE.NONE, MSO_ANCHOR.MIDDLE, MSO_ANCHOR.MIDDLE)
			
			for l1 in l:
				p = txt_f.add_paragraph()
				p.text = l1
				self.set_paragraph(p, PP_ALIGN.CENTER, 
			                   self.ppt_hymal.font_name, 
							   self.ppt_hymal.font_size,
							   self.ppt_hymal.font_col, 
							   True)
		try:
			dest_ppt.save(save_file)
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "{}".format(str(e)), 
			QtGui.QMessageBox.Yes)
			del dest_ppt
			return
			
	def slide_tab_UI(self):
		layout = QtGui.QFormLayout()

		background_layout  = QtGui.QHBoxLayout()
		background_layout.addWidget(QtGui.QLabel('Background(RGB)')) 
		self.slide_back_col = QtGui.QLineEdit()
		self.slide_back_col.setInputMask('999|999|999;-')
		font = QtGui.QFont("Courier",11,True)
		fm = QtGui.QFontMetrics(font)
		self.slide_back_col.setFixedSize(fm.width("888888888888"), fm.height())
		self.slide_back_col.setFont(font)
		col = self.ppt_slide.back_col
		self.slide_back_col.setText("%03d|%03d|%03d"%(col.r, col.g, col.b))
		
		self.back_col_picker = QtGui.QPushButton('', self)
		self.back_col_picker.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
		self.back_col_picker.setIconSize(QtCore.QSize(16,16))
		self.connect(self.back_col_picker, QtCore.SIGNAL('clicked()'), self.pick_bk_color)
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

		self.custom_slide_wid = QtGui.QLineEdit()
		self.custom_slide_hgt = QtGui.QLineEdit()
		self.custom_slide_wid.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		self.custom_slide_hgt.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		self.custom_slide_wid.setEnabled(False)
		self.custom_slide_hgt.setEnabled(False)
		self.custom_slide_wid.setText("%f"%self.ppt_slide.wid)
		self.custom_slide_hgt.setText("%f"%self.ppt_slide.hgt)
		
		slide_layout2.addWidget(self.choose_slide_size)
		slide_layout2.addWidget(self.custom_slide_wid)
		slide_layout2.addWidget(self.custom_slide_hgt)
		
		text_layout = QtGui.QGridLayout()
		text_layout.addWidget(QtGui.QLabel("Textbox(x)"), 0,0)
		self.text_sx = QtGui.QLineEdit("%f"%self.ppt_textbox.sx)
		self.text_sx.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		text_layout.addWidget(self.text_sx, 0, 1)
		text_layout.addWidget(QtGui.QLabel("inch"), 0,2)

		text_layout.addWidget(QtGui.QLabel("Textbox(y)"), 1,0)
		self.text_sy = QtGui.QLineEdit("%f"%self.ppt_textbox.sy)
		self.text_sy.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		text_layout.addWidget(self.text_sy, 1, 1)
		text_layout.addWidget(QtGui.QLabel("inch"), 1,2)
		
		text_layout.addWidget(QtGui.QLabel("Textbox(wid)"), 2,0)
		self.text_wid = QtGui.QLineEdit("%f"%self.ppt_textbox.wid)
		self.text_wid.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		text_layout.addWidget(self.text_wid, 2, 1)
		text_layout.addWidget(QtGui.QLabel("inch"), 2,2)

		text_layout.addWidget(QtGui.QLabel("Textbox(hgt)"), 3,0)
		self.text_hgt = QtGui.QLineEdit("%f"%self.ppt_textbox.hgt)
		self.text_hgt.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		text_layout.addWidget(self.text_hgt, 3, 1)
		text_layout.addWidget(QtGui.QLabel("inch"), 3,2)

		font_layout = QtGui.QGridLayout()
		font_layout.addWidget(QtGui.QLabel("Font(RGB)"), 0, 0)
		self.textbox_font_col = QtGui.QLineEdit()
		self.textbox_font_col.setInputMask('999|999|999;-')
		font = QtGui.QFont("Courier",11,True)
		fm = QtGui.QFontMetrics(font)
		self.textbox_font_col.setFixedSize(fm.width("888888888888"), fm.height())
		self.textbox_font_col.setFont(font)
		col = self.ppt_textbox.font_col
		self.textbox_font_col.setText("%03d|%03d|%03d"%(col.r, col.g, col.b))
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
		font_layout.addWidget(QtGui.QLabel("Bold"),2,2)
		self.textbox_font_bold = QtGui.QCheckBox()
		self.textbox_font_bold.setChecked(True)
		font_layout.addWidget(self.textbox_font_bold, 2,3)
		
		#effect_layout = QtGui.QGridLayout()
		#effect_layout = QtGui.QHBoxLayout()
		font_layout.addWidget(QtGui.QLabel("Deep copy"),3, 0)
		#effect1_layout.addWidget(QtGui.QLabel("Deep copy"))
		self.deep_copy = QtGui.QCheckBox()
		self.deep_copy.setChecked(self.ppt_slide.deep_copy)
		self.deep_copy.stateChanged.connect(self.deep_copy_state_changed)
		font_layout.addWidget(self.deep_copy,3, 1)
		#effect1_layout.addWidget(self.deep_copy)
		
		font_layout.addWidget(QtGui.QLabel("Word wrap"),4, 0)
		#effect1_layout.addWidget(QtGui.QLabel("Word wrap"))
		self.word_wrap = QtGui.QCheckBox()
		self.word_wrap.setChecked(self.ppt_textbox.word_wrap)
		self.word_wrap.stateChanged.connect(self.word_wrap_state_changed)
		font_layout.addWidget(self.word_wrap,4, 1)
		#effect1_layout.addWidget(self.word_wrap)
		
		#effect_layout.addWidget(QtGui.QLabel("Outline"),0, 2)
		#effect2_layout = QtGui.QGridLayout()
		font_layout.addWidget(QtGui.QLabel("Outline"),5,0)
		self.text_outline = QtGui.QCheckBox()
		self.text_outline.setChecked(self.ppt_textbox.text_outline)
		self.text_outline.stateChanged.connect(self.text_outline_state_changed)
		font_layout.addWidget(self.text_outline, 5, 1)
		#effect2_layout.addWidget(self.text_outline, 0,1)

		#font = QtGui.QFont("Courier",11,True)
		#fm = QtGui.QFontMetrics(font)
		font_layout.addWidget(QtGui.QLabel("Color(RGB)"),6, 0)
		self.text_outline_col =  QtGui.QLineEdit()
		self.text_outline_col.setInputMask('999|999|999;-')
		self.text_outline_col.setFixedSize(fm.width("888888888888"), fm.height())
		self.text_outline_col.setFont(font)
		col = self.ppt_textbox.outline.col
		self.text_outline_col.setText("%03d|%03d|%03d"%(col.r, col.g, col.b))
		self.text_outline_col_picker = QtGui.QPushButton('', self)
		self.text_outline_col_picker.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
		self.text_outline_col_picker.setIconSize(QtCore.QSize(16,16))
		self.connect(self.text_outline_col_picker, QtCore.SIGNAL('clicked()'), self.pick_text_outline_color)
		font_layout.addWidget(self.text_outline_col, 6, 1)
		font_layout.addWidget(self.text_outline_col_picker,6,2)
		
		font_layout.addWidget(QtGui.QLabel("Weight"),7, 0)
		self.text_outline_weight =  QtGui.QLineEdit("%d"%self.ppt_textbox.outline.weight)
		self.text_outline_weight.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		font_layout.addWidget(self.text_outline_weight,7,1)
		
		menu_layout = QtGui.QHBoxLayout()
		self.slide_menu_init_btn = QtGui.QPushButton('Init')
		self.slide_menu_apply_btn = QtGui.QPushButton('Apply')
		self.slide_menu_save_btn = QtGui.QPushButton('Save')
		self.connect(self.slide_menu_init_btn , QtCore.SIGNAL('clicked()'), self.set_init_slide_value)
		self.connect(self.slide_menu_apply_btn, QtCore.SIGNAL('clicked()'), self.apply_current_slide_value)
		self.connect(self.slide_menu_save_btn , QtCore.SIGNAL('clicked()'), self.save_current_slide_value)
		menu_layout.addWidget(self.slide_menu_init_btn)
		menu_layout.addWidget(self.slide_menu_apply_btn)
		menu_layout.addWidget(self.slide_menu_save_btn)
				
		layout.addRow(background_layout)
		layout.addRow(slide_layout1)
		layout.addRow(slide_layout2)
		layout.addRow(text_layout)
		layout.addRow(font_layout)
		#layout.addRow(effect_layout)
		layout.addRow(menu_layout)
		self.slide_tab.setLayout(layout)
	
	def text_outline_state_changed(self):
		if self.text_outline.isChecked():
			self.ppt_textbox.text_outline = True
		else:
			self.ppt_textbox.text_outline = False
			
	def word_wrap_state_changed(self):
		if self.deep_copy.isChecked():
			self.ppt_textbox.word_wrap = True
		else:
			self.ppt_textbox.word_wrap = False
	
	def deep_copy_state_changed(self):
		if self.deep_copy.isChecked():
			self.ppt_slide.deep_copy = True
		else:
			self.ppt_slide.deep_copy = False
			
	def pick_text_outline_color(self):
		col = QtGui.QColorDialog.getColor()
		if col.isValid():
			r,g,b,a = col.getRgb()
			self.ppt_textbox.outline.col = ppt_color(r,g,b)
			self.text_outline_col.setText("%03d,%03d,%03d"%(r, g, b))
			
	def set_init_slide_value(self):
		c1 = _default_slide_bk_col
		c2 = _default_font_col
		w, h = get_slide_size(_slide_size_type[_default_slide_size_index])
		self.ppt_slide.back_col = ppt_color(c1[0],c1[1],c1[2])
		self.ppt_slide.size_index = _default_slide_size_index
		self.ppt_slide.wid = w
		self.ppt_slide.hgt = h
		self.ppt_textbox.sx  = _default_txt_sx
		self.ppt_textbox.sy  = _default_txt_sy
		self.ppt_textbox.wid = _default_txt_wid
		self.ppt_textbox.hgt = _default_txt_hgt
		self.ppt_textbox.font_name = _default_font_name
		self.ppt_textbox.font_col = ppt_color(c2[0], c2[1], c2[2])
		self.ppt_textbox.font_size = _default_font_size
		self.ppt_textbox.font_bold = True
		self.ppt_textbox.text_outline = False
		c3 = _color_black
		self.ppt_textbox.outline.col = ppt_color(c3[0], c3[1], c3[2])
		self.ppt_textbox.outline.weight = 1
		
		self.slide_back_col.setText("%03d|%03d|%03d"%(c1[0], c1[1], c1[2]))
		self.custom_slide_size.setChecked(False)
		self.choose_slide_size.setCurrentIndex(_default_slide_size_index)
		self.custom_slide_wid.setEnabled(False)
		self.custom_slide_hgt.setEnabled(False)
		self.custom_slide_wid.setText("%f"%w)
		self.custom_slide_hgt.setText("%f"%h)
		self.custom_slide_size.setChecked(False)
		self.text_sx .setText("%f"%_default_txt_sx)
		self.text_sy .setText("%f"%_default_txt_sy)
		self.text_wid.setText("%f"%_default_txt_wid)
		self.text_wid.setText("%f"%_default_txt_hgt)
		self.textbox_font_name.setText(_default_font_name)
		self.textbox_font_col.setText("%03d|%03d|%03d"%(c2[0], c2[1], c2[2]))
		self.textbox_font_size.setText("%f"%_default_font_size)
		self.textbox_font_bold.setChecked(True)
		self.text_outline.setChecked(False)
		self.text_outline_col.setText("%03d|%03d|%03d"%(c3[0], c3[1], c3[2]))
		self.text_outline_weight.setText("%d"%self.ppt_textbox.outline.weight)
		
	def apply_current_slide_value(self):
		self.ppt_slide.back_col = get_rgb(self.slide_back_col.text(), '|')
		
		if self.custom_slide_size.isChecked():
			self.ppt_slide.wid = float(self.custom_slide_wid.text())
			self.ppt_slide.hgt = float(self.custom_slide_hgt.text())
		else:
			w,h = get_slide_size(str(self.choose_slide_size.currentText()))
			self.ppt_slide.wid = w
			self.ppt_slide.hgt = h
			
		self.ppt_textbox.sx  = float(self.text_sx .text())
		self.ppt_textbox.sy  = float(self.text_sy .text())
		self.ppt_textbox.wid = float(self.text_wid.text())
		self.ppt_textbox.hgt = float(self.text_hgt.text())

		self.ppt_textbox.font_name = self.textbox_font_name.text()
		self.ppt_textbox.font_col = get_rgb(self.textbox_font_col.text(), '|')
		self.ppt_textbox.font_size = float(self.textbox_font_size.text())
		self.ppt_textbox.font_bold = self.textbox_font_bold.isChecked()
		self.ppt_textbox.word_wrap = self.word_wrap.isChecked()
		
		self.ppt_textbox.text_outline = self.text_outline.isChecked()
		if self.ppt_textbox.text_outline:
			self.ppt_textbox.outline.col = get_rgb(self.text_outline_col.text(), '|')
			self.ppt_textbox.outline.weight = int(self.text_outline_weight.text())

		#self.global_message.appendPlainText(str(self.ppt_slide))
		#self.global_message.appendPlainText(str(self.ppt_textbox))
	def save_current_slide_value(self):
		return
	
	def pick_bk_color(self):
		col = QtGui.QColorDialog.getColor()
		if col.isValid():
			r,g,b,a = col.getRgb()
			self.ppt_slide.slide_back_col = ppt_color(r,g,b)
			self.slide_back_col.setText("%03d|%03d|%03d"%(r, g, b))

	def pick_font_color(self):
		col = QtGui.QColorDialog.getColor()
		if col.isValid():
			r,g,b,a = col.getRgb()
			self.ppt_textbox.font_col = ppt_color(r,g,b)
			self.textbox_font_col.setText("%03d|%03d|%03d"%(r, g, b))
		
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
		
		self.play = QtGui.QPushButton('', self)
		self.play.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_play.table)))
		self.play.setIconSize(QtCore.QSize(16,16))
		
		btn_layout = QtGui.QHBoxLayout()
		btn_layout.addWidget(self.move_up_btn)
		btn_layout.addWidget(self.move_down_btn)
		btn_layout.addWidget(self.sort_asnd_btn)
		btn_layout.addWidget(self.sort_dsnd_btn)
		btn_layout.addWidget(self.delete_btn)
		btn_layout.addWidget(self.delete_all_btn)
		btn_layout.addWidget(self.play)
		
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
		
		run_layout = QtGui.QHBoxLayout()
		self.run_convert = QtGui.QPushButton('', self)
		self.run_convert.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_convert.table)))
		self.run_convert.setIconSize(QtCore.QSize(48,48))
		self.run_convert.setStyleSheet("QPushButton { text-align: bottom; }")
		self.run_convert.clicked.connect(self.run_subtitle)

		self.ppt_to_pptx = QtGui.QPushButton('', self)
		self.ppt_to_pptx.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_ppt_pptx.table)))
		self.ppt_to_pptx.setIconSize(QtCore.QSize(48,48))
		self.ppt_to_pptx.clicked.connect(self.convert_ppt_to_pptx)

		self.ppt_to_img = QtGui.QPushButton('', self)
		self.ppt_to_img.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_ppt_image.table)))
		self.ppt_to_img.setIconSize(QtCore.QSize(48,48))
		self.ppt_to_img.clicked.connect(self.convert_ppt_to_image)
		run_layout.addWidget(self.run_convert)
		run_layout.addWidget(self.ppt_to_pptx)
		run_layout.addWidget(self.ppt_to_img)
		
		layout.addRow(open_layout)
		layout.addRow(self.ppt_list_table)
		layout.addRow(btn_layout)
		layout.addRow(publish_layout)
		layout.addRow(run_layout)
		self.ppt_tab.setLayout(layout)
				
	def create_outline_text(self, sorce):
		try:
			Application = win32com.client.Dispatch("PowerPoint.Application")
			#Application.Visible = True
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), 
			QtGui.QMessageBox.Yes)
			return

		Presentation = Application.Presentations.Open(sorce)
		for i, sld in enumerate(Presentation.Slides):
			sld.Select()
			# ppViewSlide : 1
			Application.ActiveWindow.ViewType = 1
			Application.ActiveWindow.Activate()
        
			try:
				sld.Shapes[0].Select()
			except Exception as e:
				continue
				
			fnt = Application.ActiveWindow.Selection.TextRange2.Font
			fnt.Line.Visible = True #msoCTrue
			c = self.ppt_textbox.outline.col
			fnt.Line.ForeColor.RGB = RGB(c.r, c.g, c.b)
			fnt.Line.Weight = self.ppt_textbox.outline.weight
		
		#try:	
		#Application.Save(sorce)
		#except Exception as e:
		#	QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), 
		#	QtGui.QMessageBox.Yes)
		#Application.Quit()
			
	def convert_ppt_to_image(self):
		nppt = self.ppt_list_table.rowCount()
		if nppt is 0: return

		try:
			Application = win32com.client.Dispatch("PowerPoint.Application")
			#Application.Visible = True
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), 
			QtGui.QMessageBox.Yes)
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
				sld.Export(os.path.join(img_folder,"%s%03d.jpg"%(ff[0],i)), "JPG")
		Application.Quit()
		
		QtGui.QMessageBox.question(QtGui.QWidget(), 'completed!', "%s"%img_folder, 
		QtGui.QMessageBox.Yes)
		
	def convert_ppt_to_pptx(self):
		nppt = self.ppt_list_table.rowCount()
		if nppt is 0: return

		try:
			Application = win32com.client.Dispatch("PowerPoint.Application")
			#Application.Visible = True
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "%s"%str(e), QtGui.QMessageBox.Yes)
			return
			
		for npr in range(nppt):
			path = self.ppt_list_table.item(npr, 2).text()
			sorc = os.path.join(path, self.ppt_list_table.item(npr,0).text())
			Presentation = Application.Presentations.Open(sorc)
			fname = os.path.splitext(self.ppt_list_table.item(npr,0).text())
			dest = os.path.join(path, "%s.pptx"%(fname[0]))
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
		
		QtGui.QMessageBox.question(QtGui.QWidget(), 'completed!', "%s"%dest, QtGui.QMessageBox.Yes)
		
	def set_common_var(self):
		self.ppt_slide = ppt_slide_info()
		self.ppt_textbox = ppt_textbox_info()
		self.ppt_hymal = ppt_hymal_info()
		
	def custom_worship_type(self):
		cid = self.publish_title.currentIndex()
		nwt = len(_worship_type)-1
		cur_wt = self.publish_title.currentText()
		
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
		
	def get_save_directory_path(self, dest):
		startingDir = os.getcwd() 
		self.save_folder = QtGui.QFileDialog.getExistingDirectory(None, 'Save folder', startingDir, QtGui.QFileDialog.ShowDirsOnly)
		if not self.save_folder: return
		
		if dest == _SAVE_FOLDER_SLIDE:
			self.save_directory_path.setText(self.save_folder)
		elif dest == _SAVE_FOLDER_HYMAL:
			self.hymal_save_path.setText(self.save_folder)

	def open_ppt_file(self):
		title = self.open_btn.text()
		self.ppt_filenames = QtGui.QFileDialog.getOpenFileNames(self, title, directory=self.src_directory_path.text(), filter="PPTX (*.pptx);;PPT (*.ppt);;All files (*.*)")
		nppt = len(self.ppt_filenames)

		if nppt: 
			cur_tab = self.tabs.currentIndex()
			if self.tabs.tabText(cur_tab) == _ppttab_text:
				self.clear_pptlist_table()
				self.ppt_list_table.setRowCount(nppt)
				for k in range(nppt):
					fdate = datetime.datetime.fromtimestamp(os.path.getmtime(self.ppt_filenames[k])).strftime("%Y-%m-%d %H:%M:%S").split(' ')
					fpath, fname = os.path.split(self.ppt_filenames[k])
					try:
						prs = pptx.Presentation(self.ppt_filenames[k])
						nslide = len(prs.slides)
					except Exception as e:
						#res = QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', str(e), QtGui.QMessageBox.Yes|QtGui.QMessageBox.Cancel)
						#if res is QtGui.QMessageBox.Cancel:
						#	return
						nslide = 0
						pass
							
					self.ppt_list_table.setItem(k, 0, QtGui.QTableWidgetItem(fname))
					self.ppt_list_table.setItem(k, 1, QtGui.QTableWidgetItem("%d"%nslide))
					self.ppt_list_table.setItem(k, 2, QtGui.QTableWidgetItem(fpath))
					self.ppt_list_table.setItem(k, 3, QtGui.QTableWidgetItem("%d"%_default_txt_nparagraph))
				self.src_directory_path.setText(fpath)
		
	def add_ppt_file(self):
		title = self.add_btn.text()
		files = QtGui.QFileDialog.getOpenFileNames(self, title, directory=self.src_directory_path.text(), filter="PPTX (*.pptx);;PPT (*.ppt);;All files (*.*)")
		nppt = len(files)
		format_error = False
		if nppt:
			
			cur_row = self.ppt_list_table.rowCount()
			for k in range(nppt):
				j = k + cur_row
				fdate = datetime.datetime.fromtimestamp(os.path.getmtime(files[k])).strftime("%Y-%m-%d %H:%M:%S").split(' ')
				fpath, fname = os.path.split(files[k])
				try:
					prs = pptx.Presentation(self.ppt_filenames[k])
				except Exception as e:
					res = QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', e.message, QtGui.QMessageBox.Yes|QtGui.QMessageBox.Cancel)
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
		#bc = self.ppt_slide.back_col
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
		txt = ''.join(p_list)
		return txt

	def create_liveppt(self):
		import copy
		nppt = self.ppt_list_table.rowCount()
		if nppt is 0: return

		if self.add_publish_date.isChecked():
			save_file = "%s %s.pptx"%(self.publish_date.text(), self.publish_title.currentText())
		else:
			save_file = "%s.pptx"%(self.publish_title.currentText())
			
		dest_ppt = pptx.Presentation()
		dest_ppt.slide_width = pptx.util.Inches(self.ppt_slide.wid)
		dest_ppt.slide_height = pptx.util.Inches(self.ppt_slide.hgt)
		blank_slide_layout = dest_ppt.slide_layouts[6]

		#self.add_empty_slide(dest_ppt, blank_slide_layout, self.ppt_slide.back_col)
		for npr in range(nppt):
			npf = os.path.join(self.ppt_list_table.item(npr, 2).text(), self.ppt_list_table.item(npr,0).text())
			npg = int(self.ppt_list_table.item(npr, 3).text())
			src = pptx.Presentation(npf)
			for slide in src.slides:
				in_text = self.slide_has_text(slide)
				if in_text=='' and self.ppt_slide.deep_copy:
					self.add_empty_slide(dest_ppt, blank_slide_layout, self.ppt_slide.back_col)
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
						
						# skip hymal chapter info: (찬송기 123장)
						match = _skip_hymal_info.search(line_text)
						if match or not line_text: continue
						p_list.append(line_text.strip())
					
					np_list = len(p_list)
					for j in range(0, np_list, npg):
						if npg > 1 and (np_list-j) < npg: break
						dest_slide = self.add_empty_slide(dest_ppt, blank_slide_layout, self.ppt_slide.back_col)
						txt_box = self.add_textbox(dest_slide,
						                           self.ppt_textbox.sx,
												   self.ppt_textbox.sy,
												   self.ppt_textbox.wid,
												   self.ppt_textbox.hgt)
						txt_f = txt_box.text_frame
						self.set_textbox(txt_f, MSO_AUTO_SIZE.NONE, MSO_ANCHOR.BOTTOM, MSO_ANCHOR.MIDDLE)
		
						k_list = []
						for k in range(npg):
							k_list.append(p_list[j+k])
						if self.ppt_textbox.word_wrap:
							for w in k_list:
								p = txt_f.add_paragraph()
								p.text = w
								self.set_paragraph(p, 
									PP_ALIGN.CENTER,
									self.ppt_textbox.font_name,
									self.ppt_textbox.font_size,
									self.ppt_textbox.font_col,
									self.ppt_textbox.font_bold)
						else:	
							p = txt_f.add_paragraph()
							p.text = ' '.join(k_list)
							self.set_paragraph(p,
									PP_ALIGN.CENTER,
									self.ppt_textbox.font_name,
									self.ppt_textbox.font_size,
									self.ppt_textbox.font_col,
									self.ppt_textbox.font_bold)

					if npg == 1: continue
					left_over = np_list%npg
					if left_over:
						dest_slide = self.add_empty_slide(dest_ppt, blank_slide_layout, self.ppt_slide.back_col)
						txt_box = self.add_textbox(dest_slide,
						                           self.ppt_textbox.sx,
												   self.ppt_textbox.sy,
												   self.ppt_textbox.wid,
												   self.ppt_textbox.hgt)
						txt_f = txt_box.text_frame
						self.set_textbox(txt_f,MSO_AUTO_SIZE.NONE, MSO_ANCHOR.BOTTOM, MSO_ANCHOR.MIDDLE)
						k_list = []
						l = j-npg
						for k in range(l):
							k_list.append(p_list[l+k])
							
						if self.ppt_textbox.word_wrap:
							for w in k_list:
								p = txt_f.add_paragraph()
								p.text = w
								self.set_paragraph(p, 
									PP_ALIGN.CENTER,
									self.ppt_textbox.font_name,
									self.ppt_textbox.font_size,
									self.ppt_textbox.font_col,
									self.ppt_textbox.font_bold)
						else:
							p = txt_f.add_paragraph()
							p.text = ' '.join(k_list)
							self.set_paragraph(p, 
									PP_ALIGN.CENTER,
									self.ppt_textbox.font_name,
									self.ppt_textbox.font_size,
									self.ppt_textbox.font_col,
									self.ppt_textbox.font_bold)
			self.add_empty_slide(dest_ppt, blank_slide_layout, self.ppt_slide.back_col)
		try:
			sfn = os.path.join(self.save_directory_path.text(),save_file)
			dest_ppt.save(sfn)
		except Exception as e:
			QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', "{}".format(str(e)), QtGui.QMessageBox.Yes)
			#del dest_ppt
			return

		return sfn

	def run_subtitle(self):
		sfn = self.create_liveppt()
		if self.text_outline.isChecked():
			self.create_outline_text(sfn)
		QtGui.QMessageBox.question(QtGui.QWidget(), 'completed!', "{}".format(sfn), QtGui.QMessageBox.Yes)

	def add_textbox(self, ds, sx, sy, wid, hgt):
		return ds.shapes.add_textbox(\
			pptx.util.Inches(sx),
			pptx.util.Inches(sy),
			pptx.util.Inches(wid),
			pptx.util.Inches(hgt)
		)
		
	# slide(az,va,ha): MSO_AUTO_SIZE.NONE, MSO_ANCHOR.BOTTOM, MSO_ANCHOR.MIDDLE
	# hymal(az,va,ha): MSO_AUTO_SIZE.NONE, MSO_ANCHOR.BOTTOM, MSO_ANCHOR.LEFT
	def set_textbox(self, tf, az, va, ha):
		tf.auto_size = az
		tf.vertical_anchor = va
		tf.horizontal_anchor= ha
	
	# slide(al,fn,fz,fc,bl) = PP_ALIGN.CENTER, self.ppt_textbox.font_name,
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

def main():
	app = QtGui.QApplication(sys.argv)
	#QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(u'Motif'))
	#QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(u'CDE'))
	QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(u'Plastique'))
	#QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(u'Cleanlooks'))
	lppt= QLivePPT()
	sys.exit(app.exec_())
	
if __name__ == '__main__':
	main()	