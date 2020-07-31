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
	
	Distribute
	pyinstaller --clean --onefile --hidden-import pptx --exclude-module numpy --exclude-module matplotlib liveppt.py
	
	
'''

import re
import os
import datetime
import pptx
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
import sys
from PyQt4 import QtCore, QtGui

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

_slide_size_type = [
	"[ 4:3],[10:7.5  ]",
	"[16:9],[10:5.625]",
	"[16:9],[13.3:7.5]"
]

_skip_hymal_info = re.compile('[\(\d\)]')

_default_txt_sx = 2.25
_default_txt_sy = 4.66
_default_txt_wid = 5.51
_default_txt_hgt = 0.4
_default_slide_bak_col = (0,32,96)
_default_font_size = 20 # point
_default_font_name = "맑은고딕"
_default_txt_nparagraph = 1
_default_slide_size_index = 1

_ppttab_text = "PPT"
_slidetab_text = "Slide"

_worship_type = ["주일예배", "수요예배", "새벽기도", 
                 "부흥회"  , "특별예베", "직접입력"]

def get_slide_size(t):
	t1 = t.split(',')
	t2 = t1[1][1:-1].split(':')
	return float(t2[0]), float(t2[1])

class ppt_col:
	def __init__(self, col = (255,255,255)):
		self.r = col[0]
		self.g = col[1]
		self.b = col[2]
		
class ppt_textbox_info:
	def __init__(self):
		self.sx = _default_txt_sx
		self.sy = _default_txt_sy
		self.wid = _default_txt_wid
		self.hgt = _default_txt_hgt
		self.font_name = _default_font_name
		self.font_col  = ppt_col()
		self.font_bold = True
		self.font_size = _default_font_size
		self.nparagraph = _default_txt_nparagraph
		self.paragraph_wrap = False

class ppt_slide_info:
	def __init__(self):
		w, h = get_slide_size(_slide_size_type[_default_slide_size_index])
		self.slide_bak_col = ppt_col(_default_slide_bak_col)
		self.slide_wid = w
		self.slide_hgt = h
		self.skip_image = True
		self.slide_size_index = _default_slide_size_index

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
		self.tabs.addTab(self.ppt_tab, _ppttab_text)
		self.tabs.addTab(self.slide_tab, _slidetab_text)

		self.ppt_tab_UI()
		self.slide_tab_UI()
		tab_layout.addWidget(self.tabs)
		self.form_layout.addRow(tab_layout)
		self.setLayout(self.form_layout)
		
		self.setWindowTitle("PPT")
		#self.setWindowIcon(QtGui.QIcon(QtGui.QPixmap(icon_encoder.table)))
		self.show()

	# Background color
	# 
	def slide_tab_UI(self):
		layout = QtGui.QFormLayout()

		background_layout  = QtGui.QHBoxLayout()
		background_layout.addWidget(QtGui.QLabel('Background(RGB)')) 
		self.background_col = QtGui.QLineEdit()
		self.background_col.setInputMask('999|999|999;-')
		font = QtGui.QFont("Courier",11,True)
		fm = QtGui.QFontMetrics(font)
		self.background_col.setFixedSize(fm.width("888888888888"), fm.height())
		self.background_col.setFont(font)
		self.background_col_picker = QtGui.QPushButton('', self)
		self.background_col_picker.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
		self.background_col_picker.setIconSize(QtCore.QSize(16,16))
		self.connect(self.background_col_picker, QtCore.SIGNAL('clicked()'), self.pick_background_col)
		background_layout.addWidget(self.background_col)
		background_layout.addWidget(self.background_col_picker)
		
		slide_layout1 = QtGui.QHBoxLayout()
		slide_layout1.addWidget(QtGui.QLabel("Custom Slide Size"))
		self.custom_slide_size = QtGui.QCheckBox()
		self.custom_slide_size.stateChanged.connect(self.custom_slide_state_changed)
		self.custom_slide_size.setChecked(False)
		slide_layout1.addWidget(self.custom_slide_size)
		
		slide_layout2 = QtGui.QHBoxLayout()
		self.choose_slide_size = QtGui.QComboBox(self)
		self.choose_slide_size.addItems(_slide_size_type)
		self.choose_slide_size.setCurrentIndex(self.ppt_slide.slide_size_index)
		self.choose_slide_size.currentIndexChanged.connect(self.set_custom_slide_size)

		self.custom_slide_wid = QtGui.QLineEdit()
		self.custom_slide_hgt = QtGui.QLineEdit()
		self.custom_slide_wid.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		self.custom_slide_hgt.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
		self.custom_slide_wid.setEnabled(False)
		self.custom_slide_hgt.setEnabled(False)
		self.custom_slide_wid.setText("%f"%self.ppt_slide.slide_wid)
		self.custom_slide_hgt.setText("%f"%self.ppt_slide.slide_hgt)

		slide_layout2.addWidget(self.choose_slide_size)
		slide_layout2.addWidget(self.custom_slide_wid)
		slide_layout2.addWidget(self.custom_slide_hgt)
		
		#self.sx = _default_txt_sx
		#self.sy = _default_txt_sy
		#self.wid = _default_txt_wid
		#self.hgt = _default_txt_hgt
		#self.font_name = _default_font_name
		#self.font_col  = ppt_col()
		#self.font_bold = True
		#self.font_size = _default_font_size
		#self.nparagraph = _default_txt_nparagraph
		#self.paragraph_wrap = False
			
		layout.addRow(background_layout)
		layout.addRow(slide_layout1)
		layout.addRow(slide_layout2)
		self.slide_tab.setLayout(layout)
	
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
		
	def pick_background_col(self):
		return
		
	def ppt_tab_UI(self):
		layout = QtGui.QFormLayout()
		open_layout  = QtGui.QHBoxLayout()
		self.add_button = QtGui.QPushButton('Add', self)
		self.add_button.clicked.connect(self.add_ppt_file)
		self.add_button.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_file_add.table)))
		self.add_button.setIconSize(QtCore.QSize(16,16))
		self.open_button = QtGui.QPushButton('Open', self)
		self.open_button.clicked.connect(self.open_ppt_file)
		self.open_button.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_folder_open.table)))
		self.open_button.setIconSize(QtCore.QSize(16,16))
		open_layout.addWidget(self.open_button)
		open_layout.addWidget(self.add_button)

		self.ppt_list_table = QtGui.QTableWidget()
		self.ppt_list_table.setColumnCount(4)
		self.ppt_list_table.setHorizontalHeaderItem(0, QtGui.QTableWidgetItem("Name"))
		self.ppt_list_table.setHorizontalHeaderItem(1, QtGui.QTableWidgetItem("Slide"))
		self.ppt_list_table.setHorizontalHeaderItem(2, QtGui.QTableWidgetItem("Path"))
		self.ppt_list_table.setHorizontalHeaderItem(3, QtGui.QTableWidgetItem("Parg"))
		#self.ppt_list_table.setHorizontalHeaderItem(4, QtGui.QTableWidgetItem("VCodec"))
		#self.ppt_list_table.setHorizontalHeaderItem(5, QtGui.QTableWidgetItem("ACodec"))
	
		header = self.ppt_list_table.horizontalHeader()
		header.setResizeMode(1, QtGui.QHeaderView.ResizeToContents)
		header.setResizeMode(3, QtGui.QHeaderView.ResizeToContents)
		
		self.move_up = QtGui.QPushButton('', self)
		self.move_up.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_arrow_up.table)))
		self.move_up.setIconSize(QtCore.QSize(16,16))
		self.connect(self.move_up, QtCore.SIGNAL('clicked()'), self.move_item_up)
		 
		self.move_down = QtGui.QPushButton('', self)
		self.move_down.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_arrow_down.table)))
		self.move_down.setIconSize(QtCore.QSize(16,16))
		self.connect(self.move_down, QtCore.SIGNAL('clicked()'), self.move_itme_down)
		
		self.sort_asc = QtGui.QPushButton('', self)
		self.sort_asc.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_table_sort_asc.table)))
		self.sort_asc.setIconSize(QtCore.QSize(16,16))

		self.sort_desc = QtGui.QPushButton('', self)
		self.sort_desc.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_table_sort_desc.table)))
		self.sort_desc.setIconSize(QtCore.QSize(16,16))

		self.delete_video = QtGui.QPushButton('', self)
		self.delete_video.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_delete.table)))
		self.delete_video.setIconSize(QtCore.QSize(16,16))
		self.connect(self.delete_video, QtCore.SIGNAL('clicked()'), self.delete_item)
		
		self.delete_all = QtGui.QPushButton('', self)
		self.delete_all.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_trash.table)))
		self.delete_all.setIconSize(QtCore.QSize(16,16))
		
		self.play = QtGui.QPushButton('', self)
		self.play.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_play.table)))
		self.play.setIconSize(QtCore.QSize(16,16))
		
		self.connect(self.delete_all, QtCore.SIGNAL('clicked()'), self.delete_all_item)
		
		btn_layout = QtGui.QHBoxLayout()
		btn_layout.addWidget(self.move_up)
		btn_layout.addWidget(self.move_down)
		btn_layout.addWidget(self.sort_asc)
		btn_layout.addWidget(self.sort_desc)
		btn_layout.addWidget(self.delete_video)
		btn_layout.addWidget(self.delete_all)
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
		#self.publish_title.setEditable(True)
		self.publish_title.addItems(_worship_type)
		self.publish_title.currentIndexChanged.connect(self.custom_worship_type)
		publish_layout.addWidget(self.publish_title, 2, 1)

		publish_layout.addWidget(QtGui.QLabel('Srce'), 3, 0)
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
		self.run_convert = QtGui.QPushButton('Run', self)
		self.run_convert.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_convert.table)))
		self.run_convert.setIconSize(QtCore.QSize(48,48))
		self.run_convert.clicked.connect(self.create_liveppt)
		run_layout.addWidget(self.run_convert)
		
		layout.addRow(open_layout)
		layout.addRow(self.ppt_list_table)
		layout.addRow(btn_layout)
		layout.addRow(publish_layout)
		layout.addRow(run_layout)
		self.ppt_tab.setLayout(layout)
				
	def set_common_var(self):
		self.ppt_slide = ppt_slide_info()
		self.ppt_textbox = ppt_textbox_info()
		
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
		
	def get_save_directory_path(self):
		startingDir = os.getcwd() 
		self.save_folder = QtGui.QFileDialog.getExistingDirectory(None, 'Save folder', startingDir, QtGui.QFileDialog.ShowDirsOnly)
		if not self.save_folder: return
		self.save_directory_path.setText(self.save_folder)
		#self.message.appendPlainText("... Folder path : {0}".format(self.save_folder))
		
	def open_ppt_file(self):
		title = self.open_button.text()
		self.ppt_filenames = QtGui.QFileDialog.getOpenFileNames(self, title, directory=self.src_directory_path.text(), filter="PPT (*.pptx);;All files (*.*)")
		nppt = len(self.ppt_filenames)
		#format_error = False
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
					except Exception as e:
						res = QtGui.QMessageBox.question(QtGui.QWidget(), 'Error', e.message, QtGui.QMessageBox.Yes|QtGui.QMessageBox.Cancel)
						if res is QtGui.QMessageBox.Cancel:
							return
							
					self.ppt_list_table.setItem(k, 0, QtGui.QTableWidgetItem(fname))
					self.ppt_list_table.setItem(k, 1, QtGui.QTableWidgetItem("%s"%len(prs.slides)))
					self.ppt_list_table.setItem(k, 2, QtGui.QTableWidgetItem(fpath))
					self.ppt_list_table.setItem(k, 3, QtGui.QTableWidgetItem("%d"%_default_txt_nparagraph))
					#self.video_list_table.setItem(k, 5, QtGui.QTableWidgetItem(mi.audiocodec))
				self.src_directory_path.setText(fpath)
		
	def add_ppt_file(self):
		title = self.add_button.text()
		files = QtGui.QFileDialog.getOpenFileNames(self, title, directory=self.src_directory_path.text(), filter="PPT (*.pptx);;All files (*.*)")
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
				#self.ppt_list_table.setItem(k, 4, QtGui.QTableWidgetItem(mi.videocodec))
				#self.ppt_list_table.setItem(k, 5, QtGui.QTableWidgetItem(mi.audiocodec))
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
	
	def add_empty_slide(self, dest_ppt, layout):
		dest_slide = dest_ppt.slides.add_slide(layout)
		bk = dest_slide.background
		bk.fill.solid()
		bc = self.ppt_slide.slide_bak_col
		bk.fill.fore_color.rgb = pptx.dml.color.RGBColor(bc.r, bc.g, bc.b)
		return dest_slide

	def create_liveppt(self):
		nppt = self.ppt_list_table.rowCount()
		if nppt is 0: return

		if self.add_publish_date.isChecked():
			save_file = "%s %s.pptx"%(self.publish_date.text(), self.publish_title.currentText())
		else:
			save_file = "%s.pptx"%(self.publish_title.currentText())
			
		dest_ppt = pptx.Presentation()
		dest_ppt.slide_width = pptx.util.Inches(self.ppt_slide.slide_wid)
		dest_ppt.slide_height = pptx.util.Inches(self.ppt_slide.slide_hgt)
		blank_slide_layout = dest_ppt.slide_layouts[6]

		self.add_empty_slide(dest_ppt, blank_slide_layout)
		for npr in range(nppt):
			npf = os.path.join(self.ppt_list_table.item(npr, 2).text(), self.ppt_list_table.item(npr,0).text())
			npg = int(self.ppt_list_table.item(npr, 3).text())
			src = pptx.Presentation(npf)
			for slide in src.slides:
				for shape in slide.shapes:
					if not shape.has_text_frame:
						continue
					p_list = []
					for paragraph in shape.text_frame.paragraphs:
						run_list = []
						for run in paragraph.runs:
							run_list.append(run.text)
						line_text = ''.join(run_list)
						
						match = _skip_hymal_info.search(line_text)
						if match or not line_text: continue
						p_list.append(line_text)
					
					np_list = len(p_list)
					for j in range(0, np_list, npg):
						if (np_list-j) < npg: break
						dest_slide = self.add_empty_slide(dest_ppt, blank_slide_layout)
						txt_box = self.add_textbox(dest_slide)
						txt_f = txt_box.text_frame
						self.set_textbox(txt_f)
		
						k_list = []
						for k in range(npg):
							k_list.append(p_list[j+k])
						
						p = txt_f.add_paragraph()
						p.text = ''.join(k_list)
						self.set_paragraph(p)

					if npg == 1: continue
					left_over = np_list%npg
					if left_over:
						dest_slide = self.add_empty_slide(dest_ppt, blank_slide_layout)
						txt_box = self.add_textbox(dest_slide)
						txt_f = txt_box.text_frame
						self.set_textbox(txt_f)
						k_list = []
						l = j-npg
						for k in range(l):
							k_list.append(p_list[l+k])
						p = txt_f.add_paragraph()
						p.text = ''.join(k_list)
						self.set_paragraph(p)
			self.add_empty_slide(dest_ppt, blank_slide_layout)
		dest_ppt.save(os.path.join(self.save_directory_path.text(),save_file))

	def add_textbox(self, ds):
		return ds.shapes.add_textbox(\
			pptx.util.Inches(self.ppt_textbox.sx),
			pptx.util.Inches(self.ppt_textbox.sy),
			pptx.util.Inches(self.ppt_textbox.wid),
			pptx.util.Inches(self.ppt_textbox.hgt)
		)
		
	def set_textbox(self, tf):
		tf.auto_size = MSO_AUTO_SIZE.NONE
		tf.vertical_anchor = MSO_ANCHOR.BOTTOM
		tf.horizontal_anchor= MSO_ANCHOR.MIDDLE
	
	def set_paragraph(self, p):
		p.alignment = PP_ALIGN.CENTER
		p.font.name = self.ppt_textbox.font_name
		p.font.size = pptx.util.Pt(self.ppt_textbox.font_size)
		fc = self.ppt_textbox.font_col
		p.font.color.rgb = pptx.dml.color.RGBColor(fc.r, fc.g, fc.b)
		p.font.bold = self.ppt_textbox.font_bold

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