from PyQt4 import QtCore, QtGui

import copy
import icon_color_picker
import icon_font_picker

from livepptcls import *
import livepptconst as const
import livepptcolor as color
import livepptfunc as func

class QLivepptSlideDataDlg(QtGui.QDialog):
    def __init__(self, slide_data, sub_title):
        super(QLivepptSlideDataDlg, self).__init__()
        self.slide_data = slide_data
        self.sub_title = sub_title
        self.initUI(slide_data)

    def initUI(self, slide_data):
        layout = QtGui.QFormLayout()

        slide_layout2 = QtGui.QHBoxLayout()
        self.choose_slide_size = QtGui.QComboBox(self)
        self.choose_slide_size.addItems(const._slide_size_type)
        self.choose_slide_size.setCurrentIndex(slide_data.slide.size_index)
        self.choose_slide_size.currentIndexChanged.connect(self.slide_size_changed)
        
        self.custom_slide_wid = QtGui.QLineEdit()
        self.custom_slide_hgt = QtGui.QLineEdit()
        self.custom_slide_wid.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
        self.custom_slide_hgt.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
        self.custom_slide_wid.setEnabled(False)
        self.custom_slide_hgt.setEnabled(False)
        self.custom_slide_wid.setText("%f"%slide_data.slide.wid)
        self.custom_slide_hgt.setText("%f"%slide_data.slide.hgt)
        
        slide_layout2.addWidget(self.choose_slide_size)
        slide_layout2.addWidget(self.custom_slide_wid)
        slide_layout2.addWidget(self.custom_slide_hgt)
        
        text_layout = QtGui.QGridLayout()
        text_layout.addWidget(QtGui.QLabel('Back(RGB)'), 0, 0) 
        self.slide_back_col = QtGui.QLineEdit(str(slide_data.slide.back_col))
  
        self.back_col_picker = QtGui.QPushButton('', self)
        self.back_col_picker.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
        self.back_col_picker.setIconSize(QtCore.QSize(16,16))
        self.connect(self.back_col_picker, QtCore.SIGNAL('clicked()'), self.pick_slide_back_color)
        text_layout.addWidget(self.slide_back_col, 0, 1)
        text_layout.addWidget(self.back_col_picker, 0, 2)        
        
        text_layout.addWidget(QtGui.QLabel("Textbox(x)"), 1,0)
        self.textbox_sx = QtGui.QLineEdit("%f"%slide_data.textbox.sx)
        self.textbox_sx.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
        text_layout.addWidget(self.textbox_sx, 1, 1)
        text_layout.addWidget(QtGui.QLabel("inch"), 1,2)
    
        text_layout.addWidget(QtGui.QLabel("Textbox(y)"), 2,0)
        self.textbox_sy = QtGui.QLineEdit("%f"%slide_data.textbox.sy)
        self.textbox_sy.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
        text_layout.addWidget(self.textbox_sy, 2, 1)
        text_layout.addWidget(QtGui.QLabel("inch"), 2,2)
        
        text_layout.addWidget(QtGui.QLabel("Textbox(wid)"), 3,0)
        self.textbox_wid = QtGui.QLineEdit("%f"%slide_data.textbox.wid)
        self.textbox_wid.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
        text_layout.addWidget(self.textbox_wid, 3, 1)
        text_layout.addWidget(QtGui.QLabel("inch"), 3,2)
    
        text_layout.addWidget(QtGui.QLabel("Textbox(hgt)"), 4,0)
        self.textbox_hgt = QtGui.QLineEdit("%f"%slide_data.textbox.hgt)
        self.textbox_hgt.setSizePolicy(QtGui.QSizePolicy.Ignored, QtGui.QSizePolicy.Preferred)
        text_layout.addWidget(self.textbox_hgt, 4, 1)
        text_layout.addWidget(QtGui.QLabel("inch"), 4,2)
    
        text_layout.addWidget(QtGui.QLabel("VAnchor"), 5, 0)
        self.choose_textframe_vanchor = QtGui.QComboBox()
        self.choose_textframe_vanchor.addItems(const.get_all_textframe_vanchortype_string())
        self.choose_textframe_vanchor.setCurrentIndex(slide_data.textbox.vanchor) # center
        text_layout.addWidget(self.choose_textframe_vanchor, 5, 1)

        text_layout.addWidget(QtGui.QLabel("PP Align"), 6, 0)
        self.choose_paragraph_align = QtGui.QComboBox()
        self.choose_paragraph_align.addItems(const.get_all_pragraph_aligntype_string())
        self.choose_paragraph_align.setCurrentIndex(slide_data.textbox.pp_align) # center
        text_layout.addWidget(self.choose_paragraph_align, 6, 1)
        
        text_layout.addWidget(QtGui.QLabel("Font(RGB)"), 7, 0)
        self.textbox_font_col = QtGui.QLineEdit(str(slide_data.textbox.font_col))
        self.textbox_font_col_picker = QtGui.QPushButton('', self)
        self.textbox_font_col_picker.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
        self.textbox_font_col_picker.setIconSize(QtCore.QSize(16,16))
        self.connect(self.textbox_font_col_picker, QtCore.SIGNAL('clicked()'), self.pick_font_color)
        text_layout.addWidget(self.textbox_font_col, 7, 1)
        text_layout.addWidget(self.textbox_font_col_picker, 7, 2)
    
        text_layout.addWidget(QtGui.QLabel("Name"), 8,0)
        self.textbox_font_name = QtGui.QLineEdit(slide_data.textbox.font_name)
        text_layout.addWidget(self.textbox_font_name, 8, 1)
        self.textbox_font_picker = QtGui.QPushButton('', self)
        self.textbox_font_picker.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_font_picker.table)))
        self.textbox_font_picker.setIconSize(QtCore.QSize(16,16))
        self.connect(self.textbox_font_picker, QtCore.SIGNAL('clicked()'), self.pick_font)
        text_layout.addWidget(self.textbox_font_picker, 8, 2)
    
        text_layout.addWidget(QtGui.QLabel("Size"), 9,0)
        self.textbox_font_size = QtGui.QLineEdit("%f"%slide_data.textbox.font_size)
        text_layout.addWidget(self.textbox_font_size, 9,1)
    
        option_layout = QtGui.QHBoxLayout()
        if self.sub_title:
            self.textbox_deep_copy = QtGui.QCheckBox("D.Copy")
            self.textbox_deep_copy.setChecked(slide_data.slide.deep_copy)
            option_layout.addWidget(self.textbox_deep_copy)
            self.textbox_word_wrap = QtGui.QCheckBox("Wrap")
            self.textbox_word_wrap.setChecked(slide_data.textbox.word_wrap)
            option_layout.addWidget(self.textbox_word_wrap)

        self.textbox_font_bold = QtGui.QCheckBox("Bold")
        self.textbox_font_bold.setChecked(True)
        option_layout.addWidget(self.textbox_font_bold)
            
        self.textbox_fill = QtGui.QCheckBox("Fill")
        self.textbox_fill.setChecked(True)
        option_layout.addWidget(self.textbox_fill)
        
        self.textbox_fill_type = QtGui.QComboBox()
        self.textbox_fill_type.addItems(["Solid", "Grad"])
        self.textbox_fill_type.setFixedWidth(50)
        self.textbox_fill_type.setCurrentIndex(slide_data.textbox.fill.type)
        self.connect(self.textbox_fill_type, QtCore.SIGNAL("currentIndexChanged(int)"), self.textbox_filltype_changed)
    
        self.textbox_solid_color = QtGui.QLineEdit(str(slide_data.textbox.fill.solid_col))
        self.textbox_solid_color.setFixedWidth(100)
        
        self.textbox_solid_color_btn = QtGui.QPushButton('', self)
        self.textbox_solid_color_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker.table)))
        self.textbox_solid_color_btn.setIconSize(QtCore.QSize(16,16))
        self.connect(self.textbox_solid_color_btn, QtCore.SIGNAL('clicked()'), self.pick_textbox_fill_solid_color)
    
        self.textbox_gradient_color1 = QtGui.QLineEdit(str(slide_data.textbox.fill.gradient_col1))
        self.textbox_gradient_color1.setFixedWidth(72)
        
        self.textbox_gradient_color2 = QtGui.QLineEdit(str(slide_data.textbox.fill.gradient_col2))
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
        self.connect(self.textbox_gradient_color2_btn, QtCore.SIGNAL('clicked()'), self.pick_textbox_fill_gradient_color2)
    
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
        self.text_box_lmargin = QtGui.QLineEdit('%f'%slide_data.textbox.left_margin)
        self.text_box_rmargin = QtGui.QLineEdit('%f'%slide_data.textbox.right_margin)
        self.text_box_lmargin.setFixedWidth(90)
        self.text_box_rmargin.setFixedWidth(90)
        margin_layout1.addWidget(self.text_box_lmargin)
        margin_layout1.addWidget(self.text_box_rmargin)
        
        margin_layout2.addWidget(QtGui.QLabel('T/B Margin'))
        self.text_box_tmargin = QtGui.QLineEdit('%f'%slide_data.textbox.top_margin)
        self.text_box_bmargin = QtGui.QLineEdit('%f'%slide_data.textbox.bottom_margin)
        self.text_box_tmargin.setFixedWidth(90)
        self.text_box_bmargin.setFixedWidth(90)
        margin_layout2.addWidget(self.text_box_tmargin)
        margin_layout2.addWidget(self.text_box_bmargin)
        
        margin_layout.addLayout(margin_layout0)
        margin_layout.addLayout(margin_layout1)
        margin_layout.addLayout(margin_layout2)
    
        menu_layout = QtGui.QHBoxLayout()
        self.slide_menu_cancel_btn = QtGui.QPushButton('cancel')
        self.slide_menu_ok_btn = QtGui.QPushButton('Ok')
        self.slide_menu_ok_btn.clicked.connect(self.accept)
        self.slide_menu_cancel_btn.clicked.connect(self.reject)
        
        menu_layout.addWidget(self.slide_menu_cancel_btn)
        menu_layout.addWidget(self.slide_menu_ok_btn)
                
        layout.addRow(slide_layout2)
        layout.addRow(text_layout)
        layout.addRow(option_layout)
        layout.addRow(margin_layout)
        layout.addRow(menu_layout)
        self.setLayout(layout)
        self.textbox_filltype_changed()

        #self.global_message.appendPlainText('... Slide Tab UI created')
            
    def pick_textbox_fill_solid_color(self):
        fc = self.slide_data.textbox.fill.solid_col
        col = QtGui.QColorDialog.getColor(QtGui.QColor(fc.r, fc.g, fc.b))
        if col.isValid():
            r,g,b,a = col.getRgb()
            self.textbox_solid_color.setText("%03d,%03d,%03d"%(r, g, b))
    
    def pick_textbox_fill_gradient_color1(self):
        gc = self.slide_data.textbox.fill.gradient_col1
        col = QtGui.QColorDialog.getColor(QtGui.QColor(gc.r, gc.g, gc.b))
        if col.isValid():
            r,g,b,a = col.getRgb()
            self.textbox_gradient_color1.setText("%03d,%03d,%03d"%(r, g, b))
    
    def pick_textbox_fill_gradient_color2(self):
        gc = self.slide_data.textbox.fill.gradient_col2
        col = QtGui.QColorDialog.getColor(QtGui.QColor(gc.r, gc.g, gc.b))
        if col.isValid():
            r,g,b,a = col.getRgb()
            self.textbox_gradient_color2.setText("%03d,%03d,%03d"%(r, g, b))
        
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
            
    def pick_slide_back_color(self):
        bc = self.slide_data.slide.back_col
        col = QtGui.QColorDialog.getColor(QtGui.QColor(bc.r, bc.g, bc.b))
        if col.isValid():
            r,g,b,a = col.getRgb()
            self.slide_back_col.setText("%03d,%03d,%03d"%(r, g, b))
    
    def pick_font_color(self):
        fc = self.slide_data.textbox.font_col
        col = QtGui.QColorDialog.getColor(QtGui.QColor(fc.r, fc.g, fc.b))
        if col.isValid():
            r,g,b,a = col.getRgb()
            self.textbox_font_col.setText("%03d,%03d,%03d"%(r, g, b))
        
    def pick_font(self):
        font, valid = QtGui.QFontDialog.getFont()
        if valid:
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
        
    def slide_size_changed(self):
        cur_text = self.choose_slide_size.currentText()
        cur_index = self.choose_slide_size.currentIndex()
        if cur_text == const.get_custom_slide_size_string():
            self.custom_slide_wid.setEnabled(True)
            self.custom_slide_hgt.setEnabled(True)
        else:
            self.custom_slide_wid.setEnabled(False)
            self.custom_slide_hgt.setEnabled(False)
        
        w,h = get_slide_size(cur_text)
        self.custom_slide_wid.setText("%f"%w)
        self.custom_slide_hgt.setText("%f"%h)
        
        # textbox sx : not change
        # textbox hgt: not change
        if self.sub_title: 
            self.textbox_sy.setText('%f'%(h-self.slide_data.textbox.hgt))
            self.textbox_hgt.setText('%f'%self.slide_data.textbox.hgt)
        else:
            self.textbox_sy.setText('%f'%(0.0))
            self.textbox_hgt.setText('%f'%h)
        self.textbox_wid.setText('%f'%w)

    def accept(self):
        self.apply_current_slide_info()
        return QtGui.QDialog.accept(self)
        
    def apply_current_slide_info(self):
        #self.global_message.appendPlainText("... Apply Slide Info")
        self.slide_data.slide.back_col = func.get_rgb(self.slide_back_col.text())
        cur_text = self.choose_slide_size.currentText()
        self.slide_data.slide.size_index = self.choose_slide_size.currentIndex()
        
        if cur_text == const.get_custom_slide_size_string():
            self.slide_data.slide.wid = float(self.custom_slide_wid.text())
            self.slide_data.slide.hgt = float(self.custom_slide_hgt.text())
        else:
            w,h = func.get_slide_size(str(cur_text))
            self.slide_data.slide.wid = w
            self.slide_data.slide.hgt = h
            
        self.slide_data.textbox.sx  = float(self.textbox_sx .text())
        self.slide_data.textbox.sy  = float(self.textbox_sy .text())
        self.slide_data.textbox.wid = float(self.textbox_wid.text())
        self.slide_data.textbox.hgt = float(self.textbox_hgt.text())
    
        self.slide_data.textbox.font_name = self.textbox_font_name.text()
        self.slide_data.textbox.font_col  = func.get_rgb(self.textbox_font_col.text())
        self.slide_data.textbox.font_size = float(self.textbox_font_size.text())
        self.slide_data.textbox.font_bold = self.textbox_font_bold.isChecked()
        self.slide_data.textbox.fill.show = self.textbox_fill.isChecked()
        self.slide_data.textbox.fill.type = self.textbox_fill_type.currentIndex()
        self.slide_data.textbox.fill.solid_col     = func.get_rgb(self.textbox_solid_color.text())
        self.slide_data.textbox.fill.gradient_col1 = func.get_rgb(self.textbox_gradient_color1.text())
        self.slide_data.textbox.fill.gradient_col2 = func.get_rgb(self.textbox_gradient_color2.text())

        if self.sub_title:
            self.slide_data.textbox.word_wrap = self.textbox_word_wrap.isChecked()
            self.slide_data.slide.deep_copy = self.textbox_deep_copy.isChecked()
            
        self.slide_data.textbox.vanchor     = self.choose_textframe_vanchor.currentIndex()
        self.slide_data.textbox.pp_align    = self.choose_paragraph_align.currentIndex()
        
        self.slide_data.textbox.left_margin  = float(self.text_box_lmargin.text())
        self.slide_data.textbox.top_margin   = float(self.text_box_tmargin.text())
        self.slide_data.textbox.right_margin = float(self.text_box_rmargin.text())
        self.slide_data.textbox.bottom_margin= float(self.text_box_bmargin.text())
        
              
def get_slide_data(slide_data, sub_title=True):
    run = QLivepptSlideDataDlg(slide_data, sub_title)
    #print(res.exec_())
    res = run.exec_()
    if res == 0:
        return res
    
    #slide_data = copy.deepcopy(res.temp_slide_data)
   