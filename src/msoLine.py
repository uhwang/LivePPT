# msoLine
# 

from collections import OrderedDict

# Single line.
msoLineSingle = 1	

#Not supported.
msoLineStyleMixed =	-2	

#Thick line with a thin line on each side.
msoLineThickBetweenThin	= 5	

# Thick line next to thin line. For horizontal lines, 
# thick line is above thin line. For vertical lines, 
# thick line is to the left of the thin line.
msoLineThickThin = 4	

# Thick line next to thin line. For horizontal lines, 
# thick line is below thin line. For vertical lines, 
# thick line is to the right of the thin line.
msoLineThinThick = 3	

# Two thin lines.
msoLineThinThin = 2

_line_type = OrderedDict()

_line_type["1"]= "msoLineSingle"
_line_type["2"]= "msoLineThinThin"
_line_type["3"]= "msoLineThinThick"
_line_type["4"]= "msoLineThickThin"
_line_type["5"]= "msoLineThickBetweenThin"

def get_linestyle_list():
	style = []
	for key, val in _line_type.items():
		style.append(val[7:])
	return style
	
# make sure the idx starts from 1
def get_linestyle_name(idx):
	return _line_type[str(idx)]
	
def index_to_style(idx):
	return idx+1
