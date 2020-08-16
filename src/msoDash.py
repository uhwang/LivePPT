from collections import OrderedDict

_max_dash_style = 11

msoLineSolid         = 1	
msoLineSquareDot     = 2	
msoLineRoundDot      = 3	
msoLineDash          = 4	
msoLineDashDot       = 5 
msoLineDashDotDot    = 6	
msoLineLongDash      = 7	
msoLineLongDashDot   = 8	
msoLineLongDashDotDot= 9	
msoLineSysDash       = 10	
msoLineSysDashDot    = 12	
msoLineSysDot        = 11	

_dasy_type = OrderedDict()

_dasy_type["1"]  = "msoLineSolid"         
_dasy_type["2"]  = "msoLineSquareDot"     
_dasy_type["3"]  = "msoLineRoundDot"      
_dasy_type["4"]  = "msoLineDash"     
_dasy_type["5"]  = "msoLineDashDot"
_dasy_type["6"]  = "msoLineDashDotDot"    
_dasy_type["7"]  = "msoLineLongDash"   
_dasy_type["8"]  = "msoLineLongDashDot"   
_dasy_type["9"]  = "msoLineLongDashDotDot"
_dasy_type["10"] = "msoLineSysDash"
_dasy_type["12"] = "msoLineSysDashDot"    
_dasy_type["11"] = "msoLineSysDot"   

def index_to_style(idx):
	return idx+1
	
def get_dashstyle_list():
	style = []
	for key, val in _dasy_type.items():
		style.append(val[7:])
	return style

def get_dashstyle_name(idx):
	return _dasy_type[str(idx)]