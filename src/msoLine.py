# msoLine
# 

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

style_to_name = {\
"1": "msoLineSingle",
"2": "msoLineThinThin",
"3": "msoLineThinThick",
"4": "msoLineThickThin",
"5": "msoLineThickBetweenThin",
}

def get_linestyle_name(style):
	return style_to_name[str(style)]
