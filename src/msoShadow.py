from collections import OrderedDict

msoShadow1 = 1	
msoShadow2	= 2	
msoShadow3	= 3	
msoShadow4	= 4	
msoShadow5	= 5	
msoShadow6	= 6	
msoShadow7	= 7	
msoShadow8	= 8	
msoShadow9	= 9	
msoShadow10	= 10	
msoShadow11	= 11	
msoShadow12	= 12	
msoShadow13	= 13	
msoShadow14	= 14	
msoShadow15	= 15	
msoShadow16	= 16	
msoShadow17	= 17	
msoShadow18	= 18	
msoShadow19	= 19	
msoShadow20	= 20	
msoShadow21	= 21	
msoShadow22	= 22	
msoShadow23	= 23	
msoShadow24	= 24	
msoShadow25	= 25	
msoShadow26	= 26	
msoShadow27	= 27	
msoShadow28	= 28	
msoShadow29	= 29	
msoShadow30	= 30	
msoShadow31	= 31	
msoShadow32	= 32	
msoShadow33	= 33	
msoShadow34	= 34	
msoShadow35	= 35	
msoShadow36	= 36	
msoShadow37	= 37	
msoShadow38	= 38	
msoShadow39	= 39	
msoShadow40	= 40	
msoShadow41	= 41	
msoShadow42	= 42	
msoShadow43	= 43	

#Not supported.
#msoShadowMixed	-2	
_shadow_type = OrderedDict()
_shadow_type["1"]  = "Inner"     
_shadow_type["2"]  = "Outer"
_shadow_type["3"]  = "Mixed"

_default_shadow_type = 2

def get_shadowstyle_list():
	return list(_shadow_type.values())
	
def get_shadowstyle_name(idx):
	return _shadow_type[str(idx)]
