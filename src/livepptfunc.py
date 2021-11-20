import re
from livepptconst import _pp_align
from livepptcolor import ppt_color

_skip_hymal_info = re.compile('찬송가')
_find_lyric_number = re.compile('\d\.')
_find_rgb = re.compile("(\d{1,3}),\s*(\d{1,3}),\s*(\d{1,3})")

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

def access_denied(e_str):
    key = ["access", "denied", "used", "another"]
    return any(x in e_str.lower() for x in key)

