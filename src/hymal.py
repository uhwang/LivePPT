# get_hymal.py

import re
import sqlite3 as db

table_name = 'hymnal'
title_chap_py = 'titchap.py'
max_hymal = 639
ref_of_responsive_reading = 700

#_default_db_file = "newhymal.hdb"
_default_db_file = "새찬송가.hdb"

def create_hyaml_py(db_file):
    '''
    Create a dictionary of (title, number) pair
    
        Find max number of hymal from DB
    Responsive Reading is attached at the end of hymal
    The max number of hymal is 639
    
    cursor.execute("SELECT * FROM {}".format(table_name))
    nhym = len(cursor.fetchall())+1
    
    cur = cursor.execute()
    cur is a tuple
    '''
    conn = db.connect(db_file)
    cursor = conn.cursor()
    
    tn_file = open(title_chap_py, "wt")
    tn_file.write("title_chap = dict()\n")
    nhym = max_nhymal + 1
    
    for i in range(1, nhym):
        sql = 'SELECT title from {} where id = {}'.format(table_name, i)
        cur = cursor.execute(sql)
        tn_file.write("title_chap[\"{}\"] = {}\n".format(cur.fetchone()[0], i))
    tn_file.close()

# usage	
#create_title_to_num("새찬송가.hdb")

# <b>1</b> number of lyrics
# <br> new line

_corus_delimiter = '[후렴]'
_find_nlyric = re.compile("\d")
_amen = "아멘"
'''
def parse_hymal(hymal_str):
    h1 = hymal_str.replace('<b>', '')
    h2 = h1.replace('</b>', '')
    h3 = h2.replace('<br>', '\n')
    h4 = h3.replace('\n\n', '\n')
    s1=re.sub(r"(<.*?>)(?!<)", '', h4)
    s2=re.sub(r"\[.*?\]", '', s1)
    #s2 = s1
    #h2 = re.sub("[<b>][</b]", hymal_str, '')
    #h3 = re.sub("[<br>]", h2, '\n')
    #h4 = re.sub('\n\n', h3, '\n')
    h5 = re.split('\d. ', s2)
    if h5[0] == '': del h5[0]
    
    hymal_list = []
    nl = len(h5)
    for i in range(nl):
        v = h5[i].split('\n')
        del v[-1]
        for j in range(len(v)): 
            v[j] = v[j].strip()
        hymal_list.append(v)
    
    return hymal_list
'''
def parse_hymal(hymal_str):
    h1 = hymal_str.replace('<b>', '')
    h2 = h1.replace('</b>', '')
    h3 = h2.replace('<br>', '\n')
    h4 = h3.replace('\n\n', '\n')
    s1=re.sub(r"(<.*?>)(?!<)", '', h4)
    h5 = re.split('\d. ', s1)

    # delete empty element 
    if h5[0] == '': del h5[0]
    
    # ----------------------------------------------------------
    # Verse 1                Corus           Verse 2    Verse 3
    # ----------------------------------------------------------
    # ['????? \n???? \n??? \n[후렴] ????', '???? \n', '???? \n']
    # -----------------------------------------------------------
 
    # find corus
    if h5[0].find(_corus_delimiter) >= 0:
        h6 = h5[0].split(_corus_delimiter)
        print(h6)
        h5.pop(0)
        h5.insert(0, ''.join(h6))
        
        # add corus to the other verses
        for ih in range(1, len(h5)):
            h5[ih] = ''.join([h5[ih], h6[1]])
        
    hymal_list = []
    nl = len(h5)
    for i in range(nl):
        v = h5[i].split('\n')
        del v[-1]
        for j in range(len(v)):
            vj = v[j].strip()
            if vj[-2:] == _amen:
                vj = vj[0:-2]
            v[j] = vj #v[j].strip()
        hymal_list.append(v)
    
    return hymal_list
    
# Responsive Reading starts from chapter 701
# make sure the num is between 701 and 836
def get_responsive_reading_by_chapter(num, db_file):
    
    conn = db.connect(db_file)
    cursor = conn.cursor()
    new_num = num +ref_of_responsive_reading
    sql = 'SELECT title, htext from {} where id = {}'.format(table_name, new_num)
    cur = cursor.execute(sql)
    rtup = cur.fetchone()
    return rtup[0], parse_hymal(rtup[1]) 
    
def get_hymal_by_chapter(num, db_file):
    
    #try:
    db_con = db.connect(db_file)
    db_cur = db_con.cursor()
    db_tbls= db_cur.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()

    db_tbl = db_tbls[0][0]
    sql = 'SELECT htext from {} where id = {}'.format(db_tbl, num)
    #sql = 'SELECT htext from {} where id = {}'.format(table_name, num)
    cur = db_cur.execute(sql)
    return parse_hymal(db_cur.fetchone()[0])

def get_hymal_by_title(title, db_file):
    
    db_con = db.connect(db_file)
    db_cur = db_con.cursor()
    db_tbls= db_cur.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()
    db_tbl = db_tbls[0][0]   
    sql = 'SELECT htext from {} where title = {}'.format(db_tbl, title)
    #sql = 'SELECT htext from {} where title = {}'.format(table_name, title)
    cur = cursor.execute(sql)
    return parse_hymal(cur.fetchone()[0])
    
def get_hymal_by_keyword(keyword):
    return
	
#print(get_hymal_by_chapter(10, _default_db_file))
#print(get_responsive_reading_by_chapter(701))
#print(parse_hymal("<b>1</b>. 내 평생에 가는길 순탄하여 <br>늘 잔잔한강 같든지 큰 풍파로 <br>무섭고 어렵든지나의 영혼은 늘 편하다<br><small><font color='#0099CC'>[후렴]</font></small> 내 영혼 평안해 <br>내 영혼 내영혼 평안해<br><br><b>2</b>. 저 마귀는 우리를 삼키려고<br>입 벌리고 달려와도 예수는<br>우리의 대장되니 끝내 싸워서 이기리라<br><br><b>3</b>. 내 지은 죄 주홍빛같더라도 <br>주 예수께 다 아뢰면 그 십자가<br>피로써 다 씻으사 흰눈보다 정하리라<br><br><b>4</b>. 저 공중에 구름이 일어나며<br>큰 나팔이 울릴때에 주 오셔서 <br>세상을 심판해도 나의 영혼은 겁 없으리<br>"))