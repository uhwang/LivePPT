# get_hymal.py

import re
import sqlite3 as db

table_name = 'hymnal'
title_chap_py = 'titchap.py'
max_hymal = 639
ref_of_responsive_reading = 700

#_default_db_file = "newhymal.hdb"
#_default_db_file = "새찬송가.hdb"
_default_db_file = "newhymal-v.0.0.1.hdb"

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
    h7 =[]
    if h5[0].find(_corus_delimiter) >= 0:
        h6 = h5[0].split(_corus_delimiter)
        corus = _corus_delimiter + h6[1].strip()
        h7.append([h6[0]])
        h7.append([corus])
        
        # add corus to the other verses
        for ih in range(1, len(h5)):
            h7.append([h5[ih]])
            h7.append([corus])
    else:
        h7 = [[h6] for h6 in h5]
    '''
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
    '''
    #         verse 1        corus           verse 2
    # h7=[['..\n..\n'],['[후렴]..\n..\n'],['..\n..\n'],['[후렴]..\n..\n']]
    #
    hymal_list = []
    for h8 in h7:
        v = h8[0].split('\n')
        if v[-1] == '': del v[-1]
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
    
    try:
        db_con = db.connect(db_file)
        db_cur = db_con.cursor()
    except Exception as e:
        e_str = "... Error(get_responsive_reading_by_chapter): DB error\n%s"%str(e)
        print(e_str)
        return -1, e_str
        
    new_num = num +ref_of_responsive_reading
    sql = 'SELECT title, htext from {} where id = {}'.format(table_name, new_num)
    rtup = db_cur.execute(sql).fetchone()

    return rtup[0], parse_hymal(rtup[1]) 
    
def get_hymal_by_chapter(num, db_file):
    
    try:
        db_con = db.connect(db_file)
        db_cur = db_con.cursor()
        db_tbls= db_cur.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()
        db_tbl = db_tbls[0][0]
    except Exception as e:
        e_str = "... Error(get_hymal_by_chapter): DB error\n\n... Check DB file size!\n...%s"%str(e)
        print(e_str)
        return -1, e_str
        
    #if db_tbl == '':
    #    e_str = "... Error(get_hymal_by_chapter): no table exist!\n... Check DB file size!"
    #    return -1, e_str
        
    sql = 'SELECT title, htext from {} where id = {}'.format(db_tbl, num)
    #sql = 'SELECT title, htext from {} where id = {}'.format(table_name, num)
    htup = db_cur.execute(sql).fetchone()
    #htup = cur.fetchone()
    return htup[0], parse_hymal(htup[1])

def get_hymal_by_title(title, db_file):
    
    try:
        db_con = db.connect(db_file)
        db_cur = db_con.cursor()
        db_tbls= db_cur.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()
        db_tbl = db_tbls[0][0]   
    except Exception as e:
        e_str = "... Error(get_hymal_by_chapter): DB error\n%s"%str(e)
        print(e_str)
        return -1, e_str

    if db_tbl == '':
        e_str = "... Error(get_hymal_by_chapter): no table exist!\n... Check DB file size!"
        return -1, e_str
        
    sql = 'SELECT htext from {} where title = {}'.format(db_tbl, title)
    #sql = 'SELECT htext from {} where title = {}'.format(table_name, title)
    cur = cursor.execute(sql)
    return parse_hymal(cur.fetchone()[0])
    
def get_hymal_by_keyword(keyword):
    return
	
#print(get_hymal_by_chapter(250, _default_db_file))
#print(get_responsive_reading_by_chapter(701))
#print(parse_hymal("<b>1</b>. 내 평생에 가는길 순탄하여 <br>늘 잔잔한강 같든지 큰 풍파로 <br>무섭고 어렵든지나의 영혼은 늘 편하다<br><small><font color='#0099CC'>[후렴]</font></small> 내 영혼 평안해 <br>내 영혼 내영혼 평안해<br><br><b>2</b>. 저 마귀는 우리를 삼키려고<br>입 벌리고 달려와도 예수는<br>우리의 대장되니 끝내 싸워서 이기리라<br><br><b>3</b>. 내 지은 죄 주홍빛같더라도 <br>주 예수께 다 아뢰면 그 십자가<br>피로써 다 씻으사 흰눈보다 정하리라<br><br><b>4</b>. 저 공중에 구름이 일어나며<br>큰 나팔이 울릴때에 주 오셔서 <br>세상을 심판해도 나의 영혼은 겁 없으리<br>"))
#print(parse_hymal("<b>1</b>. 내 평생에 가는길 순탄하여 <br>늘 잔잔한강 같든지 큰 풍파로 <br>무섭고 어렵든지나의 영혼은 늘 편하다<br><small><font color='#0099CC'>[후렴]</font></small> 내 영혼 평안해 <br>내 영혼 내영혼 평안해<br><br>"))
