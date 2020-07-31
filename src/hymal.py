# get_hymal.py

import re
import sqlite3 as db

table_name = 'hymnal'
title_num_py = 'titnum.py'
#_find_title = re.compile('\'.*\'')
max_nhymal = 639

db_file = "newhymal.hdb"

def create_title_to_num(db_file):
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
	
	tn_file = open(title_num_py, "wt")
	tn_file.write("title_num = dict()\n")
	nhym = max_nhymal + 1
	
	for i in range(1, nhym):
		sql = 'SELECT title from {} where id = {}'.format(table_name, i)
		cur = cursor.execute(sql)
		tn_file.write("title_num[\"{}\"] = {}\n".format(cur.fetchone()[0], i))
	tn_file.close()

# usage	
#create_title_to_num("새찬송가.hdb")

# <b>1</b> number of lyrics
# <br> new line

_find_nlyric = re.compile("\d")

#remove_empty_string = lambda del m[i] for i in range(len(m)) if m[i] == ''

def parse_hymal(hymal_str):
	h1 = hymal_str.replace('<b>', '')
	h2 = h1.replace('</b>', '')
	h3 = h2.replace('<br>', '\n')
	h4 = h3.replace('\n\n', '\n')
	h5 = re.split('\d.', h4)
	if h5[0] == '': del h5[0]
	
	hymal_list = []
	nl = len(h5)
	for i in range(nl):
		v = h5[i].split('\n')
		del v[-1]
		hymal_list.append(v)
		
	return hymal_list
	
def get_hymal_by_chapter(num):
	
	conn = db.connect(db_file)
	cursor = conn.cursor()
	
	sql = 'SELECT htext from {} where id = {}'.format(table_name, num)
	cur = cursor.execute(sql)
	return cur.fetchone()[0]

def get_hymal_by_title(title):
	
	conn = db.connect(db_file)
	cursor = conn.cursor()
	sql = 'SELECT htext from {} where title = {}'.format(table_name, title)
	cur = cursor.execute(sql)
	return cur.fetchone()[0]
	
def get_hymal_by_keyword(keyword):
	return
	




