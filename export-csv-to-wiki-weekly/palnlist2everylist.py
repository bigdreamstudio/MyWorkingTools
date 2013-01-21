#python
#-*- coding:utf-8 -
"""将Redmine导出的CSV文件，按执行者生成Redmine的wiki文件
@version: v0.1
@author: 周光甫
@license:
@contact:zhougf.zhupp@gmail.com
@see:
"""

import os,sys,re,string


#读取文件内容，生成二维数组
def file22array(filepath):
	f1 = open(filepath,'rb')
	t = f1.readlines()
	total_rows = len(t)
	f1.close()
	
	f2 = open(filepath,'rb')	
	i = 1
	list2array=[]	
	while i <= total_rows:
		s = f2.readline()
		slist = s.split(',') 
		#字符串的split函数默认分隔符是空格 ' ',所以我用','
		list2array.append(slist)
		i = i+1
	f2.close()
	return list2array
	
"""
#选择需要的列生成新的二维数组
def select_right_items(trans_array):
	sr_list2array = zip(*trans_array)
	sr_rows = len(sr_list2array)
	items_right_rows = []
	sr_result_list = []
	right_items = [u'#' ,u'主题' ,u'指派给' ,u'计划完成日期']
	r_items = zip(right_items)
	it_rows = len(right_items)	
	
	for i_row in xrange(it_rows):
		for row in xrange(sr_rows):
			if trans_array[row][0] == r_items[i_row][0]:
				items_right_rows.append(trans_array[row])

	sr_result_list = zip(*items_right_rows)
	return sr_result_list
"""

#选择二维数组中，含有值为columns_name的每一行，生成到新的二维数组
def select_right_columns(as_list2array , columns_name):
	s_list2array = as_list2array
	s_rows = len(s_list2array)
	#s_columns = len(s_list2array[0])
	all_right_rows = []
	s_result_list = []
	#先记录所有符合的行号，最后再统一处理，否则出现list index out of range错误
	"""
	i = 0
	while i < s_rows:
		if as_list2array[i][2] == columns_name.encode('utf-8'):
			all_right_rows.append(i)
		i = i+1
	"""
	for j in xrange(s_rows):
		if as_list2array[j][2] == columns_name.encode('utf-8'):
			all_right_rows.append(j)
	
	for row in xrange(len(all_right_rows)):
		r = all_right_rows[row]
		s_result_list.append(as_list2array[r])
	result_list = s_result_list
	return result_list

#将按特定要求生成的数组，转成wiki打印格式，存入新的数组中
def add_style(list2array):
	a_list2array = list2array
	a_rows = len(list2array)
	for row in xrange(a_rows):
		a_list2array[row][0] = '|#' + a_list2array[row][0]
		a_list2array[row][1] = '|' + a_list2array[row][1]
		a_list2array[row][2] = '|'
		strinfo = re.compile('\r')
		a_list2array[row][3] = strinfo.sub('|',a_list2array[row][3])
	
	i = 0
	m = '|Why|\\2=.周计划任务，完成项目阶段工作|\n|When|\\2=.任务执行日期：'
	n = '|How much|\\2=.预计1小时完成任务|\n|How|\\2=.略|\n'
	e = ['|_.任务编号', '|_.主题(What)', '|', '_.计划完成日期|\n']
	m_list2array = []
	m_list2array.append(e)
	while i < a_rows:
		final_date = a_list2array[i][3]
		add_des = [m + final_date + n , '', '', '']
		m_list2array.append(a_list2array[i])
		m_list2array.append(add_des)
		i = i+1
	
	"""
	m_list2array = []
	m = '|Why|\2=.项目进度需要，周计划任务|\n|When|\2=.'
	n = '\n|How|\2=.|\n|How much|\2=.|'
	for row in xrange(a_rows):
		final_date = a_list2array[row][3]
		#add_des = m + final_date + n
		add_des = [[m + final_date + n , '' , '' , '']]
		m_list2array.append(a_list2array[row]) 
		m_list2array.append(add_des)
	"""
	as_list2array = m_list2array
	return as_list2array
	
#将二维数组写到特定的文件里
def write2files(as_list2array , file_path):
	f = open(file_path,'wb')
	#f.write(as_list2array.encode('utf'))
	f.writelines(''.join(str(v) for v in row) for row in as_list2array)
	#f.writelines(''.join(c for c in row) for row in as_list2array)
	#f.writelines('|'.join(c for c in row) for row in as_list2array)
	f.close()

#整合所有的方法，完成csv文件内容转成wiki格式内容，分别生成文本文件
def palns2everywiki():

	#thefilepath = u'D:\\Python26\\bigdreamstudio\\export.csv'
	open_path = u'D:\\Python26\\bigdreamstudio\\wsp.csv'
	#write2files(select_right_items(file22array(thefilepath)) , open_path)

	read_result = file22array(open_path)
	all_name = [u'韩 雨',u'吴 勇庆',u'周 光甫',u'钱 文豪',u'惠 卿',u'吴 章强',u'薛 富玮',u'张 雪培',u'汪 唐明',u'及 松浩',u'杨 健',u'张 伟豪']
	all_name_rows = len(all_name)
	for row in xrange(all_name_rows):
		save_path = u'D:\\Python26\\bigdreamstudio\\'+all_name[row]+'.txt'
		select_rights = select_right_columns(read_result , all_name[row])
		array_len = len(select_rights)
		if  array_len >= 1 :
			write2files(add_style(select_rights) , save_path)
	
#程序执行的入口	
if __name__=="__main__":
	palns2everywiki()