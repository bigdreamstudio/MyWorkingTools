#python
#-*- coding:utf-8 -
"""将文章的标题、评价、链接，分享到SharePoint站点的知识分享中去
@version: v0.1
@author: 周光甫
@license:
@contact:zhougf930@gmail.com
@see:
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import unittest, time, re, os, sys, string

class SToCW(unittest.TestCase):
		
	def setUp(self):
		self.driver = webdriver.Firefox()
		self.driver.implicitly_wait(30)
		self.base_url = "http://192.168.1.80/TM03/PM05/Lists/List/view2.aspx"
		self.verificationErrors = []
		self.accept_next_alert = True
    
	def test_SToCW(self):
		filepath = u'D:\\Python26\\bigdreamstudio\\ShareToCoprizeWeb\\S2CW.csv'
		
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
		
		list = list2array
		list_rows = len(list)
		
		driver = self.driver
		driver.get(self.base_url)
		
		for row in xrange(list_rows):
			aTitle = ''.join(list[row][0]).decode("utf-8")
			aRatingLink = "<p><a target=\"_blank\" href=\"" + ''.join(list[row][2]).decode("utf-8") + "/\">" + ''.join(list[row][0]).decode("utf-8") + "</a></p><p>" + ''.join(list[row][1]).decode("utf-8") + "</p>"
			
			driver.find_element_by_id("zz11_NewMenu").click()
			driver.find_element_by_id("ctl00_m_g_a015e06e_f6a0_448a_ac1c_2c30ed84f0dd_ctl00_ctl04_ctl00_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(aTitle)
			driver.find_element_by_id("ctl00_m_g_a015e06e_f6a0_448a_ac1c_2c30ed84f0dd_ctl00_ctl04_ctl01_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").clear()
			driver.find_element_by_id("ctl00_m_g_a015e06e_f6a0_448a_ac1c_2c30ed84f0dd_ctl00_ctl04_ctl01_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(aRatingLink)
			# ERROR: Caught exception [ReferenceError: selectLocator is not defined]
			driver.find_element_by_id("ctl00_m_g_a015e06e_f6a0_448a_ac1c_2c30ed84f0dd_ctl00_toolBarTbl_RightRptControls_ctl00_ctl00_diidIOSaveItem").click()
			
		pString = "恭喜你！共发布成功" + str(row+1) + "篇文章！"
		print pString.decode("utf-8")
    
	def is_element_present(self, how, what):
		try: self.driver.find_element(by=how, value=what)
		except NoSuchElementException, e: return False
		return True
    
	def is_alert_present(self):
		try: self.driver.switch_to_alert()
		except NoAlertPresentException, e: return False
		return True
    
	def close_alert_and_get_its_text(self):
		try:
			alert = self.driver.switch_to_alert()
			alert_text = alert.text
			if self.accept_next_alert:
				alert.accept()
			else:
				alert.dismiss()
			return alert_text
		finally: self.accept_next_alert = True
    
	def tearDown(self):
		self.driver.quit()
		self.assertEqual([], self.verificationErrors)

if __name__ == "__main__":
    unittest.main()