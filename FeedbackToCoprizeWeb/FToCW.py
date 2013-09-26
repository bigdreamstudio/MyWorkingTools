#python
#-*- coding:utf-8 -
"""将反馈问题发送到SharePoint站点上对应的项目中去
@version: v0.2
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

class FToCW(unittest.TestCase):
	global projectId
	global listArray
	
	def readFile(filepath):	
		f1 = open(filepath,'rb')
		t = f1.readlines()
		total_rows = len(t)
		f1.close()
		
		f2 = open(filepath,'rb')
		i = 1
		list2array = []
		while i <= total_rows:
			s = f2.readline()
			slist = s.split(',') 
			#字符串的split函数默认分隔符是空格 ' ',所以我用','
			list2array.append(slist)
			i = i+1
		f2.close()
		return list2array

	def setUp(self):
		projectName = u"会员自助"
		filepath = u'D:\\Python26\\bigdreamstudio\\FeedbackToCoprizeWeb\\myconfig.pycfg'
		listArray = readFile(filepath)
		listArray_rows = len(listArray)
		for drow in xrange(listArray_rows):
			if projectName == ''.join(listArray[drow][0].decode("utf-8")):
				url4feedback = ''.join(listArray[drow][1]).decode("utf-8")
				projectId = ''.join(listArray[drow][2]).decode("utf-8")
				drow = listArray_rows				
		self.driver = webdriver.Firefox()
		self.driver.implicitly_wait(30)
		self.base_url = url4feedback
		self.verificationErrors = []
		self.accept_next_alert = True

	def test_FToCW(self):
		filepath = u'D:\\Python26\\bigdreamstudio\\FeedbackToCoprizeWeb\\F2CW.csv'
		list = readFile(filepath)
		list_rows = len(list)
		driver = self.driver
		driver.get(self.base_url)		
		for row in xrange(list_rows):
			fTitle = ''.join(list[row][0]).decode("utf-8")
			fLiability = ''.join(list[row][1]).decode("utf-8")
			fAccessPath = ''.join(list[row][2]).decode("utf-8")
			fDescription = ''.join(list[row][3]).decode("utf-8")
			fAttachingTask = ''.join(list[row][4]).decode("utf-8")
			fComment = ''.join(list[row][5]).decode("utf-8")			
			driver.find_element_by_id("zz11_NewMenu").click()
			driver.find_element_by_id(projectId + "_ctl00_ctl04_ctl00_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(fTitle)
			driver.find_element_by_id(projectId + "_ctl00_ctl04_ctl01_ctl00_ctl00_ctl04_ctl00_ctl00_UserField_downlevelTextBox").clear()
			driver.find_element_by_id(projectId + "_ctl00_ctl04_ctl01_ctl00_ctl00_ctl04_ctl00_ctl00_UserField_downlevelTextBox").send_keys(fLiability)
			# ERROR: Caught exception [ReferenceError: selectLocator is not defined]
			# ERROR: Caught exception [ReferenceError: selectLocator is not defined]
			driver.find_element_by_id(projectId + "_ctl00_ctl04_ctl04_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(fAccessPath)
			driver.find_element_by_id(projectId + "_ctl00_ctl04_ctl05_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").clear()
			driver.find_element_by_id(projectId + "_ctl00_ctl04_ctl05_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(fDescription)
			driver.find_element_by_id(projectId + "_ctl00_ctl04_ctl06_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(fAttachingTask)
			# ERROR: Caught exception [ReferenceError: selectLocator is not defined]
			driver.find_element_by_id(projectId + "_ctl00_ctl04_ctl10_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").clear()
			driver.find_element_by_id(projectId + "_ctl00_ctl04_ctl10_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(fComment)
			# ERROR: Caught exception [ReferenceError: selectLocator is not defined]
			driver.find_element_by_id(projectId + "_ctl00_toolBarTbl_RightRptControls_ctl00_ctl00_diidIOSaveItem").click()			
		pString = "恭喜你！共发布成功" + str(row+1) + "个反馈！"
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