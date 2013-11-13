#python
#-*- coding:utf-8 -
"""
@description: 将每周团队协作工作的结果，以表扬和批评的方式发到到SharePoint站点相应的位置中，并且发表扬后提示出来。使用到的数据文件PAT-FT.csv中的每列内容，从左至右依次表示：发的是表扬或批评、表扬或批评的任务名称、责任人姓名
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

class PraiseAndTips(unittest.TestCase):
	def setUp(self):
		self.driver = webdriver.Firefox()
		self.driver.implicitly_wait(30)
		self.base_url = "http://192.168.1.80/"
		self.verificationErrors = []
		self.accept_next_alert = True

	def test_praise_and_tips(self):
		pWeek = u"201311-2"
		pTheDate = u"2013/11/08"
		#读取文件数据
		filepath = u'D:\\Python26\\bigdreamstudio\\PraiseAndTips\\PAT-FT.csv'
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
		
		list = list2array
		list_rows = len(list)
		p = 0
		t = 0
		f = 0
		driver = self.driver
		for row in xrange(list_rows):
			pomCri = ''.join(list[row][0]).decode("utf-8")
			pTitle = ''.join(list[row][1]).decode("utf-8")
			pLiability = ''.join(list[row][2]).decode("utf-8")

			if pomCri == u"表扬":
				url4Praise = self.base_url + "/TM03/PM02/Lists/List4/view2.aspx"
				url4Tips = self.base_url + "/TM03/PM03/Lists/List/view2.aspx"
				#发表扬
				driver.get(url4Praise)
				driver.find_element_by_id("zz11_NewMenu").click()
				driver.find_element_by_id("ctl00_m_g_a170c922_8d1a_4a87_9494_4644b4889091_ctl00_ctl04_ctl00_ctl00_ctl00_ctl04_ctl00_ctl00_UserField_downlevelTextBox").clear()
				driver.find_element_by_id("ctl00_m_g_a170c922_8d1a_4a87_9494_4644b4889091_ctl00_ctl04_ctl00_ctl00_ctl00_ctl04_ctl00_ctl00_UserField_downlevelTextBox").send_keys(pLiability)
				driver.find_element_by_id("ctl00_m_g_a170c922_8d1a_4a87_9494_4644b4889091_ctl00_ctl04_ctl01_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").clear()
				driver.find_element_by_id("ctl00_m_g_a170c922_8d1a_4a87_9494_4644b4889091_ctl00_ctl04_ctl01_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(u"表扬：在“" + pTitle + u"”协作工作中积极配合、效果突出(" + pLiability + u")")
				driver.find_element_by_id("ctl00_m_g_a170c922_8d1a_4a87_9494_4644b4889091_ctl00_ctl04_ctl02_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").clear()
				driver.find_element_by_id("ctl00_m_g_a170c922_8d1a_4a87_9494_4644b4889091_ctl00_ctl04_ctl02_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(u"为表扬您，在“" + pTitle + u"”协作工作中，积极配合，效果突出，特发表扬以彰成效！")
				# ERROR: Caught exception [ReferenceError: selectLocator is not defined]
				driver.find_element_by_id("ctl00_m_g_a170c922_8d1a_4a87_9494_4644b4889091_ctl00_ctl04_ctl04_ctl00_ctl00_ctl04_ctl00_ctl00_DateTimeField_DateTimeFieldDate").clear()
				driver.find_element_by_id("ctl00_m_g_a170c922_8d1a_4a87_9494_4644b4889091_ctl00_ctl04_ctl04_ctl00_ctl00_ctl04_ctl00_ctl00_DateTimeField_DateTimeFieldDate").send_keys(pTheDate)
				driver.find_element_by_id("ctl00_m_g_a170c922_8d1a_4a87_9494_4644b4889091_ctl00_toolBarTbl_RightRptControls_ctl00_ctl00_diidIOSaveItem").click()
				p = p + 1
				#发提示
				driver.get(url4Tips)
				driver.find_element_by_id("zz11_NewMenu").click()
				driver.find_element_by_id("ctl00_m_g_078b61fe_d5c0_46e3_9af1_a96ba9040999_ctl00_ctl04_ctl00_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").clear()
				driver.find_element_by_id("ctl00_m_g_078b61fe_d5c0_46e3_9af1_a96ba9040999_ctl00_ctl04_ctl00_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(u"已发：表扬-在“" + pTitle + u"”协作工作中积极配合、效果突出(" + pLiability + u")")
				driver.find_element_by_id("ctl00_m_g_078b61fe_d5c0_46e3_9af1_a96ba9040999_ctl00_ctl04_ctl01_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").clear()
				driver.find_element_by_id("ctl00_m_g_078b61fe_d5c0_46e3_9af1_a96ba9040999_ctl00_ctl04_ctl01_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(url4Praise)
				driver.find_element_by_id("ctl00_m_g_078b61fe_d5c0_46e3_9af1_a96ba9040999_ctl00_ctl04_ctl02_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").clear()
				driver.find_element_by_id("ctl00_m_g_078b61fe_d5c0_46e3_9af1_a96ba9040999_ctl00_ctl04_ctl02_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(u"<p>已发表扬，如标题。</p>\n<p>请 " + pLiability + u" 参照http://192.168.1.80/publish/Coprize/01%E5%B2%97%E4%BD%8D%E8%81%8C%E8%83%BD/F.%E4%BB%BB%E5%8A%A1%E8%AE%A8%E8%AE%BA/Coprize/" + pWeek +"/" + u"链接内容，核对有无漏发，有无表扬等级错误的问题。</p>\n<p>谢谢！</p>")
				driver.find_element_by_id("ctl00_m_g_078b61fe_d5c0_46e3_9af1_a96ba9040999_ctl00_ctl04_ctl03_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").clear()
				driver.find_element_by_id("ctl00_m_g_078b61fe_d5c0_46e3_9af1_a96ba9040999_ctl00_ctl04_ctl03_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(u"团队协作." + pWeek + u".Coprize.团队协作交流")
				driver.find_element_by_id("ctl00_m_g_078b61fe_d5c0_46e3_9af1_a96ba9040999_ctl00_toolBarTbl_RightRptControls_ctl00_ctl00_diidIOSaveItem").click()
				t = t + 1
			else:
				if pomCri == u"批评":
					url4feedback = self.base_url + "/TM03/PM03/Lists/List1/view.aspx"
					#发反馈(批评)
					driver.get(url4feedback)
					driver.find_element_by_id("zz11_NewMenu").click()
					driver.find_element_by_id("ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a_ctl00_ctl04_ctl00_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").clear()
					driver.find_element_by_id("ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a_ctl00_ctl04_ctl00_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(u"批评：在“" + pTitle + u"”协作工作中执行不力、消极应付(" + pLiability + u")")
					driver.find_element_by_id("ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a_ctl00_ctl04_ctl01_ctl00_ctl00_ctl04_ctl00_ctl00_UserField_downlevelTextBox").clear()
					driver.find_element_by_id("ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a_ctl00_ctl04_ctl01_ctl00_ctl00_ctl04_ctl00_ctl00_UserField_downlevelTextBox").send_keys(pLiability)
					driver.find_element_by_id("ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a_ctl00_ctl04_ctl02_ctl00_ctl00_ctl04_ctl00_DropDownChoice").click()
					# ERROR: Caught exception [ReferenceError: selectLocator is not defined]
					driver.find_element_by_css_selector(u"option[value=\"已关闭\"]").click()
					driver.find_element_by_id("ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a_ctl00_ctl04_ctl04_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").clear()
					driver.find_element_by_id("ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a_ctl00_ctl04_ctl04_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(u"http://192.168.1.80/publish/Coprize/01%E5%B2%97%E4%BD%8D%E8%81%8C%E8%83%BD/F.%E4%BB%BB%E5%8A%A1%E8%AE%A8%E8%AE%BA/Coprize/" + pWeek +"/")
					driver.find_element_by_id("ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a_ctl00_ctl04_ctl05_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").clear()
					driver.find_element_by_id("ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a_ctl00_ctl04_ctl05_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(u"因你在“" + pTitle + u"”协作工作中，执行不力、消极应付，特发批评以儆效尤！")
					driver.find_element_by_id("ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a_ctl00_ctl04_ctl06_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").clear()
					driver.find_element_by_id("ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a_ctl00_ctl04_ctl06_ctl00_ctl00_ctl04_ctl00_ctl00_TextField").send_keys(u"团队协作." + pWeek + u".Coprize.团队协作交流")
					driver.find_element_by_id("ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a_ctl00_ctl04_ctl11_ctl00_ctl00_ctl04_ctl00_DropDownChoice").click()
					# ERROR: Caught exception [ReferenceError: selectLocator is not defined]
					driver.find_element_by_css_selector(u"option[value=\"内容细致\"]").click()
					driver.find_element_by_id("ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a_ctl00_toolBarTbl_RightRptControls_ctl00_ctl00_diidIOSaveItem").click()
					f = f + 1
		pString = "恭喜你！共发成功" + str(p) + "个表扬，" + str(t) + "个提示，" + str(f) + "个反馈(批评)！"
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