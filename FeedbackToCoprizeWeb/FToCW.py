#python
#-*- coding:utf-8 -
"""将反馈问题发送到SharePoint站点上对应的项目中去
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

class FToCW(unittest.TestCase):

	def setUp(self):
		global projectId
		projectName = u"会员自助"
		if projectName == u"岗位职能":
			url4feedback = u"http://192.168.1.80/TM03/PM03/Lists/List1/view.aspx"
			projectId = u"ctl00_m_g_9345aa16_429a_41c7_97e9_6e7c4fc9e81a"
		else:
			if projectName == u"经营管理":
				url4feedback = u"http://192.168.1.80/TM03/PM01/Lists/List1/view.aspx"
				projectId = u"ctl00_m_g_46f0dace_940f_4c8f_a9bf_badda4266792"
			else:
				if projectName == u"团队建设":
					url4feedback = u"http://192.168.1.80/TM03/PM02/Lists/List2/view.aspx"
					projectId = u"ctl00_m_g_3ec9f653_2d5b_4fbd_a861_dc503321dc87"
				else:
					if projectName == u"技术研发":
						url4feedback = u"http://192.168.1.80/TM03/PM06/Lists/List2/view.aspx"
						projectId = u"ctl00_m_g_c963a282_9381_40f0_8ca7_08492bbb4c05"
					else:
						if projectName == u"企业运营":
							url4feedback = u"http://192.168.1.80/TM03/PM04/Lists/List2/view.aspx"
							projectId = u"ctl00_m_g_0323a78a_87f0_4d76_ab7a_abe7c81fd4ab"
						else:
							if projectName == u"对外服务":
								url4feedback = u"http://192.168.1.80/TM02/PM03/Lists/List2/view.aspx"
								projectId = u"ctl00_m_g_174d41b7_5f88_4bc1_82a4_1d86e90e82e5"
							else:
								if projectName == u"市场销售":
									url4feedback = u"http://192.168.1.80/TM02/PM01/Lists/List2/view.aspx"
									projectId = u"ctl00_m_g_5430aefd_f783_42cb_bc8f_b93f718c82c2"
								else:
									if projectName == u"运维支持":
										url4feedback = u"http://192.168.1.80/TM02/PM02/Lists/List5/view.aspx"
										projectId = u"ctl00_m_g_a6e2c10d_10ea_485a_9303_16178186d29f"
									else:
										if projectName == u"产品推广":
											url4feedback = u"http://192.168.1.80/TM02/PM04/Lists/List/view.aspx"
											projectId = u"ctl00_m_g_27bf19b2_a359_498f_a32a_4912239fc323"
										else:
											if projectName == u"门户自助":
												url4feedback = u"http://192.168.1.80/TM01/PM02/Lists/List/view.aspx"
												projectId = u"ctl00_m_g_4ca8e542_c610_4b10_8919_f301010549c6"
											else:
												if projectName == u"会员管理":
													url4feedback = u"http://192.168.1.80/TM01/PM03/Lists/List4/view.aspx"
													projectId = u"ctl00_m_g_6993004e_2174_4bf7_9a74_95c3bce5a28f"
												else:
													if projectName == u"仓储服务":
														url4feedback = u"http://192.168.1.80/TM04/PM01/Lists/List2/view.aspx"
														projectId = u"ctl00_m_g_5cb9bdff_81d8_4ebd_a722_36e95f1786c6"
													else:
														if projectName == u"库存管理":
															url4feedback = u"http://192.168.1.80/TM04/PM02/Lists/List2/view.aspx"
															projectId = u"ctl00_m_g_2ac0b6ec_c72e_4050_81bc_fdbf668a831f"
														else:
															if projectName == u"数据交换":
																url4feedback = u"http://192.168.1.80/TM04/PM03/Lists/List2/view.aspx"
																projectId = u"ctl00_m_g_f130c109_2fc5_448d_a738_15cd9c70b857"
															else:
																if projectName == u"会员自助":
																	url4feedback = u"http://192.168.1.80/TM04/PM04/Lists/List/view.aspx"
																	projectId = u"ctl00_m_g_729705ff_d691_45c6_be45_803394049edf"
		self.driver = webdriver.Firefox()
		self.driver.implicitly_wait(30)
		self.base_url = url4feedback
		self.verificationErrors = []
		self.accept_next_alert = True

	def test_FToCW(self):
		filepath = u'D:\\Python26\\bigdreamstudio\\FeedbackToCoprizeWeb\\F2CW.csv'
		
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