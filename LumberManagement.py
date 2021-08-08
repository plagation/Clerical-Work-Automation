# -*- coding: utf-8 -*-
"""
Created on Sun Aug  8 12:09:52 2021

@author: kyle.conrad
"""
from Helpers import Helpers

from datetime import datetime
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

class LumberManagement:
    def __init__(self, driver, customer, loadType, dor, bolNum, poNum, units, truckingCompany, driverName, timeGate):
        self.driver = driver
        self.customer = customer
        self.loadType = loadType
        self.dor = dor
        self.bolNum = bolNum
        self.poNum = poNum
        self.units = units
        self.truckingCompany = truckingCompany
        self.driverName = driverName
        self.timeGate = timeGate
        self.timeCheck = datetime.now().strftime("%H:%M")
        
    def navigate_to_LumberManagement(self):
        try:
            self.driver.switch_to_window('portal')
            self.driver.get("http://lms.nascochi.us/Truck/NewLumber")
            WebDriverWait(self.driver, 30).until(lambda x: x.find_element(By.XPATH, '//*[@id="ShipCode"]'))
            
        except KeyboardInterrupt:
            Helpers.Exceptions.user_interrupt(self.driver)
            
        except:
            Helpers.Exceptions.unexpected_exception("navigate to Lumber Management")
            
    def upload(self):
        try:
            print("\nAttempting to upload " + self.dor + " to Lumber Management...")
            shipCode = self.driver.find_element_by_id("Code").get_attribute("value")
            self.driver.find_element_by_id("ShipCode").send_keys(shipCode)
            Select(self.driver.find_element_by_id("CustomerName")).select_by_value(self.customer)
            Select(self.driver.find_element_by_id("LoadType")).select_by_value(self.loadType)
            self.driver.find_element_by_id("OrderNumber").send_keys(self.dor)
            self.driver.find_element_by_id("BolNumber").send_keys(self.bolNum)
            self.driver.find_element_by_id("CustomerPo").send_keys(self.poNum)
            self.driver.find_element_by_id("Units").send_keys(self.units)
            self.driver.find_element_by_id("Carrier").send_keys(self.truckingCompany)
            self.driver.find_element_by_id("TruckDriver").send_keys(self.driverName)
            self.driver.find_element_by_id("InGate").send_keys(self.timeGate)
            self.driver.find_element_by_id("CheckIn").send_keys(self.timeCheck)
            self.driver.find_element_by_xpath('/html/body/div[2]/form/button[2]').click()
            print("Upload successful")
            
        except KeyboardInterrupt:
            Helpers.Exceptions.user_interrupt(self.driver)
            
        except: 
            Helpers.Exceptions.unexpected_exception("upload " + self.dor)
            