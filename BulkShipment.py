# -*- coding: utf-8 -*-
"""
Created on Fri Jul 23 13:28:22 2021

@author: kyle.conrad
"""

import Shipment

import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

class BulkShipment(Shipment.Shipment):
    def __init__(self, driver, driverName, truckingCompany, dor, licenseNumber, clientCode, tare):
        self.driver = driver
        self.driverName = driverName
        self.truckingCompany = truckingCompany
        self.dor = dor
        self.licenseNumber= licenseNumber
        self.clientCode = clientCode
        self.tare = tare
        
    def set_lot(self):
        self.lot = ""
        
    def set_analysis(self):
        self.analysis = ""
        
    def set_poNumber(self):
        self.poNumber = ""
        
    def set_remarks(self):
        self.set_clerk()
        self.remarks = "Tarp load\nDrivers accpet paperwork as is, signing waived due to COVID-19 restrictions\n{}".format(self.clerk)
        
    def shipped_cargo_fill(self):
        pass
        
    def set_variables(self):
        self.set_poNumber()
        self.set_lot()
        self.set_analysis()
        self.set_remarks()
        
    def CreateShipment(self):
        self.set_variables()
        
        #navigate to tc3 shipments
        self.driver.get('https://tos.qsl.com/client-inventories/shipment-of-materials')
        element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/section[1]/div[2]/div[2]/button'))
        time.sleep(1.5)
        
        #create new shipment (select port) of bulk material via truck
        element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/section[1]/div[2]/div[2]/button'))
        time.sleep(1.5)
        element.click()
        element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '/html/body/div[3]/div/form/div/div[3]/label[1]'))
        element.click()
        self.driver.find_element_by_xpath('/html/body/div[3]/div/form/header/menu/button[2]').click()
        
        #fill in truck information
        element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="react-select-2-input"]'))
        element.send_keys(self.truckingCompany + Keys().RETURN)
        self.driver.find_element_by_xpath('//*[@id="driverName"]').send_keys(self.driverName)
        self.driver.find_element_by_xpath('//*[@id="carrierBill"]').send_keys(self.poNumber)
        self.driver.find_element_by_xpath('//*[@id="transportationNumber"]').send_keys(self.licenseNumber)
        self.driver.find_element_by_xpath('//*[@id="specialInstructions"]').send_keys(self.remarks)
        
        #add destination
        self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/fieldset[2]/article/header/button').click()
        element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="react-select-3-input"]'))
        element.send_keys(self.receiver + Keys().RETURN)
        self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/fieldset[2]/article/section/section/div/menu/section/div[5]/i').click()
        element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="formattedAddress"]'))
        element.send_keys(self.address)
        self.driver.find_element_by_xpath('/html/body/div[5]/div/header/menu/button[2]').click()
        
        #save shipment
        self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/section[2]/div/div[2]/button[2]').click()
        
        #navigate to shipped items
        element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/form/section[1]/div[3]/nav/menu/a[2]'))
        element.click()
        
        #enter base information for shipped items
        self.shipped_cargo_fill()