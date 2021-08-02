# -*- coding: utf-8 -*-
"""
Created on Thu Jul 22 11:58:09 2021

@author: kyle.conrad
"""

import Shipment

import re
import time
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

class RusalShipment(Shipment.Shipment):
    def set_product(self):
        if "INGOTS" in self.convertedFile:
            self.product = "INGOTS"
        else:
            self.product = "BILLETS"
        
    def set_poNumber(self):
        fileLines = self.convertedFile.split("\n")
        trucksIndex = fileLines.index("TRUCKS")
        i = 3
        while i < 7:
            if fileLines[trucksIndex + i] == "":
                break;
            i += 1
        if i == 7:
            self.poNumber = "COULD NOT FIND PO_NUMBER"
        numLines = i - 2
        self.poNumber = fileLines[trucksIndex + numLines + 3].strip()
        
    
    def set_remarks(self):
        self.set_clerk()
        self.set_poNumber()
        self.remarks = "{}\nPO#: {}\nMaterial must be free of dirt debris.\nMetal straps must be replaced with plastic straps.\nDrivers accept paperwork as is, signing waived due to COVID-19 restrictions.\n{}".format(self.dor, self.poNumber, self.clerk)
        
    def set_pickupNumber(self):
        self.pickupNumber = self.dor
        
    #incomplete needs methods for dealing with multiple piece counts and for dealing with billets
    def set_shippedBundles(self):
        if self.product == "INGOTS":
            self.shippedBundles = re.findall(r'\d{1,2} BUNDLES', self.convertedFile)[0].replace(" BUNDLES", "")
        else: 
            self.convertedFile = re.sub(r'TKS','TRUCKS', self.convertedFile)
            self.shippedBundles = re.findall(r'\d{1,2}\sTRUCKS\s@\s\d{1,2}\sPCS', self.convertedFile)
            
    def set_variables(self):
        self.convert_to_text("Z:/SCALE OFFICE/RUSAL/Current Releases/" + self.dor + ".pdf")
        self.set_product()
        self.set_pickupNumber()
        self.set_poNumber()
        self.set_remarks()
        self.set_carrierBill()
        self.set_address()
        self.set_receiver()
        self.set_shippedBundles()
    
    #incomplete needs methods for dealing with multiple B/Ls on site and for dealing with multiple piece counts
    def add_load(self):
        WebDriverWait(self.driver,50).until(lambda x: x.find_element(By.XPATH,'//*[@id="viewport"]/article/section/form/section[1]/div[3]/nav/menu/a[2]'))
        self.bolNum = self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/section[1]/div[2]/div[1]/h3').text
        self.driver.get('https://tos.qsl.com/client-inventories/shipment-of-materials/' + self.bolNum + '/loading-request')
        WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/article/header[2]/aside/button'))
        self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/article/header[2]/aside/button').click()
        WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '/html/body/div[13]/div/div/article/form/div/div/button'))
        self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/form/div/div/button').click()
        WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '/html/body/div[13]/div/div/article/section/div/div/div[1]/div[4]/div/div[2]/span/button'))
        time.sleep(2.5)
        self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section/div/div/div[1]/div[4]/div/div[2]/span/button').click()
        WebDriverWait(self.driver, 10).until(lambda x: x.find_element(By.XPATH, '//*[@id="quantity"]'))
        self.driver.find_element_by_xpath('//*[@id="quantity"]').send_keys(self.shippedBundles  + Keys().RETURN)
        time.sleep(.5)
        self.driver.find_element_by_xpath('/html/body/div[13]/div/div/div[1]/header/menu/button[3]').click()
        time.sleep(2)
        self.driver.find_element_by_xpath('/html/body/div[13]/div/div/div[1]/header/menu/button[2]').click()
        

test = RusalShipment("driver", "driverName", "truckingCompany", "RAC-015171", "licenseNumber", "clientCode")

#Specific to network (Modification for scale necessary) // testing only 
#test.convert_to_text("C:/Users/kyleb/Documents/GitHub/Clerical Apps/References/" + test.dor + "-VTE-B80-000.pdf")

test.set_variables()
f = open("testfile.txt", 'w')
f.write(test.convertedFile)
f.close()

print(test.shippedBundles)
