# -*- coding: utf-8 -*-
"""
Created on Fri Jul 23 13:26:19 2021

@author: kyle.conrad
"""
import BulkShipment
from Helpers import Helpers

import time, traceback

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

class ErametShipment(BulkShipment.BulkShipment):
    
    def split_file(self):
        self.splitFile = self.convertedFile.split("\n")
        
    def set_lot(self):
        try:
            print("Attempting to receive lot information...")
            for i in range(0, len(self.splitFile)):
                if "LOT:" in self.splitFile[i].upper():
                    self.lot = self.splitFile[i].replace("Lot: ", "")
                    break
            print("Retreived lot  " + self.lot)
        except:
            Helpers.Exceptions.unexpected_exception("retreive lot number")
        
    def set_poNumber(self):
        try:
            print("Attempting to retreive PO number...")
            for i in range(0, len(self.splitFile)):
                if "BUYER'S" in self.splitFile[i].upper():
                    self.poNumber = self.splitFile[i + 2]
                    break
            print("Retrieved PO Number "+ self.poNumber)
        except:
            Helpers.Exceptions.unexpected_exception("reteive PO Number")
            
    def set_productSubtype(self):
        
        self.set_lot()
        try:
            print("Attempting to set product subtype...")
            if "KV" in self.lot:
                if "Low Carbon" in self.convertedFile:
                    self.productSubtype = "LCSiMn 2x1/2"
                else:
                    self.productSubtype = "SiMn (Silicomanganese)"
            else: 
                for i in range(0, len(self.splitFile)):
                    if "SIZE" in self.splitFile[i].upper():
                        self.productSubtype = "LCFeMn " + self.splitFile[i + 4].replace(" ", "").lower()
                        break
            print("Set product subtype as " + self.productSubtype)
        except:
            Helpers.Exceptions.unexpected_exception("set product subtype")
            
    def set_address(self):
        self.address = ""
        
        try:
            print("Attempting to set receiver address...")
            for i in range(0,len(self.splitFile)):
                if self.splitFile[i].upper() in ("UNITED STATES", "CANADA"):
                    addressEndIndex = i
                    break
                
            i = addressEndIndex
            while self.splitFile[i] != "":
                i -= 1
                
            self.addressStartIndex = i + 2
            
            for i in range(self.addressStartIndex, addressEndIndex+1):
                if i != addressEndIndex:
                    self.address += self.splitFile[i] + "\n"
                else:
                    self.address += self.splitFile[i]
            print("Retrieved receiver address")
            
        except:
            Helpers.Exceptions.unexpected_exception("retreive receiver address")
        
        
    def set_receiverRemarks(self):
        self.receiverRemarks = ""
        
        try:
            print("Attempting to set receiver remarks...")
            for i in range(len(self.splitFile)):
                
                if "Shipping Instructions" in self.splitFile[i]:
                    startInd = i + 2
                    i = i + 2
                    
                elif "Plant Instructions" in self.splitFile[i]:
                    endInd = i-1
                    break
            
            for i in range(startInd, endInd):
                if self.splitFile[i] != "":
                    self.receiverRemarks += self.splitFile[i] + "\n"
                
            print("Successfully set receiver remarks")
        except:
            Helpers.Exceptions.unexpected_exception("set receiver remarks")
        
    def set_receiver(self):
        self.set_address()
        
        try:
            print("Attempting to set receiver...")
            self.receiver = self.splitFile[self.addressStartIndex-1]
            print("Receiver set as " + self.receiver)
        except:
            Helpers.Exceptions.unexpected_exception("set receiver")
            
        
    def shipped_cargo_fill(self):
        print("Adding items to shipment...")
        
        actions = ActionChains(self.driver)
        try:
            element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="react-select-17-input"]'))
            element.send_keys("Eramet" + Keys().RETURN)
            element = self.driver.find_element_by_xpath('//*[@id="react-select-18-input"]')
            actions.send_keys_to_element(element, "Manganese Alloy").pause(.1).send_keys_to_element(element, Keys().RETURN).perform()
            time.sleep(.2)
            element = self.driver.find_element_by_xpath('//*[@id="react-select-19-input"]')
            actions.send_keys_to_element(element, self.productSubtype).pause(.1).send_keys_to_element(element, Keys().RETURN).perform()
            time.sleep(.2)
            self.driver.find_element_by_xpath('//*[@id="react-select-20-input"]').send_keys("lb" + Keys().RETURN)
            time.sleep(.2)
            self.driver.find_element_by_xpath('//*[@id="react-select-21-input"]').send_keys("Scale Reading" + Keys().RETURN)
            time.sleep(.2)
            self.driver.find_element_by_xpath('//*[@id="tareWeight"]').send_keys(self.tare)
            time.sleep(.2)
            self.driver.find_element_by_xpath('//*[@id="grossWeight"]').send_keys(str(int(self.tare) + 1))
            time.sleep(.2)
            for i in range(1, 100):
                testLot = self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/section[3]/div/section/article[' + str(i) + ']/section/div[3]/output').text
                if testLot == self.lot:
                    self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/section[3]/div/section/article[' + str(i) + ']/section/div[6]/div/input').send_keys("1")
                    break
            self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/section[1]/div[2]/div[2]/button[5]').click()
            print("Items successfully added to shipment")
        
        except KeyboardInterrupt:
            Helpers.Exceptions.user_interrupt(self.driver)
        
        except:
            Helpers.Exceptions.unexpected_exception("add items to shipment")
            
        
        
    def set_remarks(self):
        self.set_clerk()
        self.split_file()
        self.set_poNumber()
        self.set_receiverRemarks()
        
        try:
            print("Setting remarks...")
            self.remarks = "Order#: {}\n{}Tarp Load\nDrivers accept paperwork as is, signing waived due to COVID-19 restrictions\n{}".format(self.dor, self.receiverRemarks, self.clerk)
            print("Remarks successfully set")
        
        except KeyboardInterrupt:
            Helpers.Exceptions.user_interrupt(self.driver)
        
        except:
            Helpers.Exceptions.unexpected_exception("set remarks")
            
    def set_variables(self):
        self.convert_to_text("Z:/SCALE OFFICE/Bulk Operations/Eramet/Current Releases/" + self.dor + "_W.PDF")
        self.set_remarks()
        self.set_productSubtype()
        self.set_receiver()
        self.set_pickupNumber()
        self.set_carrierBill()
        
        
"""
test = ErametShipment("driver", "driverName", "truckingCompany", "MKT-4-99688-1-4", "licenseNumber", "E", 5000)

test.convert_to_text("Z:/SCALE OFFICE/Bulk Operations/Eramet/Current Releases/" + test.dor + "_W.PDF")
test.split_file()
test.set_lot()
test.set_reeiverRemarks()
test.set_productSubtype()
f = open("testfile.txt", 'w')
f.write(test.convertedFile)
f.close()
"""
