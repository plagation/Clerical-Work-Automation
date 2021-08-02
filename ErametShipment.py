# -*- coding: utf-8 -*-
"""
Created on Fri Jul 23 13:26:19 2021

@author: kyle.conrad
"""
import BulkShipment

import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains

class ErametShipment(BulkShipment.BulkShipment):
    def split_file(self):
        self.splitFile = self.convertedFile.split("\n")
        
    def set_lot(self):
        for i in range(0, len(self.splitFile)):
            if "LOT:" in self.splitFile[i].upper():
                self.lot = self.splitFile[i].replace("Lot: ", "")
                break
        
    def set_analysis(self):
        self.set_lot()
        self.analysis = ""
        self.driver.get('https://tos.qsl.com/client-inventories/inventories/378/piles?isArchived=false&includeEmptyItems=true')
        WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/section[3]/div[3]/div[1]/div[4]/div/div[1]/div[4]/span'))
        
        isValidEntry = True
        i = 4
        
        while isValidEntry:
            try:        
                element = self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/section[3]/div[3]/div[1]/div[' + str(i) + ']/div/div[1]/div[4]/span')
                if element.text == self.lot:
                    element.click()
                    element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="remark"]'))
                    if element.text != "":
                        self.analysis = element.text
                        break
                i += 1
            except NoSuchElementException:
                self.analysis = "No Analysis Found for Given Lot!"
        
    def set_poNumber(self):
        for i in range(0, len(self.splitFile)):
            if "BUYER'S" in self.splitFile[i].upper():
                self.poNumber = self.splitFile[i + 2]
                break
        
    def set_productSubtype(self):
        if "KV" in self.lot:
            self.productSubtype = "LCSiMn 2x1/2"
        else: 
            for i in range(0, len(self.splitFile)):
                if "SIZE" in self.splitFile[i].upper():
                    self.productSubtype = "LCFeMn " + self.splitFile[i + 4].replace(" ", "").lower()
                    break
            
    def set_address(self):
        self.address = ""
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
        
        
    def set_receiver(self):
        self.set_address()
        self.receiver = self.splitFile[self.addressStartIndex-1]
        
    def shipped_cargo_fill(self):
        actions = ActionChains(self.driver)
        
        element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="react-select-16-input"]'))
        element.send_keys("Eramet" + Keys().RETURN)
        element = self.driver.find_element_by_xpath('//*[@id="react-select-17-input"]')
        actions.send_keys_to_element(element, "Manganese Alloy").pause(.1).send_keys_to_element(element, Keys().RETURN).perform()
        time.sleep(.2)
        element = self.driver.find_element_by_xpath('//*[@id="react-select-18-input"]')
        actions.send_keys_to_element(element, self.productSubtype).pause(.1).send_keys_to_element(element, Keys().RETURN).perform()
        time.sleep(.2)
        self.driver.find_element_by_xpath('//*[@id="react-select-19-input"]').send_keys("lb" + Keys().RETURN)
        time.sleep(.2)
        self.driver.find_element_by_xpath('//*[@id="react-select-20-input"]').send_keys("Scale Reading" + Keys().RETURN)
        time.sleep(.2)
        self.driver.find_element_by_xpath('//*[@id="tareWeight"]').send_keys(self.tare)
        time.sleep(.2)
        self.driver.find_element_by_xpath('//*[@id="grossWeight"]').send_keys(str(int(self.tare) + 1))
        time.sleep(.2)
        self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/section[1]/div[2]/div[2]/button[5]').click()
        
    def set_remarks(self):
        self.set_clerk()
        self.split_file()
        self.set_poNumber()
        self.set_analysis()
        self.remarks = "Order#: {}\nPO#: {}\nLot#: {}\nAnalysis: {}\nTarp Load\nDrivers accept paperwork as is, signing waived due to COVID-19 restrictions\n{}".format(self.dor, self.poNumber, self.lot, self.analysis, self.clerk)
        
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
test.set_productSubtype()
f = open("testfile.txt", 'w')
f.write(test.convertedFile)
f.close()


print(test.productSubtype)
"""