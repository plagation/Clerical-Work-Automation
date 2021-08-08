# -*- coding: utf-8 -*-
"""
Created on Thu Jul 22 08:59:01 2021

@author: kyle.conrad
"""
from Helpers import Helpers

codeMap = {"BOS":5, "AL":3, "C":6, "E":16, "PI":45, "LBF IN":24, "LBF OUT":25, "P":29, "RUS":56, "CMS":11}

class TruckManagement:
    def __init__(self, driver, timeIn, dor, gatePass, licenseNumber, driverName, truckingCompany, clientCode):
        self.timeIn = timeIn
        self.dor = dor
        self.gatePass = gatePass
        self.licenseNumber = licenseNumber
        self.driverName = driverName
        self.truckingCompany = truckingCompany
        self.driver = driver
        self.clientCode = clientCode
        
    def Upload(self):
        try:
            print("\nAttempting to upload to Truck Management")
            self.driver.find_element_by_xpath('//*[@id="timeIn"]').send_keys(self.timeIn)
            self.driver.find_element_by_xpath('//*[@id="driverName"]').send_keys(self.driverName)
            self.driver.find_element_by_xpath('//*[@id="license"]').send_keys(self.licenseNumber)
            self.driver.find_element_by_xpath('//*[@id="truckingCo"]').send_keys(self.truckingCompany)
            self.driver.find_element_by_xpath('//*[@id="gatePass"]').send_keys(self.gatePass)
            self.driver.find_element_by_xpath('//*[@id="dor"]').send_keys(self.dor)
            self.driver.find_element_by_xpath('//*[@id="code_chosen"]').click()
            self.driver.find_element_by_xpath('//*[@id="code_chosen"]/div/ul/li[' + str(codeMap[self.clientCode]) + ']').click()
            self.driver.find_element_by_xpath('/html/body/div[2]/form/div/div/div[9]/input').click()
            print("Truck Management Upload Successful")
            
        except KeyboardInterrupt:
            Helpers.Exceptions.user_interrupt(self.driver)
            
        except:
            print("An error occured while attempting to add driver to Truck Management")