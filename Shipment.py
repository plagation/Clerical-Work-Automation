# -*- coding: utf-8 -*-
"""
Created on Thu Jul 22 11:21:00 2021

@author: kyle.conrad
"""
import time, re, getpass, subprocess, sys, traceback

from Helpers import Helpers

from pathlib import Path

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

from io import StringIO

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

class Shipment:
    dirPath = str(Path.cwd()).replace("\\", "/")
    
    def __init__(self, driver, driverName, truckingCompany, dor, licenseNumber, clientCode):
        self.driver = driver
        self.driverName = driverName
        self.truckingCompany = truckingCompany
        self.dor = dor
        self.licenseNumber= licenseNumber
        self.clientCode = clientCode
        
    def set_carrierBill (self):
        self.carrierBill = self.dor
        
    def set_clerk(self):
        try:
            print("\nAttempting to set clerk...")
            self.clerk = ""
            for name in getpass.getuser().split("."):
                self.clerk = self.clerk + name[0].upper()
            print("Successfully set clerk as " + self.clerk)
        except:
            print("An unexpected error occurred while attempting to set clerk: \n")
            print(traceback.print_exc())
            
    def set_remarks(self):
        self.set_clerk()
        self.remarks = "Drivers accept paperwork as is, signing waived due to COVID-19 restrictions/n" + self.clerk

    def set_receiver(self):
        self.receiver = ""
    
    def set_address(self):
        self.address = ""
    
    def set_pickupNumber(self):
        self.pickupNumber = ""
        
    def set_bolNum(self):
        try:
            print("\nRetrieving BOL number...")
            url = self.driver.current_url
            self.bolNum = re.findall(r'\d{5,6}', url)[0]
            
            if len(self.bolNum) == 5:
                self.bolNum = "0" + self.bolNum
                
            print("BOL Number successfully set as " + self.bolNum)
            
        except KeyboardInterrupt:
            Helpers.Exceptions.user_interrupt(self.driver)
            
        except:
            Helpers.Exceptions.unexpected_exception("set BOL number")
            
    def convert_to_text(self, path):
        try:
            
            print("\nAttempting to parse release " + self.dor)
            
            rsrcmgr = PDFResourceManager()
            retstr = StringIO()
            codec = 'utf-8'
            laparams = LAParams()
            device = TextConverter(rsrcmgr, retstr,codec = codec, laparams=laparams)
            fp = open(path, 'rb')
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            password = ""
            maxpages = 0
            caching  = True
            pagenos = set()
            
            for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages,password = password,caching = caching, check_extractable=True):
                interpreter.process_page(page)
            
            self.convertedFile = retstr.getvalue()
            
            fp.close()
            device.close()
            retstr.close()
        
            print("Successfully parsed release")
            
        except FileNotFoundError:
            print("Release not found in default location, please select release...")
            file = Helpers.FileBrowser.get_file()
            
            if file == "":
                print("Exiting program...")
                sys.exit(1)
            else:
                self.convert_to_text(file)
                
        except KeyboardInterrupt:
            Helpers.Exceptions.user_interrupt(self.driver)
            
        except:
            Helpers.Exceptions.unexpected_exception("parse release " + self.dor)
        
    def set_variables(self):
        self.set_carrierBill()
        self.set_remarks()
        self.set_address()
        self.set_pickupNumber()
        self.set_receiver()
        
    def CreateShipment(self):
        self.set_variables()
        try:
            print("\nAttempting to create shipment...")
            self.driver.get('https://tos.qsl.com/client-inventories/shipment-of-materials')
            element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/section[1]/div[2]/div[2]/button'))
            time.sleep(1.5)
            element.click()
            #set port if not preset
            element = WebDriverWait(self.driver, 10).until(lambda x: x.find_element(By.XPATH, '/html/body/div[3]/div/form/header/menu/button[2]'))
            if self.driver.find_element_by_xpath('//*[@id="site"]/div/div[1]/div[1]').text != "Iroquois Landing":
                self.driver.find_element_by_xpath('//*[@id="react-select-18-input"]').send_keys("Iroquois Landing" + Keys().RETURN)
            element.click()
            WebDriverWait(self.driver,50).until(lambda x: x.find_element(By.XPATH, '//*[@id="driverName"]'))
            self.driver.find_element_by_xpath('//*[@id="react-select-2-input"]').send_keys(self.truckingCompany + Keys().RETURN)
            self.driver.find_element_by_xpath('//*[@id="driverName"]').send_keys(self.driverName)
            self.driver.find_element_by_xpath('//*[@id="carrierBill"]').send_keys(self.carrierBill)
            self.driver.find_element_by_xpath('//*[@id="transportationNumber"]').send_keys(self.licenseNumber)
            self.driver.find_element_by_xpath('//*[@id="specialInstructions"]').send_keys(self.remarks)
            self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/fieldset[2]/article/header/button').click()
            time.sleep(1)
            self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/fieldset[2]/article/section/section/div/menu/section/div[9]/button').click()
            element = WebDriverWait(self.driver,50).until(lambda x: x.find_element(By.XPATH, '//*[@id="scoped-search"]'))
            element.send_keys(self.pickupNumber)
            
            try:
                WebDriverWait(self.driver,10).until(lambda x: x.find_element(By.XPATH, '/html/body/div[6]/div/section/div/div/div[1]/div[4]/div'))
            
            except TimeoutException: 
                print("Order does not contain any items!")
                self.driver.quit()
                raise
                
            self.driver.find_element_by_xpath('/html/body/div[6]/div/section/div/div/div[1]/div[4]/div').click()
            time.sleep(.5)
            self.driver.find_element_by_xpath('/html/body/div[6]/div/header/menu/button[2]').click()
            time.sleep(2)
            
            if self.receiver != "":
                self.driver.find_element_by_xpath('//*[@id="shipmentDestinations[0].receiver"]/div/div[2]/button').click()
                self.driver.find_element_by_xpath('//*[@id="react-select-5-input"]').send_keys(self.receiver + Keys().RETURN)
                
                if self.address != "":
                    try:
                        self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/fieldset[2]/article/section/section/div/menu/div/div[1]/section/div[4]/div/label').click()
                    except:
                        pass
                    self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/fieldset[2]/article/section/section/div/menu/div/div[1]/section/div[4]/i').click()
                    element = WebDriverWait(self.driver, 20).until(lambda x: x.find_element(By.ID, "formattedAddress"))
                    element.send_keys(self.address)
                    self.driver.find_element_by_xpath('/html/body/div[7]/div/header/menu/button[2]').click()
                else:  
                    self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/fieldset[2]/article/section/section/div/menu/div/div[1]/section/div[3]/div/div/div/div').click()
                    self.driver.find_element_by_xpath('//*[@id="react-select-15-option-0"]').click()
            print("Shipment successfully created")
            
        except KeyboardInterrupt:
            Helpers.Exceptions.user_interrupt(self.driver)
    
        except:
            Helpers.Exceptions.unexpected_exception("create shipment")
            
    #Save newly created entry
        try:
            print("\nAttempting to save shipment...")
            element = WebDriverWait(self.driver,20).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/form/section[2]/div/div[2]/button[2]'))
            element.click()
            print("Shipment successfully saved")
            
        except KeyboardInterrupt:
            Helpers.Exceptions.user_interrupt(self.driver)
            
        except:
            Helpers.Exceptions.unexpected_exception("save shipment")
            
        time.sleep(1)
        self.set_bolNum()
        
        self.add_load()
            
    def add_load(self):
        try:
            print("\nAttempting to add load...")
            element = WebDriverWait(self.driver,50).until(lambda x: x.find_element(By.XPATH,'//*[@id="viewport"]/article/section/form/section[1]/div[3]/nav/menu/a[2]'))
            element.click()
            print("Load successfully added")
        
        except KeyboardInterrupt:
            Helpers.Exceptions.user_interrupt(self.driver)
        
        except: 
            Helpers.Exception.unexpected_exception("add load")