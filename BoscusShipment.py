# -*- coding: utf-8 -*-
"""
Created on Thu Jul 22 11:57:54 2021

@author: kyle.conrad
"""

import Shipment
import time
import re
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException

class BoscusShipment(Shipment.Shipment):
    def set_receiver(self):
        self.receiver = re.findall(r'\w+\s{0,1}\w*,\s\w{2}\s+\d{5}', self.convertedFile)
        self.receiver = "MENARDS " + self.receiver[0].split(',')[0]
        
    def set_address(self):
        self.address = ""
        
    def set_poNumber(self):
        self.poNumber = re.findall(r'PO:\w{4}\d+', self.convertedFile)[0].replace("PO:", "")
                
    def set_skuNumber(self):
        fileLines = self.convertedFile.split("\n")
        self.skuNumber = fileLines[-7].strip().upper()
        if "FLEX" in self.skuNumber:
            if "SKU" in fileLines[-8].upper():
                self.skuNumber = fileLines[-8].strip().upper()
            else: 
                i = -len(self.bundleCounts) - 6
                self.skuNumber = fileLines[i - 1].strip().upper()
                while i <= -8:
                    self.skuNumber = self.skuNumber + "/" + fileLines[i].strip().upper()
                    i += 1
        
    def set_remarks(self):
        self.set_clerk()
        self.set_poNumber()
        self.set_skuNumber()
        self.remarks = ("{}\nPO#: {}\nSKU#: {}\nWeights are estimates only and as such are subject to review and revision as necessary."
        "\nDrivers accept paperwork as is, signing waived due to COVID-19 restrictions.\n{}"
        ).format(self.dor, self.poNumber, self.skuNumber, self.clerk) 
    
    def set_pickupNumber(self):
        self.pickupNumber = "Nasco Stock"
        
    def set_variables(self):
        self.parse_pdf()
        self.set_pickupNumber()
        self.set_carrierBill()
        self.set_address()
        self.set_poNumber()
        self.set_receiver()
        self.set_remarks()
        self.set_skuNumber()
        
    def parse_pdf(self):
        #Specific to network (Modification for scale necessary)
        self.convert_to_text("Z:/SCALE OFFICE/Boscus/Current Releases/" + self.dor + "-VTE-B80-000.pdf")
        self.descriptions = re.findall(r'\d{1}x\d{1}x\d+[\'( 92\-5/8)]+[\'\- 96]+', self.convertedFile)
        i = 0
        while i < len(self.descriptions):
            self.descriptions[i] = self.descriptions[i].replace("'", "").replace(" ", "")
            if "92" in self.descriptions[i]:
                self.descriptions[i] = self.descriptions[i][0:4] + "92 5/8"
            elif "96" in self.descriptions[i]:
                self.descriptions[i] = self.descriptions[i][0:4] + "96"
            i += 1
        
        self.pieceCounts = re.findall(r'\s{2}\d{3}\s{2}', self.convertedFile)

        self.bundleCounts = re.findall(r'\s{2}\d{1,2}\s{2}', self.convertedFile)
        
        if len(self.bundleCounts) > len(self.pieceCounts):
            self.bundleCounts.pop(len(self.bundleCounts)-1)
            
        i = 0 
        while i < len(self.bundleCounts):
            self.bundleCounts[i] = self.bundleCounts[i].replace(" ", "")
            self.pieceCounts[i] = self.pieceCounts[i].replace(" ", "")
            i += 1
            
    def add_load(self):
        WebDriverWait(self.driver,50).until(lambda x: x.find_element(By.XPATH,'//*[@id="viewport"]/article/section/form/section[1]/div[3]/nav/menu/a[2]'))
        self.bolNum = self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/section[1]/div[2]/div[1]/h3').text
        self.driver.get('https://tos.qsl.com/client-inventories/shipment-of-materials/' + self.bolNum + '/loading-request')
        WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/article/header[2]/aside/button'))
        self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/article/header[2]/aside/button').click()
        WebDriverWait(self.driver, 30).until(lambda x: x.find_element(By.XPATH, '//*[@id="react-select-9-input"]'))
        self.driver.find_element_by_xpath('//*[@id="react-select-9-input"]').send_keys("Description" + Keys().RETURN)
        time.sleep(.25)
        self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/form/div/div/button').click()
        WebDriverWait(self.driver,50).until(lambda x: x.find_element(By.XPATH, '/html/body/div[13]/div/div/article/section/div/div/div[1]/div[1]/div[3]/span/span[1]/div/div/span'))
        time.sleep(2.5)
        
        #adds order and description to viewable attributes and removes dimensions
        actions = ActionChains(self.driver)
        actions.context_click(self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section/div/div/div[1]/div[1]/div[3]/span/span[1]/div/div/span')).perform()
        WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '/html/body/nav[2]/section/div[2]/div[3]/div/div/span[23]/div/label'))
        self.driver.find_element_by_xpath('/html/body/nav[2]/section/div[2]/div[3]/div/div/span[23]/div/label').click()
        self.driver.find_element_by_xpath('/html/body/nav[2]/section/div[2]/div[3]/div/div/span[7]/div/label').click()
        self.driver.find_element_by_xpath('/html/body/nav[2]/section/div[2]/div[3]/div/div/span[9]/div/label').click()
        self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section/div/div/div[1]/div[1]/div[4]/span/span[1]/div/div/span').click()
        self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section/div/div/div[1]/div[1]/div[9]/span/span[1]/div/div/span/span[2]/div').click()
        time.sleep(3)
        self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section/div/div/div[1]/div[1]/div[9]/span/span[1]/div/div/span/span[2]/div').click()
        time.sleep(3)
        
        #loops through all sizes ordered, adding them to the order
        i = 0
        while i < len(self.descriptions):
            
            j = 0
            n = 0
            self.driver.find_element_by_xpath('//*[@id="identifiers.0.identifierValue"]').send_keys(self.descriptions[i] + Keys().RETURN)
            time.sleep(3)
            
            try:
                
                kStart = int(self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/header/div/h6[1]').text.split("/")[1])
                firstEntry = True
                
            except:
                
                kStart = int(self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/header[2]/div/h6[1]').text.split("/")[1])
                firstEntry = False
                
            k = kStart
            
            while j+1 <= k and n != int(self.bundleCounts[i]):
                
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                
                #check that you have not added bundles to the order yet (changes HTML)
                if firstEntry:
                    
                    siteDescription = self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section/div/div/div[1]/div[' + str(j + 4) + ']/div/div[1]/div[8]/span').text
                    sitePieces = self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section/div/div/div[1]/div[' + str(j + 4) + ']/div/div[1]/div[5]/span').text
                    
                    #cross-check the description and the pieceCount of each item in system with order
                    if self.descriptions[i] == siteDescription.replace("'","").replace('"',"") and self.pieceCounts[i] == sitePieces:
                        
                        siteQuantity = self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section/div/div/div[1]/div[' + str(j + 4) + ']/div/div[1]/div[7]/span').text
                        self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section/div/div/div[1]/div[' + str(j + 4) + ']/div/div[2]/span/button').click()
                        time.sleep(.5)
                        
                        if int(siteQuantity) == 1:
                            n += 1
                            
                        else:
                            
                            WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="quantity"]'))
                            
                            if int(siteQuantity) + n >= int(self.bundleCounts[i]):
                                
                                self.driver.find_element_by_xpath('//*[@id="quantity"]').send_keys(str(int(self.bundleCounts[i]) - n) + Keys().RETURN)
                                time.sleep(.5)
                                n += int(self.bundleCounts[i]) - n
                            
                            else: 
                                
                                self.driver.find_element_by_xpath('//*[@id="quantity"]').send_keys(siteQuantity + Keys().RETURN)
                                time.sleep(.5)
                                n += int(siteQuantity)
                            
                        k -= 1
                        j -= 1
                        
                else:
                    
                    siteDescription = self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section[2]/div/div/div[1]/div[' + str(j + 4) + ']/div/div[1]/div[8]/span').text
                    sitePieces = self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section[2]/div/div/div[1]/div[' + str(j + 4) + ']/div/div[1]/div[5]/span').text

                    #cross-check the description and the pieceCount of each item in system with order
                    if self.descriptions[i] == siteDescription.replace("'","").replace('"',"") and self.pieceCounts[i] == sitePieces:
                        
                        siteQuantity = self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section[2]/div/div/div[1]/div[' + str(j + 4) + ']/div/div[1]/div[7]/span').text
                        self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section[2]/div/div/div[1]/div[' + str(j + 4) + ']/div/div[2]/span/button').click()
                        time.sleep(.5)
                        
                        if int(siteQuantity) == 1:
                            n += 1
                            
                        else:
                            
                            WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="quantity"]'))
                            
                            if int(siteQuantity) + n >= int(self.bundleCounts[i]):
                                
                                self.driver.find_element_by_xpath('//*[@id="quantity"]').send_keys(str(int(self.bundleCounts[i]) - n) + Keys().RETURN)
                                time.sleep(.5)
                                n += int(self.bundleCounts[i]) - n
                                
                            else:
                                self.driver.find_element_by_xpath('//*[@id="quantity"]').send_keys(siteQuantity + Keys().RETURN)
                                time.sleep(.5)
                                n += int(siteQuantity)
                                
                        k -= 1
                        j -= 1
                
                j += 1
                print(str(n) + "/" + str(self.bundleCounts[i]) + " bundles added of " + self.descriptions[i])
                self.driver.find_element_by_xpath('//*[@id="identifiers.0.identifierValue"]').clear()
                
            i += 1
            
        time.sleep(.5)
        self.driver.find_element_by_xpath('/html/body/div[13]/div/div/div[1]/header/menu/button[3]').click()
        time.sleep(2)
        
        try:
            self.driver.find_element_by_xpath('/html/body/div[13]/div/div/div[1]/header/menu/button[2]').click()
        except NoSuchElementException:
            self.driver.find_element_by_xpath('/html/body/div[13]/div/div/header/menu/button[2]').click()
            