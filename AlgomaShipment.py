# -*- coding: utf-8 -*-
"""
Created on Thu Jul 22 11:57:30 2021

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

class AlgomaShipment(Shipment.Shipment):
    def set_remarks(self):
        self.set_clerk()
        self.remarks = "Tarp load\nDrivers accept paperwork as is, signing waived due to COVID-19 restrictions.\n{}".format(self.clerk)
        
    def set_pickupNumber(self):
        self.pickupNumber = self.dor
        
    def add_load(self):
        WebDriverWait(self.driver,50).until(lambda x: x.find_element(By.XPATH,'//*[@id="viewport"]/article/section/form/section[1]/div[3]/nav/menu/a[2]'))
        self.bolNum = self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/form/section[1]/div[2]/div[1]/h3').text
        self.driver.get('https://tos.qsl.com/client-inventories/shipment-of-materials/' + self.bolNum + '/loading-request')
        WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/article/header[2]/aside/button'))
        self.driver.find_element_by_xpath('//*[@id="viewport"]/article/section/article/header[2]/aside/button').click()
        WebDriverWait(self.driver, 30).until(lambda x: x.find_element(By.XPATH, '/html/body/div[13]/div/div/article/form/div/div/button'))
        self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/form/div/div/button').click()
        
        WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '/html/body/div[13]/div/div/article/section/div/div/div[1]/div[4]/div/div[2]/span/button'))
        element = self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section/div/div/div[1]/div[1]/div[3]/span/span[1]/div/div')
        
        actions = ActionChains(self.driver)
        actions.context_click(element).perform()
        WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '/html/body/nav[2]/section/div[2]/div[3]/div/div/span[23]/div/label'))
        self.driver.find_element_by_xpath('/html/body/nav[2]/section/div[2]/div[3]/div/div/span[23]/div/label').click()
        self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section/div/div/div[1]/div[1]/div[4]/span/span[1]/div/div/span').click()
        WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '/html/body/div[13]/div/div/article/section/div/div/div[1]/div[4]/div/div[1]/div[9]/span'))
        coilQuantity = self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/header/div/h6[1]').text.split("/")[1]
        firstInDate = float('inf')
        dateArr = []
        
        for i in range(4, 4 + int(coilQuantity)):
            testVal = self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section/div/div/div[1]/div[' + str(i) + ']/div/div[1]/div[9]/span').text
            
            if "HOT" in testVal.upper():
                
                if firstInDate != 0:
                    dateArr = []
                    
                dateArr.append(i)
                firstInDate = 0
            
            else:
                if testVal[1] == "/":
                    testVal = "0" + testVal
                
                if testVal[4] == "/":
                    testVal = testVal[0,3] + "0" + testVal[3,]
                    
                testNum = int(testVal.replace("/", ""))
                
                if testNum < firstInDate:
                    dateArr = [i]
                    firstInDate = testNum
                    
                elif testNum == firstInDate:
                    dateArr.append(i)
                    
        k = 0
        for i in dateArr:
            j = i - k
            if k == 0:
                self.driver.execute_script("window.scrollBy(0,window.innerHeight)")
                time.sleep(.5)
                self.driver.find_element_by_xpath("/html/body/div[13]/div/div/article/section/div/div/div[1]/div[" + str(j) + "]/div/div[2]/span/button").click()
            else: 
                self.driver.execute_script("window.scrollBy(0,window.innerHeight)")
                time.sleep(.5)
                self.driver.find_element_by_xpath('/html/body/div[13]/div/div/article/section[2]/div/div/div[1]/div[' + str(j) + ']/div/div[2]/span/button').click()
            k += 1
            time.sleep(.5)
            
        time.sleep(3)
        
        try:
            self.driver.find_element_by_xpath('/html/body/div[13]/div/div/div[1]/header/menu/button[3]').click()
        except NoSuchElementException:
            self.driver.find_element_by_xpath('/html/body/div[13]/div/div/header/menu/button[3]').click()
            
        element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '/html/body/div[13]/div/div/div[1]/header/menu/button[2]'))
        element.click()
        
        element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/section[1]/div[2]/div[2]/div[2]'))
        element.click()
        
        element = WebDriverWait(self.driver,50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/section[1]/div[2]/div[2]/div[2]/ul/li[1]'))
        element.click()
        
        element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '/html/body/div[5]/div/form/header/menu/button[2]'))
        element.click()
        
        element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/section[1]/div[2]/div[1]/h3'))
        bol = element.text
        time.sleep(3)
        