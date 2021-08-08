# -*- coding: utf-8 -*-
"""
Created on Mon Aug  2 14:02:46 2021

@author: kyle.conrad
"""

from AlgomaShipment import AlgomaShipment
from Helpers import Helpers

import time, subprocess
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException

class CooperConsolidatedShipment(AlgomaShipment):
    
    def add_load(self):
        try:
            print("\nAttempting to add items to shipment...")
            WebDriverWait(self.driver,50).until(lambda x: x.find_element(By.XPATH,'//*[@id="viewport"]/article/section/form/section[1]/div[3]/nav/menu/a[2]'))
            self.driver.get('https://tos.qsl.com/client-inventories/shipment-of-materials/' + self.bolNum + '/loading-request')
           
            element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/article/header[2]/aside/button'))
            element.click()
            
            element = WebDriverWait(self.driver, 30).until(lambda x: x.find_element(By.XPATH, '/html/body/div[13]/div/div/article/form/div/div/button'))
            element.click()
            
            element = WebDriverWait(self.driver, 30).until(lambda x: x.find_element(By.XPATH, '/html/body/div[13]/div/div/div[1]/header/menu/button[1]'))
            element.click()
            
            element = WebDriverWait(self.driver, 30).until(lambda x: x.find_element(By.XPATH, '/html/body/div[14]/div/footer/button[2]'))
            element.click()
            
            element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/section[1]/div[2]/div[2]/div[2]'))
            time.sleep(1)
            element.click()
            
            element = WebDriverWait(self.driver,50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/section[1]/div[2]/div[2]/div[2]/ul/li[1]'))
            element.click()
            
            element = WebDriverWait(self.driver, 50).until(lambda x: x.find_element(By.XPATH, '/html/body/div[5]/div/form/header/menu/button[2]'))
            element.click()
            
            time.sleep(2)
            print("Items successfully added to shipment")
        except:
            Helpers.Exceptions.unexpected_exception("add items to shipment")
        
        try:
            print("\nAttempting to print loading request...")
            self.print_loading_request()
            print("Loading request successfully printed")
            
        except:
            Helpers.Exceptions.unexpected_exception("print loading request")