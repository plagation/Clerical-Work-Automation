# -*- coding: utf-8 -*-
"""
Created on Fri Jul 23 13:26:21 2021

@author: kyle.conrad
"""

import BulkShipment

import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

class MillerShipment(BulkShipment.BulkShipment):
    def set_lot(self):
        self.lot = ""
        
    def set_analysis(self):
        self.analysis = ""
        
    def set_poNumber(self):
        self.poNumber = ""
        
    def set_remarks(self):
        self.set_clerk()
        self.set_poNumber()
        self.set_lot()
        self.set_analysis()
        self.remarks = "Order#: {}\nPO#: {}\nLot#: {}\nAnalysis: {}\nFreight Terms: {}\nTarp load\nDrivers accept paperwork as is, signing waived due to COVID-19 restrictions.\n{}".format(self.dor, self.poNumber, self.lot, self.analysis, self.clerk)