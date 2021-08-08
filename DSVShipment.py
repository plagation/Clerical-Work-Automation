# -*- coding: utf-8 -*-
"""
Created on Sat Aug  7 20:12:38 2021

@author: kyleb
"""
from Shipment import Shipment

class DSVShipment(Shipment):
    
    def __init__(self, orderNum):
        self.orderNum = orderNum
        
    def set_mark(self):
        self.mark = "Placeholder, will parse excel file using orderNum to find last hub for given orderNum and it's mark"
        
    def set_driverName(self):
        self.driverName = self.clerk
    
    def set_truckingCompany(self):
        self.truckingCompany = "Sycamore"
        
    def set_licenseNumber(self):
        self.licenseNumber = self.clerk
    
    def set_dor(self):
        self.dor = self.clerk
        
    def set_remarks(self):
        self.remarks = ""

    def print_loading_certificates(self):
        "Placeholder"
    
    def set_variables(self):
        self.set_mark()
        self.set_clerk()
        self.set_remarks()
        self.set_driverName()
        self.set_truckingCompany()
        self.set_licenseNumber()
        self.set_dor()
        
            
    def add_load(self):
        "Placeholder"
        self.set_variables()
        