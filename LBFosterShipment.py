# -*- coding: utf-8 -*-
"""
Created on Thu Jul 22 11:58:24 2021

@author: kyle.conrad
"""

import Shipment

class LBFosterShipment(Shipment.Shipment):
    def set_pickupNumber(self):
        if(self.dor.contains(281778)):
            self.pickupNumber = "New Stock"
        else:
            self.pickupNumber = "Shipments"