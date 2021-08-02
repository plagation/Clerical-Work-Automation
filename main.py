# -*- coding: utf-8 -*-

import time
import getpass
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import TruckManagement, AlgomaShipment, BoscusShipment, RusalShipment, LBFosterShipment,BoscusTruckReception, ErametShipment, MillerShipment, CooperConsolidatedShipment
import easygui

def Boscus():
    if "BO9" in dor:
        return BoscusTruckReception.BoscusTruckReception(driver, driverName, truckingCompany, dor, licenseNum, clientCode)
    else:
        return BoscusShipment.BoscusShipment(driver, driverName, truckingCompany, dor, licenseNum, clientCode)
        
def Algoma():
    return AlgomaShipment.AlgomaShipment(driver, driverName, truckingCompany, dor, licenseNum, clientCode)
    
def LBFoster():
    return LBFosterShipment.LBFosterShipment(driver, driverName, truckingCompany, dor, licenseNum, clientCode)
    
def Rusal():
    return RusalShipment.RusalShipment(driver, driverName, truckingCompany, dor, licenseNum, clientCode, loadNum)

def Eramet():
    return ErametShipment.ErametShipment(driver, driverName, truckingCompany, dor, licenseNum, clientCode, tare)

def Miller():
    return MillerShipment.MillerShipment(driver, driverName, truckingCompany, dor, licenseNum, clientCode, tare)

def CooperConsolidated():
    return CooperConsolidatedShipment.CooperConsolidatedShipment(driver, driverName, truckingCompany, dor, licenseNum, clientCode)
    
def login_TC3():
    driver.maximize_window()
    driver.get("https://identity.qsl.com/u/login/identifier?state=hKFo2SBFbGlaT3Zjb0tZa3lrRGN2NGV3ZlprMzJYNlN0ZVNlaKFur3VuaXZlcnNhbC1sb2dpbqN0aWTZIDBDc2M4VnJkRGpDT0pzTEU4TTQ5RUhJQ0x3YTZEQ2N3o2NpZNkgYzRibDAyWVZjN1NWRWE2aXQ4RlFSVDRNOENaeU9RMWo")
    WebDriverWait(driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="username"]'))
    driver.find_element_by_xpath('//*[@id="username"]').send_keys(username + "@qsl.com"+ Keys.RETURN)
    """
    WebDriverWait(driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="i0118"]'))
    driver.find_element_by_xpath('//*[@id="i0118"]').send_keys(password + Keys.RETURN)
    WebDriverWait(driver,50).until(lambda x: x.find_element(By.XPATH, '//*[@id="idBtn_Back"]'))
    driver.find_element_by_xpath('//*[@id="idBtn_Back"]').click()
    """
    WebDriverWait(driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/div/div[1]/div[2]/div'))
    
    
def login_truck_management():
    driver.execute_script("window.open('about::blank','portal');")
    driver.switch_to.window("portal")
    driver.get("http://portal.nascochi.us/")
    driver.find_element_by_xpath('//*[@id="dismissModal"]').click()
    driver.find_element_by_xpath('//*[@id="truckAuth"]').click()
    WebDriverWait(driver, 20).until(lambda x: x.find_element(By.XPATH, '//*[@id="username"]'))
    time.sleep(.2)
    driver.find_element_by_xpath('//*[@id="username"]').send_keys(username)
    driver.find_element_by_xpath('//*[@id="password"]').send_keys(password)
    driver.find_element_by_xpath('//*[@id="signin"]').click()
    WebDriverWait(driver,50).until(lambda x: x.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/ul[1]/li[3]/a'))
    driver.get("http://tms.nascochi.us/Scale/Scale")
    WebDriverWait(driver,50).until(lambda x: x.find_element(By.XPATH, '//*[@id="timeIn"]'))

clientOptions = {"ALG":Algoma, "LBF IN":LBFoster, "LBF OUT":LBFoster, "BOS":Boscus, "RUS":Rusal, "ERA":Eramet, "MIL":Miller, "COOPCON":CooperConsolidated}
algomaNumbers = [12850, 14016, 19843, 34062, 38367, 41941, 42615, 45209, 48365, 56427, 65714, 74629, 87254, 82513, 87693, 84261]
cooperConsolidatedNumbers = ["0036584", "0036585"]

username = ""
password = ""

if username == "":
    username = getpass.getuser()
    
if password == "":
    password = easygui.enterbox("Enter password for Truck Management: ")

timeIn = easygui.enterbox("Enter time in: ")
driverName = easygui.enterbox("Enter driver name: ").lower().title()
truckingCompany = easygui.enterbox("Enter trucking company: ").lower().title()
while True:
    gatePass = easygui.enterbox("Enter gate pass number: ")
    try:
        int(gatePass)
        break
    except ValueError:
        print("Invalid Entry! Gatepass number must be an integer!")
        
dor = easygui.enterbox("Enter pickup number: ").upper()
licenseNum = easygui.enterbox("Enter license number: ").upper()

try:
    if "BOS" in dor or "BO9" in dor:
        clientCode = client = "BOS"
        
    elif "RAC" in dor:
        clientCode = client = "RUS"
        while True:
            loadNum = easygui.enterbox("Enter the load number: ")
            try:
                int(loadNum)
                break
            
            except ValueError:
                print("Invalid entry! Load number must be an integer!")
                
    elif "MKT" in dor:
        clientCode = "E"
        client = "ERA"
        
    elif dor in cooperConsolidatedNumbers:
        clientCode = "C"
        client = "COOPCON"
        
    elif int(dor) in algomaNumbers:
        clientCode = "C"
        client = "ALG"
        
    elif len(dor) == 12:
        clientCode = "PI"
        client = "MIL"
        
    else:
        clientCode = easygui.enterbox("Enter client code: ").upper()
        
except ValueError:
    clientCode = easygui.enterbox("Enter client code: ").upper()
    
if clientCode in ("PI", "E"):
    while True:
        tare = easygui.enterbox("Enter tare weight:")
        try:
            if int(tare) % 20 == 0:
                break
            else:
                print("Invalid Entry! Tare must be a multiple of 20!")
                
        except ValueError:
            print("Invalid Entry! Entry must be an integer!")
            
#open browser and perform initial login for TC3 and truck management
driver = webdriver.Chrome()

login_TC3()
tabOne = driver.window_handles[0]
login_truck_management()

#truckUpload = TruckManagement.TruckManagement(driver, timeIn, dor, gatePass, licenseNum, driverName, truckingCompany, clientCode)
#truckUpload.Upload()

driver.switch_to.window(tabOne)

shipment = clientOptions[client]()
shipment.CreateShipment()

#driver.quit()
