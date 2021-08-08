# -*- coding: utf-8 -*-

import TruckManagement, AlgomaShipment, BoscusShipment, RusalShipment, LBFosterShipment,BoscusTruckReception, ErametShipment, MillerShipment, CooperConsolidatedShipment
from Helpers import Helpers

import getpass, re, time, easygui, sys, traceback


from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

def Boscus():
    if "BO9" in dor:
        return BoscusTruckReception.BoscusTruckReception(driver, driverName, truckingCompany, dor, licenseNum, clientCode, timeIn)
    else:
        return BoscusShipment.BoscusShipment(driver, driverName, truckingCompany, dor, licenseNum, clientCode, timeIn)
        
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
    try:
        print("\nAttempting to login to TC3...")
        driver.get("https://identity.qsl.com/u/login/identifier?state=hKFo2SBFbGlaT3Zjb0tZa3lrRGN2NGV3ZlprMzJYNlN0ZVNlaKFur3VuaXZlcnNhbC1sb2dpbqN0aWTZIDBDc2M4VnJkRGpDT0pzTEU4TTQ5RUhJQ0x3YTZEQ2N3o2NpZNkgYzRibDAyWVZjN1NWRWE2aXQ4RlFSVDRNOENaeU9RMWo")
        element = WebDriverWait(driver, 120).until(lambda x: x.find_element(By.XPATH, '//*[@id="username"]'))
        element.send_keys(username + "@qsl.com"+ Keys.RETURN)
        try:
            print("Attempting to enter password...")
            element = WebDriverWait(driver, 5).until(lambda x: x.find_element(By.XPATH, '//*[@id="i0118"]'))
            element.send_keys(passwordTC3 + Keys.RETURN)
            element = WebDriverWait(driver, 10).until(lambda x: x.find_element(By.XPATH, '//*[@id="idBtn_Back"]'))
            element.click()
        except TimeoutException:
            print("No password field found")
            pass
        
        WebDriverWait(driver, 50).until(lambda x: x.find_element(By.XPATH, '//*[@id="viewport"]/article/section/div/div[1]/div[2]/div'))
        print("TC3 Login Successful")
    except:
        Helpers.Exceptions.unexpected_exception("login to TC3")
        driver.quit()
        sys.exit(1)
    
def login_truck_management():
    try:
        print("\nAttempting to login to Truck Management...")
        driver.execute_script("window.open('about::blank','portal');")
        driver.switch_to.window("portal")
        driver.get("http://portal.nascochi.us/")
        driver.find_element_by_xpath('//*[@id="dismissModal"]').click()
        driver.find_element_by_xpath('//*[@id="truckAuth"]').click()
        element = WebDriverWait(driver, 20).until(lambda x: x.find_element(By.XPATH, '//*[@id="username"]'))
        time.sleep(.2)
        element.send_keys(username)
        driver.find_element_by_xpath('//*[@id="password"]').send_keys(passwordTM)
        driver.find_element_by_xpath('//*[@id="signin"]').click()
        WebDriverWait(driver,50).until(lambda x: x.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/ul[1]/li[3]/a'))
        driver.get("http://tms.nascochi.us/Scale/Scale")
        WebDriverWait(driver,50).until(lambda x: x.find_element(By.XPATH, '//*[@id="timeIn"]'))
        print("Truck Management Login Successful")
    except:
        print("Error occured attempting to login to Truck Management")

def set_gatepass_values():
    
    global timeIn, driverName, truckingCompany, licenseNum, gatePass, dor, clientCode, client, tare, loadNum
    
    algomaNumbers = [12850, 14016, 19843, 34062, 38367, 41941, 42615, 45209, 48365, 56427, 65714, 74629, 87254, 82513, 87693, 84261]
    cooperConsolidatedNumbers = ["0036584", "0036585"]
    try:
    
        while True:
            timeIn = easygui.enterbox("Enter time in: ")
            if timeIn == None:
                sys.exit(1)
            elif ":" not in timeIn:
                print("Not a valid time! Time must be in format HH:MM!")
            else:
                if len(timeIn.split(":")[0]) == 1: 
                    timeIn = "0" + timeIn
                break
        driverName = easygui.enterbox("Enter driver name: ").lower().title()
        if driverName == None:
            sys.exit(1)
        truckingCompany = easygui.enterbox("Enter trucking company: ").lower().title()
        if truckingCompany == None:
            sys.exit(1)
        licenseNum = easygui.enterbox("Enter license number: ").upper()
        if licenseNum == None:
            sys.exit(1)
        while True:
            gatePass = easygui.enterbox("Enter gate pass number: ")
            if gatePass == None:
                sys.exit(1)
            try:
                int(gatePass)
                break
            except ValueError:
                print("Invalid Entry! Gatepass number must be an integer!")
                
        dor = easygui.enterbox("Enter pickup number: ").upper()
        if dor == None:
            sys.exit(1)
        try:
            if "BOS" in dor or "BO9" in dor:
                clientCode = client = "BOS"
                
            elif "RAC" in dor:
                clientCode = client = "RUS"
                while True:
                    if len(re.findall(r'^RAC-\d{6}$', dor)) == 1:
                        break
                    if "-" not in dor:
                        dor = "RAC-" + dor[3:]
                    if len(re.findall(r'^RAC-\d{6}$', dor)) == 0:
                        print("Invalid pickup number! Rusal pickup numbers must have 6 digits at their end, enter a valid pickup number!")
                        dor = easygui.enterbox("Enter pickup number: ").upper()
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
                if clientCode == None:
                    sys.exit(1)
                if clientCode in ("LBF IN", "LBF OUT"):
                    client = "LBFoster"
                    
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
                    
    except AttributeError:
        raise SystemExit
        
clientOptions = {"ALG":Algoma, "LBF IN":LBFoster, "LBF OUT":LBFoster, "BOS":Boscus, "RUS":Rusal, "ERA":Eramet, "MIL":Miller, "COOPCON":CooperConsolidated}

username = ""
passwordTC3 = ""
passwordTM = ""

try: 
    if username == "":
        username = getpass.getuser()

    if passwordTC3 == "":
        passwordTC3 = easygui.enterbox("Enter password for TC3: ")
    
    if passwordTM == "":
        passwordTM = easygui.enterbox("Enter password for Truck Management: ")
    
    set_gatepass_values()
        
    #open browser and perform initial login for TC3 and truck management
    options = Options()
    options.headless = True
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1920x1080')
    options.add_argument('start-maximized')
    options.add_experimental_option("prefs", {
        "download.default_directory": "C:\\Users\\" + getpass.getuser() + "\\Downloads",
        "download.prompt_for_download": False})
    
    driver = webdriver.Chrome(options = options)
    
    driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': "C:\\Users\\" + getpass.getuser() + "\\Downloads"}}
    command_result = driver.execute("send_command", params)
    
    login_TC3()
    tabOne = driver.window_handles[0]
    
    login_truck_management()
    
    #truckUpload = TruckManagement.TruckManagement(driver, timeIn, dor, gatePass, licenseNum, driverName, truckingCompany, clientCode).Upload()
    
    driver.switch_to.window(tabOne)
    
    shipment = clientOptions[client]().CreateShipment()
    
    driver.quit()
    
except SystemExit:
    print("Interrupted by user")
    sys.exit(1)
    
except KeyboardInterrupt:
    Helpers.Exceptions.user_interrupt(driver)
    
except:
    Helpers.Exceptions.unexpected_exception("run main method")


