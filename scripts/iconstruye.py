from selenium import webdriver

from time import *
from datetime import date
from time import *
from selenium import webdriver
from selenium.webdriver.support.select import Select
from scripts.fileOperations import FileOperation
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as EC


class botService:

    def __init__(self, driverOptions, downloadPath, driverPath):
        self.driverOptions = driverOptions
        self.downloadPath = downloadPath
        self.pathDriver = driverPath
        self.url = "https://www.iconstruye.com/index.html"
        self.operation = FileOperation.Operation(downloadPath)
        self.i=0
        self.username = "fvalenzuela"
        self.password = "Santi2017"
        self.orgName = "mkci"
        self.today = date.today()
        self.lastDate = self.today.strftime("%d-%m-%Y")
        self.initialDate = "01-06-2019"
        self.customers = [
            {
                "customerID": "51083",
                "nameFile": "E011077"
            },
            {
                "customerID": "55020",
                "nameFile": "E011083"
            },
            {
                "customerID": "55021",
                "nameFile": "E011084"
            },
            {
                "customerID": "55700",
                "nameFile": "E011085"
            },
            {
                "customerID": "55914",
                "nameFile": "E011086"
            },
            {
                "customerID": "52016",
                "nameFile": "E024022"
            },
            {
                "customerID": "51841",
                "nameFile": "E032001"
            },

            #nuevos cleintes
            {
                "customerID": "58356",
                "nameFile": "E024024"
            },
            {
                "customerID": "58739",
                "nameFile": "E024025"
            },
            {
                "customerID": "58716",
                "nameFile": "E031003"
            },
            {
                "customerID": "58713",
                "nameFile": "E011088"
            },
        ]

    def run(self):


        for customer in self.customers:
            self.operation.removeFile()
            self.downloadReports(customer["customerID"])
            self.operation.renameFile(customer["nameFile"])
            self.operation.convertFormat()
            self.operation.sendEmail(customer["nameFile"],"Se adjunta el archivo asociado.")
            print("se envio: "+customer["nameFile"])
            self.operation.removeFile()

        # self.operation.removeFile()
        # self.downloadReportsFacturas()
        #
        # self.operation.removeFile()
        # self.downloadReportsSubContrato()




    def getNameTc(self,i):
        if(i==0):
            return "BCI-siime-saldo-TC.xlsx"
        if (i == 1):
            return "BCI-h&a-saldo-TC.xlsx"
        if (i == 2):
            return "BCI-iappem-saldo-TC.xlsx"

    def getNameSaldo(self,i):
        if(i==0):
            return "BCI-siime-saldo-cuentacorriente.xls"
        if (i == 1):
            return "BCI-h&a-saldo-cuentacorriente.xls"
        if (i == 2):
            return "BCI-iappem-saldo-cuentacorriente.xls"



    def login(self, driver):
        driver.get(self.url)
        sleep(5)
        driver.find_element_by_class_name("js-open-login").click()
        sleep(5)
        iframe = driver.find_element_by_class_name("login ")
        driver.switch_to.frame(iframe)
        driver.find_element_by_id("txtUsuario").send_keys(self.username)
        driver.find_element_by_id("txtOrganizacion").send_keys(self.orgName)
        driver.find_element_by_id("txtClave").send_keys(self.password)
        sleep(5)
        driver.find_element_by_xpath("//input[@name='btnIngresar']").click()
        driver.switch_to_default_content()

    def downloadReportsSubContrato(self):
        with webdriver.Chrome(executable_path=self.pathDriver, chrome_options=self.driverOptions) as driver:
            try:
                self.login(driver)
                sleep(5)
                actions = ActionChains(driver)
                target= driver.find_element_by_css_selector("div#ControlMenu1_sec table tr td:nth-of-type(7)")
                actions.move_to_element(target).perform()
                sleep(1)
                target2=target.find_element_by_class_name("SubMenu ").find_element_by_css_selector("div table tr tr:nth-of-type(6)")
                actions.move_to_element(target2).perform()
                sleep(1)
                target3=target2.find_element_by_class_name("SubMenu ").find_element_by_css_selector("div table tr tr:nth-of-type(3)")
                target3.click()
                sleep(5)
                driver.switch_to.frame("detalle")
                select_element = driver.find_element_by_id("lstOrgC")
                select_object = Select(select_element)
                select_object.select_by_value('-1')
                sleep(10)
                select_element = driver.find_element_by_id("lstMoneda")
                select_object = Select(select_element)
                select_object.select_by_value('-1')
                driver.find_element_by_id("rngFechaCreacion1FECHADESDE").clear()
                driver.find_element_by_id("rngFechaCreacion1FECHADESDE").send_keys(self.initialDate)
                driver.find_element_by_id("rngFechaCreacion1FECHAHASTA").clear()
                driver.find_element_by_id("rngFechaCreacion1FECHAHASTA").send_keys(self.lastDate)
                sleep(7)
                driver.find_element_by_id("btnBuscar").click()
                sleep(5)
                try:
                    driver.switch_to.alert.accept()
                except:
                    pass
                sleep(5)
                driver.find_element_by_id("lnkExcel").click()
                sleep(7)
                driver.close()
            except:
                driver.close()
                self.downloadReportsSubContrato()

    def downloadReports(self,customer):
        with webdriver.Chrome(executable_path=self.pathDriver, chrome_options=self.driverOptions) as driver:
            try:
                self.login(driver)
                sleep(7)
                actions = ActionChains(driver)
                target= driver.find_element_by_css_selector("div#ControlMenu1_sec table tr td:nth-of-type(6)")
                actions.move_to_element(target).perform()
                sleep(7)
                target2=target.find_element_by_class_name("SubMenu ").find_element_by_css_selector("div table tr tr:nth-of-type(12)")
                target2.click()
                sleep(7)
                driver.switch_to.frame("ventana")
                select_element = driver.find_element_by_id("lstCentroGestion")
                select_object = Select(select_element)
                select_object.select_by_value(customer)
                sleep(3)
                driver.find_element_by_id("ctrRangoFechasFECHADESDE").clear()
                driver.find_element_by_id("ctrRangoFechasFECHADESDE").send_keys(self.initialDate)
                select_element = driver.find_element_by_id("lstEstadosLinea")
                select_object = Select(select_element)
                select_object.select_by_value('-1')
                sleep(7)
                driver.find_element_by_id("btnBuscar").click()
                sleep(10)
                driver.switch_to.alert.accept()
                sleep(5)
                driver.find_element_by_id("lnkExcel").click()
                sleep(7)
                driver.close()
            except :
                driver.close()
                self.downloadReports(customer)

    def downloadReportsFacturas(self):
        with webdriver.Chrome(executable_path=self.pathDriver, chrome_options=self.driverOptions) as driver:
            try:
                self.login(driver)
                sleep(5)
                actions = ActionChains(driver)
                target= driver.find_element_by_css_selector("div#ControlMenu1_sec table tr td:nth-of-type(10)")
                actions.move_to_element(target).perform()
                sleep(1)
                target2=target.find_element_by_class_name("SubMenu ").find_element_by_css_selector("div table tr tr:nth-of-type(4)")
                target2.click()
                sleep(5)
                driver.switch_to.frame("ventana")
                select_element = driver.find_element_by_id("lstOrgC")
                select_object = Select(select_element)
                select_object.select_by_value('-1')
                driver.find_element_by_id("RgFechaRecepcion_desde_I").clear()
                driver.find_element_by_id("RgFechaRecepcion_desde_I").send_keys(self.initialDate)
                driver.find_element_by_id("btnVerificaBuscar").click()
                sleep(7)
                driver.find_element_by_id("btnBuscar").click()
                sleep(5)
                driver.switch_to.alert.accept()
                sleep(5)
                driver.find_element_by_id("lnkExcel").click()
                sleep(7)
                driver.close()
            except :
                driver.close()
                self.downloadReportsFacturas()





