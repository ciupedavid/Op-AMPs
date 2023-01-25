import unittest
import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from decimal import Decimal
import math
from datetime import datetime
import time
import json
import chromedriver_autoinstaller

class OPAMP(unittest.TestCase):

    def setUp(self):
        # driver instance
        chromedriver_autoinstaller.install()
        self.driver = webdriver.Chrome()
        with open(r'paths.json') as d:
            self.nimbleData = json.load(d)['Nimble'][0]

    def test_export(self):
        driver = self.driver
        driver.set_window_position(-1000, 0)
        driver.maximize_window()
        driver.get('https://beta-tools.analog.com/noise/')
        if (self.nimbleData['filter_frequency'] != 0): #self.nimbleData['resistance_input'
            driver.get('https://beta-tools.analog.com/noise/#session=HO1dbVi7oU273IbstN6GiA&step=nsONZTDZTletQ0rqfFMQ0g')
            print("Filter was added.")
        else:
            driver.get('https://beta-tools.analog.com/noise/#session=0460LuqFaUGaPwEazOdZ9w&step=ZiZ21WxLTiiX7jikxx07Aw')
            print("No filther was added.")
        time.sleep(1)

        #converting kino, Mega
        d1 = {'k': 1000,'M': 1000000}
        def text_to_num1(text):
            if text[-1] in d1:
                num, magnitude = text[:-1], text[-1]
                return float(num) * d1[magnitude]
            else:
                return Decimal(text)
        new_rvalue = text_to_num1(self.nimbleData['rvalue'])

        #converting from femto, pico, nano, micro
        d2 = {'f': 0.000000000000001, 'p':0.000000000001, 'n':0.000000001, 'u':0.000001}
        def text_to_num2(text):
            if text[-1] in d2:
                num, magnitude = text[:-1], text[-1]
                return float(num) * d2[magnitude]
            else:
                return Decimal(text)
        new_c2value = text_to_num2(self.nimbleData['c2value'])
        
        # RVALUE Value to Position
        # If the value is less than or equal to 0 then we use minpos  
        def val_to_pos_rvalue(value):
            minpos = 1
            maxpos = 10000
            minval = math.log(10)
            maxval = math.log(10000000)
            scale = (maxval - minval) / (maxpos - minpos)
            global rposition
            rposition = minpos + (math.log(value) - minval) / scale
            return (rposition)
        val_to_pos_rvalue(new_rvalue)
        
        # C2 Value to Position
        def val_to_pos_c2value(value):
            minpos = 1
            maxpos = 10000
            minval = math.log(1e-13)
            maxval = math.log(0.000001)
            scale = (maxval - minval) / (maxpos - minpos)
            global c2position
            c2position = minpos + (math.log(value) - minval) / scale
            return (c2position)
        val_to_pos_c2value(new_c2value)

        # Accept cookies
        WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.XPATH, "//body/div[@id='base-container']/div[@id='noise-spinner']/div[1]")))
        time.sleep(7)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="cookie-consent-container"]/div/div/div[2]/div[1]/a'))).click()
        time.sleep(1)
        
        # Click on already dragged Amplifier
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//body/div[@id='base-container']/div[@id='main-content-container']/div[@id='application-view']/div[@id='build-signal-chain-tab-content']/div[@id='adi-signal-chain-row']/div[@id='analog-signal-chain-group']/div[@id='signal-chain-drop-area']/table[1]/tr[1]/td[1]/div[1]/div[2]/div[2]/*[1]"))).click()

        # Amp Configuration
        time.sleep(1)
        driver.execute_script(f"document.querySelector('#amp-gain-2').value={self.nimbleData['gain']}")
        
        # Select specific AMP
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#text6747-2-6 > tspan.schematic-edit-icon.schematic-part-edit-selection-link.schematic-edit-selection-link"))).click()
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#filter-0"))).send_keys(self.nimbleData['device'])
        time.sleep(1)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#device-table > div.slick-pane.slick-pane-top.slick-pane-left > div.slick-viewport.slick-viewport-top.slick-viewport-left > div > div"))).click()
        time.sleep(1)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#partSelectModal2 > div.modal.fade.in.show > div > div > form > div > button.btn.btn-primary"))).click()        
        time.sleep(5)

        # Resitance slider value
        self.scrollToRValue(rposition, driver)
        
        # C2 slider value
        self.scrollToC2Value(c2position, driver)

        # Use this Amplifier button
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#config-signal-chain-item-modal > div.modal.fade.in.show > div > div > form > div > button.btn.btn-primary"))).click()
        
        #Select Next Steps tab
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#next-steps-tab"))).click()
        time.sleep(1)

        device = self.nimbleData['device']
        downloads_path = self.nimbleData['downloads_path']
        gain = self.nimbleData['gain']
        #current_date = self.nimbleData['current_date']
        current_date = datetime.today().strftime(("%B %d, %Y"))
        print("date: ", current_date)

        l = driver.current_url
        device_url = device + 'URL G' + gain + '.txt'
        with open(device_url, 'w') as f:
            f.write(l)
        print(l)
        time.sleep(1)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "body.ember-application:nth-child(2) div.tab-content:nth-child(2) div.download-area div.download-individual-buttons div.download-button-row:nth-child(1) button.btn.btn-primary:nth-child(1) > span:nth-child(1)"))).click()
        time.sleep(1)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "body.ember-application:nth-child(2) div.tab-content:nth-child(2) div.download-area div.download-individual-buttons div.download-button-row:nth-child(1) button.btn.btn-primary:nth-child(2) > span:nth-child(1)"))).click()

        time.sleep(3)
        project_path = os.getcwd()

        if not os.path.exists(project_path + '\\' + device):
            os.makedirs(project_path + '\\' + device)
        dir_list = os.listdir()
        #print(dir_list)
        #print(project_path + device)

        ltspice_download_path = downloads_path + 'LTSpice ' + current_date + '.zip'
        shutil.move(ltspice_download_path, project_path + '/' + device)
        nimble_download_path = downloads_path + 'Raw Data Export - ' + current_date + '.zip'
        shutil.move(nimble_download_path, project_path + '/' + device)
        device_download_path = device + 'URL G' + gain + '.txt'
        shutil.move(device_download_path, project_path + '/' + device)
        time.sleep(1)

        downloaded_nimble_path = project_path + '\\' + device + '\\' + 'Raw Data Export - ' + current_date + '.zip'
        new_nimble_name = project_path + '\\' + device + '\\' + 'Nimble - ' + device + ' G' + gain + '.zip'
        downloaded_ltspice_path = project_path + '\\' + device + '\\' + 'LTspice ' + current_date + '.zip'
        new_ltspice_name = project_path + '\\' + device + '\\' + 'LTspice - ' + device + ' G' + gain + '.zip'
        os.rename(downloaded_nimble_path, new_nimble_name)
        os.rename(downloaded_ltspice_path, new_ltspice_name)

        time.sleep(2)

    @staticmethod
    def scrollToRValue(value: int, driver):
        driver.execute_script(f"document.querySelector('#rscale-slider').value = {rposition}; document.querySelector('#rscale-slider').dispatchEvent(new Event('input'));")
         
    @staticmethod
    def scrollToC2Value(value: int, driver):
        driver.execute_script(f"document.querySelector('#c2-slider').value = {c2position}; document.querySelector('#c2-slider').dispatchEvent(new Event('input'));")    

    def tearDown(self):
        self.driver.quit()

if __name__ == '__main__':
    unittest.main()

# WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, ""))).click()