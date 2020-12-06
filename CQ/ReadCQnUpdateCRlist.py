#
# All packages import
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

#
# get browser handle by driver and open URL with browser
#driver = webdriver.Firefox()
driver = webdriver.Chrome('chromedriver')
driver.get("https://clearquest.alstom.hub/cqweb/restapi/CQat/atvcm/QUERY/47363098?format=HTML&noframes=true")
print(driver.title)

#
# login to CQ to get access of the query (auto enter username and password)
time.sleep(1)
driver.switch_to.active_element
username = driver.find_element_by_name('loginId_Id')
username.send_keys("pdixit")
time.sleep(1)
password = driver.find_element_by_name('passwordId')
password.send_keys("passwd@JUN2018")

#
# auto click excel download
time.sleep(1)
connect_button = driver.find_element_by_id('loginButtonId')
connect_button.click()
time.sleep(5)
export_button = driver.find_element_by_id('dijit_form_ComboButton_1_arrow')
export_button.click()
time.sleep(5)
export_button = driver.find_element_by_id('dijit_MenuItem_35_text')
export_button.click()

time.sleep(10)

print(driver.current_url)

#
#close the driver
driver.close()