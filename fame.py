from bs4 import BeautifulSoup
import time
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options



def login(username, password):
    driver.get('https://libguides.kcl.ac.uk/az.php?a=f')
    driver.find_element_by_xpath('//*[@id="s-lg-az-name-f"]/div[5]/div[1]/a').click()
    driver.switch_to.window(driver.window_handles[1])

    driver.find_element_by_xpath('//*[@id="organisation-select-view"]/p[2]/a').click()
    driver.find_element_by_xpath('// *[ @ id = "combobox"]/ option[1123]').click()
    time.sleep(1)
    driver.find_element_by_xpath(' // *[ @ id = "submit-btn"]').click()
    usernameBox = WebDriverWait(driver, 3).until(EC.visibility_of_element_located
                                                  ((By.XPATH, '//*[@id="i0116"]')))
    usernameBox.send_keys(username)
    driver.refresh()
    Next = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'idSIButton9')))
    Next.click()
    usernameBox = WebDriverWait(driver, 3).until(EC.visibility_of_element_located
                                                  ((By.XPATH, '//*[@id="i0116"]')))
    usernameBox.send_keys(username)
    Next = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'idSIButton9')))
    Next.click()
    passwordBox = driver.find_element_by_xpath('//*[@id="i0118"]')
    passwordBox.send_keys(password)
    Next = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'idSIButton9')))
    Next.click()
    Next = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'idBtn_Back')))
    Next.click()
    time.sleep(8)
    # go to default set results view
    # driver.get("https://fame4.bvdinfo.com/version-20201112/fame/1/Companies/List/Display/6e88ea9f-7831-eb11-9009-00155d115b03");
    driver.get("https://fame4.bvdinfo.com/version-20201112/fame/1/Companies/List")
    driver.implicitly_wait(10)

def export_set(start, end):
    # while start < 7778452:
    while start < 266420:
        # click 'excel'
        excel = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/section[3]/div[1]/div[2]/div[2]/div[1]/ul/li[3]')))
        excel.click()
        time.sleep(2)
        # choose range
        # options = WebDriverWait(driver, 2).until(
        #     EC.element_to_be_clickable((By.XPATH, '//*[@id="component_RangeOptionSelectedId"]/option[4]')))
        # options.click()


        file_name = WebDriverWait(driver, 2).until(
            EC.visibility_of_element_located((By.NAME, 'component.FileName')))
        from_box = WebDriverWait(driver, 2).until(
            EC.visibility_of_element_located((By.NAME, 'component.From')))
        to_box = WebDriverWait(driver, 2).until(
            EC.visibility_of_element_located((By.NAME, 'component.To')))
        file_name.clear()
        file_name.send_keys('unfiltered_{}-{}'.format(start,end))
        from_box.send_keys(start)
        to_box.send_keys(end)
        # click 'export'
        time.sleep(1)
        submit = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="exportDialogForm"]/div[2]/a[2]')))
        submit.click()
        time.sleep(1)
        # close down download popup
        close = WebDriverWait(driver, 1).until(
            EC.visibility_of_element_located((By.XPATH, '/ html / body / section[3] / div[6] / div[1] / img')))
        close.click()

        print(start, end)
        start, end = range(start, end)
        time.sleep(5)


def range(start,end):
#check web interface and get max numof row allowed to be exported 
    start = start +1730
    end = end + 1730
    return start, end


if __name__ == "__main__":
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": r"D:\fame\unfiltered",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
    })
    driver = webdriver.Chrome("D:/PyCharm projects/famebvd/chromedriver.exe", chrome_options=chrome_options)


    username = "ssssss"
    # replace with username
    password = "sssssss"
    # replace with password
    start_url= 'https://libguides.kcl.ac.uk/az.php?a=f'
    start = 1
    end = 262962

    login(username, password)
    export_set(start, end)







