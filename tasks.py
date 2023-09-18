from RPA.Browser.Selenium import Selenium, By
from selenium.webdriver.common.keys import Keys
import sqlite3
import time
from RPA.Excel.Files import Files
from robocorp.tasks import task

def read_excel_worksheet(path, worksheet):
    lib = Files()
    lib.open_workbook(path)
    try:
        return lib.read_worksheet_as_table(worksheet, header=True)
    finally:
        lib.close_workbook()

@task
def main():
    
    excel_data = read_excel_worksheet("namedata.xlsx", "Sheet1")
    browser = Selenium()
    browser.auto_close = False
    browser.open_available_browser("https://meroshare.cdsc.com.np/#/login", maximized=True)


    for row in excel_data:
        person_name = row["Name"]
        Bank_name = row["Bank"]
        user_name = row["username"]
        pass_id = row["pass"]

        crn_num = row["CRN"]
        pin_num = row["PIN"]
        
        browser.wait_until_element_is_visible('//span[@class="select2-selection__rendered"]')

        (browser.find_element('//span[@class="select2-selection__rendered"]')).click()
        
         
        browser.input_text('//input[@class="select2-search__field"]', Bank_name)
        browser.press_keys(None, "ENTER")
        
        browser.press_keys(None, "TAB")
        
        # browser.wait_until_element_is_visible('//div[@class="input-group"]/input[@name="username"]')
        # button = browser.find_element('//div[@class="input-group"]/input[@name="username"]')
        # button.click() 

        browser.input_text('//div[@class="input-group"]/input[@name="username"]', user_name)
        browser.press_keys(None, "TAB")

        browser.input_text('//div[@class="input-group"]/input[@name="password"]', pass_id)
        browser.press_keys(None, "RETURN")
        # time.sleep(2)
        # (browser.find_element('//i[@class="mdi mdi-menu"]')).click()
        # (browser.find_element('//i[@class="mdi mdi-menu"]')).click()


        #myASBA Click
        
        browser.wait_until_element_is_visible('//*[@id="sideBar"]/nav/ul/li[8]/a/span')

        (browser.find_element('//*[@id="sideBar"]/nav/ul/li[8]/a/span')).click()

        # Apply for issue ma kaam garxa

        # (browser.find_element('//*[@id="main"]/div/app-asba/div/div[1]/div/div/ul/li[1]/a')).click()
        time.sleep(2)



        

        company_list = browser.find_elements(f'//div[@class="company-list"]//span[@class="share-of-type"]')
        for index, result in enumerate(company_list):
                IPO_name = browser.get_text(result)
                print(IPO_name)
                if IPO_name== "IPO" :
                    try:
                        ordinaryshare_name = browser.get_text(browser.find_element(f'//div[@class="company-list"][{index + 1}]//span[@tooltip="Share Group"]'))
                        print(ordinaryshare_name)
                        if ordinaryshare_name=="Ordinary Shares":
                            (browser.find_element(f'//div[@class="company-list"][{index + 1}]//button[@class="btn-issue"]/i[text()="Apply"]')).click()
                            time.sleep(1)
                            # to select the bank

                            (browser.find_element(f'//*[@id="selectBank"]')).click()
                            time.sleep(1)
                            #   to find min kitta quantity
                            (browser.find_element(f'//*[@id="selectBank"]/option[2]')).click()
                            minimum_quantity=browser.get_text(browser.find_element(f'//div[@class="section-block"]//label[text()="Minimum Quantity"]/following-sibling::div/span'))
                            print(minimum_quantity)
                            
                            (browser.find_element(f'//*[@id="appliedKitta"]')).click()
                            browser.press_keys(None, minimum_quantity)

                            #crn number
                            (browser.find_element(f'//*[@id="crnNumber"]')).click()
                            browser.press_keys(None, crn_num)
                            # i accept button
                            (browser.find_element(f'//*[@id="disclaimer"]')).click()
                            # subbmit
                            (browser.find_element(f'//button[@class="btn btn-gap btn-primary"]//span[text()="Proceed"]')).click()
                            #PIN COde
                            (browser.find_element(f'//*[@id="transactionPIN"]')).click()
                            browser.press_keys(None, str(pin_num))
                            # Apply FInal
                            (browser.find_element(f'//button[@class="btn btn-gap btn-primary"]//span[text()="Apply "]')).click()














                    except:
                        messsage="sorry no ordinary IPO share found" 
                        print(messsage)   



                else: 
                    message_reply="NO any IPO FOUND Currently"
                    print(message_reply)





        # logout
        (browser.find_element('//i[@class="msi msi-logout header-menu__icon"]')).click()        


    

        
      
         
if __name__ == "__main__":
    main()
