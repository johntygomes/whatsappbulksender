import time
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
#https://web.whatsapp.com/send?phone=++91+87270+67403
import openpyxl
import xlrd

my_shhet_index=1
my_main_file="collectednumbers.xlsx"

loc = my_main_file
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(my_shhet_index)

file=my_main_file
wb = openpyxl.load_workbook(filename=file)
ws = wb.worksheets[my_shhet_index]

driver=webdriver.Chrome()
driver.get("https://web.whatsapp.com/")
action = ActionChains(driver)
time.sleep(25)                     
#grup=driver.find_element_by_xpath('/html/body/div[1]/div/div/div[3]/div/div[2]/div[1]/div/div/div[1]/div/div ')
#grup.click()
time.sleep(2)

startrow=699
noofrows=800
for i in range(startrow,noofrows):
    current_number=sheet.cell_value(i,0)
    try:
        my_number_link="https://api.whatsapp.com/send?phone="+current_number
        driver.get(my_number_link)
        time.sleep(10)
        #msg_btn_wa_per=driver.find_element_by_link_text("MESSAGE")
        #msg_btn_wa_per.click()
        continue_to_chat_btn=driver.find_element_by_link_text("CONTINUE TO CHAT")
        continue_to_chat_btn.click()
        time.sleep(10)
        msg_btn_wa_per=driver.find_element_by_link_text("use WhatsApp Web")
        msg_btn_wa_per.click()    
        time.sleep(25)
        main_message_1="Hii Bro\u2763."
        ipbox=driver.find_element_by_xpath("/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[2]/div/div[2]")
        ipbox.send_keys(main_message_1)
        sendbtn=driver.find_element_by_xpath("/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[3]/button/span")
        sendbtn.click()
        #main_message_1="https://www.youtube.com/channel/UCPefCNVJIWojIXkujKFGE3g"
        #ipbox=driver.find_element_by_xpath("/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[2]/div/div[2]")
        #ipbox.send_keys(main_message_1)
        #sendbtn=driver.find_element_by_xpath("/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[3]/button/span")
        #sendbtn.click()        
        main_message_2="How Are You\u2764"
        ipbox=driver.find_element_by_xpath("/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[2]/div/div[2]")
        ipbox.send_keys(main_message_2)
        sendbtn=driver.find_element_by_xpath("/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[3]/button/span")
        sendbtn.click()
        time.sleep(10)
        ws.cell(row=i+1, column=3, value="SENT")
    except Exception as e:
        print(e)  
wb.save(file)
