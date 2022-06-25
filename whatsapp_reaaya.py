# -*- coding: utf-8 -*-
"""
Created on Thu Jun 16 00:37:38 2022

@author: MAH
"""
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import re
import pandas as pd
from IPython.display import display
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
import socket
import os
from time import sleep


def convpdf(pdf):    
    imgs_path = []
    images = convert_from_path(pdf)
    for i, image in enumerate(images):
        fname = 'مضحي استمارة '+str(i+1)+'.png'
        path_img = "C:/Users/User/Desktop/whatsapp API/Istemara_images/" + fname
        imgs_path.append(path_img)
        image.save(path_img, "PNG")
    return imgs_path
        

images_list = convpdf("Adahi_For_WhatsApp.pdf")
        
        
        
        
def getphn(imgpath):
    path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    image_path = imgpath
      
    # Opening the image & storing it in an image object
    img = Image.open(image_path)
      
    # Providing the tesseract executable
    # location to pytesseract library
    pytesseract.tesseract_cmd = path_to_tesseract
      
    # Passing the image object to image_to_string() function
    # This function will extract the text from the image
    text = pytesseract.image_to_string(img, lang="ara")
 
    before ="خلوي: "
    after = " المستلم"
    regex = re.compile('{}(.*?){}'.format(re.escape(before), re.escape(after)))
    res = regex.findall(text) 
    if len(res) == 0:
        return None
    elif len(res[0][0]) > 8:
        return None        
    return res[0][0:10]


def element_presence(by,xpath,time):
    element_present = EC.presence_of_element_located((By.XPATH, xpath))
    WebDriverWait(driver, time).until(element_present)

def is_connected():
    try:
        socket.create_connection(("www.google.com", 80))
        return True
    except :
        is_connected()



def send_img(num,filepath):
    driver.get("https://web.whatsapp.com/send?phone={}&source=&data=#".format(num))
    try:
        driver.switch_to_alert().accept()
    except Exception as e:
       print('Error',e)  
    try:
        element_presence(By.XPATH,'//div[@title="Attach"]',30)
        attachment = driver.find_element(By.XPATH, '//div[@title="Attach"]')
        attachment.click()
        sleep(3)

        img_vid = driver.find_element(By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
        img_vid.send_keys(filepath)

        send = driver.find_element(By.XPATH,'//span[@data-testid="send"]')
        send.click()
    except Exception as e:
        print("Not Found: ",e)




#Creating dataframe to include images paths and phone number
excel_image_list = [] 
phone_list = []
for p in images_list:
    phone_number = getphn(p)
    if phone_number == None:                                                                                                                          
        print(phone_number, "should be NONE")
    else:
        phone_list.append("961" + phone_number)
        excel_image_list.append(p)
        print(phone_number)

excel_dict = {'Image': excel_image_list,
              'Phone': phone_list} 
    
df = pd.DataFrame(excel_dict)
display(df)
  
df2 = {'Image': excel_image_list[0], 'Phone': phone_list[0]}
df = df.append(df2, ignore_index = True)
df3 = {'Image': excel_image_list[1], 'Phone': phone_list[1]}
df = df.append(df3, ignore_index = True)

display(df)


#openning driver
driver_loc = 'chromedriver.exe'
os.environ["webdriver.chrome.driver"] = driver_loc
driver = webdriver.Chrome(driver_loc)
driver.get("http://web.whatsapp.com")
driver.maximize_window()
sleep(10)
driver.implicitly_wait(10)

# To send text/imgage & video/ document just change the function name and its parameter  
for index, row in df.iterrows():
    print(row['Image'], row['Phone'])
    try:
        send_img(row['Phone'],row['Image']) # Change here
        sleep(17)
    except Exception as e:
        print(Exception)
        sleep(10)
        is_connected()










#df2 = {images_list[0] : df[0][0]}

#df = df.append(df2, ignore_index = True)
#print (df) 
#df.to_excel('whatsapp_adahi_excel.xlsx')
    

        
