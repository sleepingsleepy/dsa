# -*- coding: utf-8 -*-
# WINDOW:
# autogui 

import pyautogui as pg
import clipboard as clp
import pandas as pd
import time

def find_image(image_path, max_attempts=20, retry_interval=2):  
    attempts = 0  
    while attempts < max_attempts:  
        try:  
            center =pg.locateCenterOnScreen(image_path)  
            if center is not None:  
                return center  # 如果找到图像，返回中心坐标  
        except pg.ImageNotFoundException:  
            print("Image not found. Retrying...")  
            time.sleep(retry_interval)  
            attempts += 1  
    return None  # 如果超过最大尝试次数仍未找到图像，则返回 None  

pg.PAUSE = 1
loc_tabnine = 123, 897 #the x and y coordinate of the chat box
time.sleep(5)

print('loop start')

    #loc = pg.locateCenterOnScreen('down_arrow.png',
    #                              grayscale=False, confidence=0.8)
loc = pg.locateCenterOnScreen('new_conversation.png',
                                grayscale=False, confidence=0.8)
    # find the clear history button

df = pd.read_excel('refined_code.xlsx', header=None)  
text_to_type = df[2].values.tolist() #the column of the code   
 
i = 0 
for line in text_to_type:  
    
            pg.click(loc)# start a new conversation
           
            #loc1 = 1140, 297
            
            loc1 = 1250, 295
            
            
            
            if line != None:
                pg.press('tab')
                pg.press('f2') #enter the editor in excel 
                pg.hotkey('ctrl', 'a')
                pg.hotkey('ctrl', 'c') #copy the code
                pg.click(loc_tabnine)
                pg.hotkey('ctrl', 'v')
                pg.press('enter')#enter the code to tabnine
                         
            #wait
            
            pg.click(loc1)
            
            pg.press('f2') 
            pg.hotkey('ctrl', 'a')
            pg.hotkey('ctrl', 'c') 
            
            pg.press('tab', 2)
            loc2 = pg.position()
            
            pg.click(loc_tabnine)#move to tabnine box
            pg.hotkey('ctrl', 'v')                
            pg.press('enter')#enter the word
            
            image_path = 'copy.png'  
            image_center = find_image(image_path)
            j = 0
            
            if(image_center != None):
                  
                pg.hotkey('ctrl', 'c') #if find no image center, then quit the program
            
            #wait to generate answer
        
            loc3 = pg.locateCenterOnScreen('copy.png',
                                        grayscale=False, confidence=0.9)
            pg.click(loc3)#copy the answer
            time.sleep(0.5)
        
            pg.click(loc2)
            pg.press('f2')
            pg.hotkey('ctrl', 'v')                
            pg.press('enter')
            
            pg.press('left', 2)
            
            i += 1
            print(i)
        
        