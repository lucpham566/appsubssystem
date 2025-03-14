## eBay_11ShopsVer_2.03#工事中より3_03

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import SessionNotCreatedException, WebDriverException
from selenium.webdriver.common.by import By


options = webdriver.ChromeOptions()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
import zipfile
import urllib.request                            # ライブラリを取り込む
import os
import time
import datetime

import sys
import random
from selenium.webdriver.common.keys import Keys
import tkinter as tk
import tkinter.ttk as ttk
import datetime
import threading
import tkinter.font as font
import psutil
from selenium.webdriver.common.action_chains import ActionChains
from ebaysdk.trading import Connection as Trading
from ebaysdk.exception import ConnectionError

import re
import requests
import win32com.client

import logging

# Cấu hình logging
logging.basicConfig(filename='app.log', level=logging.DEBUG)

xl = win32com.client.GetObject(Class="Excel.Application")
#driver.quit()

service = Service('ccm/chromedriver.exe')  # Đường dẫn tới chromedriver
driver = webdriver.Chrome(service=service, options=options)


f1 = open('ccm/Appid.uud', 'r')
Appid=f1.read()
f1.close()

f2 = open('ccm/Devid.uud', 'r')
Devid=f2.read()
f2.close()

f3 = open('ccm/Certid.uud', 'r')
Certid=f3.read()
f3.close()

f4 = open('ccm/Token.uud', 'r')
Token=f4.read()
f4.close()

def DlePrdct(pr_id):
    #print("DlePrdct")
    pr_id=int(pr_id)
    api = Trading(appid=Appid, devid=Devid, certid=Certid, token=Token,config_file=None)
    api.execute('EndFixedPriceItem', { "EndingReason":"LostOrBroken","ItemID":pr_id})

def ChngQuantity(pr_id,pr_num):
    #print("ChngQuantity")
    pr_id=int(pr_id)
    request_data = {'InventoryStatus':{'ItemID': pr_id, 'Quantity':pr_num}}
    api = Trading(appid=Appid, devid=Devid, certid=Certid, token=Token, config_file=None)
    api.execute('ReviseInventoryStatus', request_data)

def ChngPrice(pr_id,NewPrice):
    #print("ChngPrice")
    pr_id=int(pr_id)
    request_data = {'InventoryStatus':{'ItemID': pr_id, 'StartPrice':NewPrice}}#ChngPrice
    api = Trading(appid=Appid, devid=Devid, certid=Certid, token=Token, config_file=None)
    api.execute('ReviseInventoryStatus', request_data)
    
def ReCalPrice(get_price,intex_chrg):
    global PriceRange
    global RateDepPri
    global RateYD
    global JpnTaxRate
    global eBayChrgRate
    global tim_sh

    index_price=((get_price/(1+float(JpnTaxRate)))/RateYD+intex_chrg)*(1+float(eBayChrgRate))
    for j in range(len(PriceRange)):
        if PriceRange[j+1]>index_price:
            IndPrdRate=RateDepPri[j]
            print(IndPrdRate)
            break
    sel_price=((get_price/(1+float(JpnTaxRate)))/RateYD*(float(IndPrdRate))+intex_chrg)*(1+float(eBayChrgRate))
    return  sel_price

def fin():
    sys.exit()

def get_det_1(prd_url,r_nm):
    global euu
    
    global driver
    global tr_sh
    global zkt
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD

    driver.get(prd_url)#①
    time.sleep(euu)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    try:
        p_id=driver.find_elements(By.CLASS_NAME,"fs-c-productNumber__number")[0].text
    except:
        pass
    try:  
        p_name=driver.find_elements(By.CLASS_NAME,"fs-c-productNameHeading__name")[0].text
    except:
        pass
   
    try:
        tex_yn=""
        tex_yn=driver.find_elements(By.CLASS_NAME,"fs-c-productNotice--outOfStock")[0].text
        #print(tex_yn)
    except:
        pass    
    if "在庫がございません" in tex_yn:
        sel_not="out of stock"
    else:
        sel_not="in stock"

    try:  
        elm_det=driver.find_elements(By.TAG_NAME,"table")[0].find_elements(By.TAG_NAME,"tr")
    except:
        pass
    try:
        for n in range(len(elm_det)-1):
            #print("TD")
            #print(elm_det[n+1].find_elements(By.TAG_NAME,"td")[0].text)
            p_det= p_det+elm_det[n+1].find_elements(By.TAG_NAME,"td")[0].text+" : "+elm_det[n+1].find_elements(By.TAG_NAME,"td")[1].text+"<br>"
    except:
        pass
    try:
        p_price=driver.find_elements(By.CLASS_NAME,"fs-c-price__value")[0].text
        p_price=p_price.replace(',', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements(By.CLASS_NAME,"fs-c-productThumbnail")[0].find_elements(By.TAG_NAME,"img")
    except:
        pass
    try:
        for m in range(len( elm_pic)):
             pic_url=pic_url+elm_pic[m].get_attribute("src")+"|"
        pic_url=pic_url.replace('-xs', '-l')
    except:
        pass    
    if p_name=="":
        sel_not="out of stock"
    
    if p_name!="":
        tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    if p_det!="":
        tr_sh.Cells(r_nm,6).Value= p_det
    if p_price!="":
        tr_sh.Cells(r_nm,7).Value= p_price
    if pic_url!="":
        tr_sh.Cells(r_nm,8).Value=pic_url

    if tr_sh.Cells(r_nm,3).Value is not None:
        if sel_not=="in stock":
            # if abs(float(tr_sh.Cells(r_nm,9).Value)-float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm,3).Value),int(tr_sh.Cells(r_nm,16).Value))###価格変更
            #         tr_sh.Cells(r_nm,9).Value=NgrYD###価格変更時為替レート記入
            #     except:
            #         pass
                
            if tr_sh.Cells(r_nm,11).Value is not None:###販売停止中（timestumpが存在）であれば,販売停止timestumpを削除して在庫を１に
                
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,1)
                    tr_sh.Cells(r_nm,11).Value=""
                except:
                    pass
                
        elif sel_not=="out of stock":
            z_range=0
            try:
                z_range=time.time()-tr_sh.Cells(r_nm,11).Value
            except:
                pass
            
            if tr_sh.Cells(r_nm,11).Value is not None and zkt<z_range:###販売停止中で（timestumpが存在して）タイマー経過していれば
                #削除して商品IO削除、販売停止timestumpを削除
                try:
                    DlePrdct(tr_sh.Cells(r_nm,3).Value)
                    tr_sh.Cells(r_nm,3).Value =""
                    tr_sh.Cells(r_nm,11).Value =""
                except:
                    pass
                
            elif tr_sh.Cells(r_nm,11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,0)
                    tr_sh.Cells(r_nm,11).Value =time.time()
                except:
                    pass
    
    
    
def get_det_3(prd_url,r_nm):
    global driver
    global tr_sh
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    try:
        p_id=driver.find_elements(By.CLASS_NAME,"")[0].text
    except:
        pass
    try:
        p_name=driver.find_elements(By.CLASS_NAME,"product-name")[0].text
    except:
        pass
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements(By.CLASS_NAME,"detail-sold-out")[0].text
    except:
        pass
    
    if tex_yn=="SOLD OUT":
        sel_not="out of stock"
    else:
        sel_not="in stock"
        
    try:
        p_det=driver.find_elements(By.CLASS_NAME,"description-text")[1].text
    except:
        pass
    try:    
        p_price=driver.find_elements(By.CLASS_NAME,"product-price-block")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('￥', '')
        p_price=p_price.replace('（税込）', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements(By.CLASS_NAME,"item-image-small")[0].find_elements(By.TAG_NAME,"img")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("src")+"|"
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
    
    if p_name!="":
        tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    if p_det!="":
        tr_sh.Cells(r_nm,6).Value= p_det
    if p_price!="":
        tr_sh.Cells(r_nm,7).Value= p_price
    if pic_url!="":
        tr_sh.Cells(r_nm,8).Value=pic_url

    if tr_sh.Cells(r_nm,3).Value is not None:
        if sel_not=="in stock":
            # if abs(float(tr_sh.Cells(r_nm,9).Value)-float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm,3).Value),int(tr_sh.Cells(r_nm,16).Value))###価格変更
            #         tr_sh.Cells(r_nm,9).Value=NgrYD###価格変更時為替レート記入
            #     except:
            #         pass
                
            if tr_sh.Cells(r_nm,11).Value is not None:###販売停止中（timestumpが存在）であれば,販売停止timestumpを削除して在庫を１に
                
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,1)
                    tr_sh.Cells(r_nm,11).Value=""
                except:
                    pass
                
        elif sel_not=="out of stock":
            z_range=0
            try:
                z_range=time.time()-tr_sh.Cells(r_nm,11).Value
            except:
                pass
            
            if tr_sh.Cells(r_nm,11).Value is not None and zkt<z_range:###販売停止中で（timestumpが存在して）タイマー経過していれば
                #削除して商品IO削除、販売停止timestumpを削除
                try:
                    DlePrdct(tr_sh.Cells(r_nm,3).Value)
                    tr_sh.Cells(r_nm,3).Value =""
                    tr_sh.Cells(r_nm,11).Value =""
                except:
                    pass
                
            elif tr_sh.Cells(r_nm,11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,0)
                    tr_sh.Cells(r_nm,11).Value =time.time()
                except:
                    pass

def get_det_5(prd_url,r_nm):
    global driver
    global tr_sh
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)
    #driver.get("https://takayama78online.jp/shop/g/g1334026530012/")#⑤
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""

    try:
        p_id=driver.find_element_by_id("spec_code_").find_elements(By.TAG_NAME,"dd")[0].text
    except:
        pass
    try:
        p_name=driver.find_elements(By.CLASS_NAME,"goods_name_")[0].text
    except:
        pass
        #document.getElementById("spec_stock_")
    
    zaik_an="aa"
    try:
        zaik_an=driver.find_elements(By.CLASS_NAME,"mainspec_")[0].find_elements(By.CLASS_NAME,"stock_")[0].text
    except:
        pass
    
    if "在庫あり" in zaik_an:
        sel_not="in stock"
    else:
        sel_not="out of stock"

    try:
        elm_det=driver.find_elements(By.TAG_NAME,"table")[1].find_elements(By.TAG_NAME,"tr")
        for n in range(len(elm_det)):
            p_det= p_det+elm_det[n].find_elements(By.TAG_NAME,"td")[0].text+" : "+elm_det[n].find_elements(By.TAG_NAME,"td")[1].text+"<br>"
    except:
        pass
    try:
        p_price=driver.find_elements(By.CLASS_NAME,"price_")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('円', '')
        p_price=int(p_price)

    except:
        pass
    try:
        elm_pic=driver.find_elements(By.CLASS_NAME,"goodsimg_")[0].find_elements(By.TAG_NAME,"a")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("href")+"|"           
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
    
    if p_name!="":
        tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    if p_det!="":
        tr_sh.Cells(r_nm,6).Value= p_det
    if p_price!="":
        tr_sh.Cells(r_nm,7).Value= p_price
    if pic_url!="":
        tr_sh.Cells(r_nm,8).Value=pic_url

    if tr_sh.Cells(r_nm,3).Value is not None:
        if sel_not=="in stock":
            # if abs(float(tr_sh.Cells(r_nm,9).Value)-float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm,3).Value),int(tr_sh.Cells(r_nm,16).Value))###価格変更
            #         tr_sh.Cells(r_nm,9).Value=NgrYD###価格変更時為替レート記入
            #     except:
            #         pass
                
            if tr_sh.Cells(r_nm,11).Value is not None:###販売停止中（timestumpが存在）であれば,販売停止timestumpを削除して在庫を１に
                
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,1)
                    tr_sh.Cells(r_nm,11).Value=""
                except:
                    pass
                
        elif sel_not=="out of stock":
            z_range=0
            try:
                z_range=time.time()-tr_sh.Cells(r_nm,11).Value
            except:
                pass
            
            if tr_sh.Cells(r_nm,11).Value is not None and zkt<z_range:###販売停止中で（timestumpが存在して）タイマー経過していれば
                #削除して商品IO削除、販売停止timestumpを削除
                try:
                    DlePrdct(tr_sh.Cells(r_nm,3).Value)
                    tr_sh.Cells(r_nm,3).Value =""
                    tr_sh.Cells(r_nm,11).Value =""
                except:
                    pass
                
            elif tr_sh.Cells(r_nm,11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,0)
                    tr_sh.Cells(r_nm,11).Value =time.time()
                except:
                    pass

def get_det_7(prd_url,r_nm):
    global driver
    global tr_sh
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)   
    #driver.get("https://kanteikyoku-web.jp/shop/products/detail/82246")#⑦
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    try:
        p_id=driver.find_elements(By.CLASS_NAME,"")[0].text
    except:
        pass
    try:
        p_name=driver.find_elements(By.CLASS_NAME,"ec-headingTitle")[0].text
    except:
        pass   
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements(By.CLASS_NAME,"add-cart")[0].text
    except:
        pass    
    if "カートに追加" in tex_yn:
        sel_not="in stock"
    else:
        sel_not="out of stock"

    try:
        elm_det=driver.find_elements(By.CLASS_NAME,"ec-productRole__description")[0].find_elements(By.TAG_NAME,"tr")
        for n in range(len(elm_det)-1):
            if elm_det[n+1].find_elements(By.TAG_NAME,"th")[0].text=="管理番号":
                p_id=elm_det[n+1].find_elements(By.TAG_NAME,"td")[0].text
            p_det= p_det+elm_det[n+1].find_elements(By.TAG_NAME,"th")[0].text+" : "+elm_det[n+1].find_elements(By.TAG_NAME,"td")[0].text+"<br>"
    except:
        pass
    try:
        p_price=driver.find_elements(By.CLASS_NAME,"ec-price__price")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('￥', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements(By.CLASS_NAME,"item_nav")[0].find_elements(By.TAG_NAME,"img")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("src")+"|"
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
    
    if p_name!="":
        tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    if p_det!="":
        tr_sh.Cells(r_nm,6).Value= p_det
    if p_price!="":
        tr_sh.Cells(r_nm,7).Value= p_price
    if pic_url!="":
        tr_sh.Cells(r_nm,8).Value=pic_url

    if tr_sh.Cells(r_nm,3).Value is not None:
        if sel_not=="in stock":
            # if abs(float(tr_sh.Cells(r_nm,9).Value)-float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm,3).Value),int(tr_sh.Cells(r_nm,16).Value))###価格変更
            #         tr_sh.Cells(r_nm,9).Value=NgrYD###価格変更時為替レート記入
            #     except:
            #         pass
                
            if tr_sh.Cells(r_nm,11).Value is not None:###販売停止中（timestumpが存在）であれば,販売停止timestumpを削除して在庫を１に
                
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,1)
                    tr_sh.Cells(r_nm,11).Value=""
                except:
                    pass
                
        elif sel_not=="out of stock":
            z_range=0
            try:
                z_range=time.time()-tr_sh.Cells(r_nm,11).Value
            except:
                pass
            
            if tr_sh.Cells(r_nm,11).Value is not None and zkt<z_range:###販売停止中で（timestumpが存在して）タイマー経過していれば
                #削除して商品IO削除、販売停止timestumpを削除
                try:
                    DlePrdct(tr_sh.Cells(r_nm,3).Value)
                    tr_sh.Cells(r_nm,3).Value =""
                    tr_sh.Cells(r_nm,11).Value =""
                except:
                    pass
                
            elif tr_sh.Cells(r_nm,11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,0)
                    tr_sh.Cells(r_nm,11).Value =time.time()
                except:
                    pass
    


def get_det_9(prd_url,r_nm):
    global driver
    global tr_sh
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)   
    #⑨eLADY 
    time.sleep(8)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    
    try:
        p_name=driver.find_elements(By.CLASS_NAME,"detail__name")[0].text
    except:
        pass   
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements(By.CLASS_NAME,"detail__btn_list--cart")[0].text
    except:
        pass    
    if "カートに入れる" in tex_yn:
        sel_not="in stock"
    else:
        sel_not="out of stock"

    try:
        elm_det=driver.find_element_by_id("info1_content").find_elements(By.CLASS_NAME,"y")
        for n in range(len(elm_det)):
            p_det=p_det+elm_det[n].text+": "
            p_det= p_det+driver.find_element_by_id("info1_content").find_elements(By.CLASS_NAME,"x")[n].text+"<br>"
    except:
        pass
    
    try:
        driver.find_elements(By.CLASS_NAME,"pc_view_label")[1].click()
        time.sleep(1)
        elm_det=driver.find_element_by_id("info2_content").find_elements(By.CLASS_NAME,"y")
        for n in range(len(elm_det)):
            p_det=p_det+elm_det[n].text+": "
            p_det= p_det+driver.find_element_by_id("info2_content").find_elements(By.CLASS_NAME,"x")[n].text+"<br>"
    except:
        pass

    try:
        p_price=driver.find_elements(By.CLASS_NAME,"price")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('¥', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements(By.CLASS_NAME,"product-image-thumbs")[0].find_elements(By.TAG_NAME,"img")
        for m in range(len( elm_pic)):
            pic_url=pic_url+(elm_pic[m].get_attribute("src")).replace('/250', '/1000')+"|"
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
    
    if p_name!="":
        tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    if p_det!="":
        tr_sh.Cells(r_nm,6).Value= p_det
    if p_price!="":
        tr_sh.Cells(r_nm,7).Value= p_price
    if pic_url!="":
        tr_sh.Cells(r_nm,8).Value=pic_url

    if tr_sh.Cells(r_nm,3).Value is not None:
        if sel_not=="in stock":
            # if abs(float(tr_sh.Cells(r_nm,9).Value)-float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm,3).Value),int(tr_sh.Cells(r_nm,16).Value))###価格変更
            #         tr_sh.Cells(r_nm,9).Value=NgrYD###価格変更時為替レート記入
            #     except:
            #         pass
                
            if tr_sh.Cells(r_nm,11).Value is not None:###販売停止中（timestumpが存在）であれば,販売停止timestumpを削除して在庫を１に
                
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,1)
                    tr_sh.Cells(r_nm,11).Value=""
                except:
                    pass
                
        elif sel_not=="out of stock":
            z_range=0
            try:
                z_range=time.time()-tr_sh.Cells(r_nm,11).Value
            except:
                pass
            
            if tr_sh.Cells(r_nm,11).Value is not None and zkt<z_range:###販売停止中で（timestumpが存在して）タイマー経過していれば
                #削除して商品IO削除、販売停止timestumpを削除
                try:
                    DlePrdct(tr_sh.Cells(r_nm,3).Value)
                    tr_sh.Cells(r_nm,3).Value =""
                    tr_sh.Cells(r_nm,11).Value =""
                except:
                    pass
                
            elif tr_sh.Cells(r_nm,11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,0)
                    tr_sh.Cells(r_nm,11).Value =time.time()
                except:
                    pass

def get_det_11(prd_url,r_nm):
    global driver
    global tr_sh
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)   
    #⑪大黒屋
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    
    try:
        p_name=driver.find_elements(By.TAG_NAME,"h1")[0].text
    except:
        pass   
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements(By.CLASS_NAME,"btn-primary")[1].text
    except:
        pass    
    if "カートに入れる" in tex_yn:
        sel_not="in stock"
    else:
        sel_not="out of stock"

    try:
        driver.find_elements(By.CLASS_NAME,"mat-expansion-panel-header-title")[0].click()
        time.sleep(1)
        elm_det=driver.find_elements(By.CLASS_NAME,"pt-4")[1].find_elements(By.TAG_NAME,"div")
        for n in range(int(len(elm_det)/3)):
            p_det= p_det+elm_det[3*n+1].text+" : "+elm_det[3*n+2].text+"<br>"
    except:
        pass
    try:
        p_det= p_det+driver.find_elements(By.CLASS_NAME,"pt-2")[1].get_attribute('innerHTML')
    except:
        pass
    
    
    try:
        p_price=driver.find_elements(By.CLASS_NAME,"my-2")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('¥', '')
        p_price=p_price.replace(' ', '')
        p_price=p_price.replace('(税込)', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements(By.CLASS_NAME,"swiper-container")[1].find_elements(By.TAG_NAME,"img")
        for m in range(len( elm_pic)):
            pic_url=(pic_url+elm_pic[m].get_attribute("src")+"|").replace('-thumb', '')
            
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
    
    if p_name!="":
        tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    if p_det!="":
        tr_sh.Cells(r_nm,6).Value= p_det
    if p_price!="":
        tr_sh.Cells(r_nm,7).Value= p_price
    if pic_url!="":
        tr_sh.Cells(r_nm,8).Value=pic_url

    if tr_sh.Cells(r_nm,3).Value is not None:
        if sel_not=="in stock":
            # if abs(float(tr_sh.Cells(r_nm,9).Value)-float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm,3).Value),int(tr_sh.Cells(r_nm,16).Value))###価格変更
            #         tr_sh.Cells(r_nm,9).Value=NgrYD###価格変更時為替レート記入
            #     except:
            #         pass
                
            if tr_sh.Cells(r_nm,11).Value is not None:###販売停止中（timestumpが存在）であれば,販売停止timestumpを削除して在庫を１に
                
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,1)
                    tr_sh.Cells(r_nm,11).Value=""
                except:
                    pass
                
        elif sel_not=="out of stock":
            z_range=0
            try:
                z_range=time.time()-tr_sh.Cells(r_nm,11).Value
            except:
                pass
            
            if tr_sh.Cells(r_nm,11).Value is not None and zkt<z_range:###販売停止中で（timestumpが存在して）タイマー経過していれば
                #削除して商品IO削除、販売停止timestumpを削除
                try:
                    DlePrdct(tr_sh.Cells(r_nm,3).Value)
                    tr_sh.Cells(r_nm,3).Value =""
                    tr_sh.Cells(r_nm,11).Value =""
                except:
                    pass
                
            elif tr_sh.Cells(r_nm,11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,0)
                    tr_sh.Cells(r_nm,11).Value =time.time()
                except:
                    pass

def get_det_15(prd_url,r_nm):
    global driver
    global tr_sh
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)
    #⑮ヤフオク
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""    
    try:
        p_name=driver.find_elements(By.CLASS_NAME,"ProductTitle__text")[0].text###OK
    except:
        pass   

    try:
        tex_yn=""
        tex_yn=driver.find_elements(By.CLASS_NAME,"ClosedHeader__tag")[0].text
    except:
        pass    
    if "このオークションは終了" in tex_yn:
        sel_not="out of stock"   
    else:
        sel_not="in stock"

    try:
        if p_det=="":
            #document.getElementsByClassName("ProductExplanation__commentBody")
            #p_det=driver.find_elements(By.CLASS_NAME,"ProductExplanation__commentBody")[0].get_attribute('outerHTML')#OKKKKK
            p_det=driver.find_elements(By.CLASS_NAME,"ProductExplanation__commentBody")[0].get_attribute('outerText')
            #p_det=driver.find_elements(By.CLASS_NAME,"ProductExplanation__commentBody")[0].find_elements(By.TAG_NAME,"div")[0].get_attribute('outerHTML')#A1
            #p_det=driver.find_elements(By.CLASS_NAME,"ProductExplanation__commentBody")[0].find_elements(By.TAG_NAME,"div")[0].get_attribute('outerText')
            
            #document.getElementsByClassName("ProductExplanation__commentArea")
            #p_det=driver.find_elements(By.CLASS_NAME,"ProductExplanation__commentArea")[0].get_attribute('innerHTML')
    except:
        pass
    
    try:
        if p_det=="":
            elm_det=driver.find_elements(By.TAG_NAME,"tbody")[1].find_elements(By.TAG_NAME,"tr")
            for n in range(len(elm_det)):
                try:
                    p_det= p_det+elm_det[n].find_elements(By.TAG_NAME,"td")[0].text+" : "+elm_det[n].find_elements(By.TAG_NAME,"td")[1].text+"<br>"
                except:
                    pass
    except:
        pass
    
    try:
        if p_det=="":
            p_det=driver.find_elements(By.TAG_NAME,"tbody")[1].get_attribute('innerHTML')
    except:
        pass
    
    try:
        p_price=driver.find_elements(By.CLASS_NAME,"Price__value")[0].text
    except:
        pass
    #document.getElementsByClassName("Price__tax")
    
    try:
        p_price.find('（')
        p_price= p_price[: p_price.find('（')]
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('¥', '')
        p_price=p_price.replace('円', '')
        p_price=int(p_price)
    except:
        pass
    
    #document.getElementsByClassName("Price__tax")
    if len(driver.find_elements(By.CLASS_NAME,"Price__tax"))!=0 and driver.find_elements(By.CLASS_NAME,"Price__tax")[0].text!="（税 0 円）":
        try:
            p_price=driver.find_elements(By.CLASS_NAME,"Price__tax")[0].text
            p_price=p_price.replace(',', '')
            p_price=p_price.replace('（税込 ', '')
            p_price=p_price.replace(' 円）', '')
            p_price=int(p_price)
            
        except:
            pass
    
    
    
    
    try:
        elm_pic=driver.find_elements(By.CLASS_NAME,"ProductImage__thumbnail")
        for m in range(len( elm_pic)):
            pind_url=elm_pic[m].find_elements(By.TAG_NAME,"img")[0].get_attribute("src")
            pind_url=pind_url[:pind_url.find('.jpg')+4]
            pic_url=pic_url+ pind_url+"|"
    except:
        pass

    if p_name=="":
        sel_not="out of stock"
    
    if p_name!="":
        tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    if p_det!="":
        tr_sh.Cells(r_nm,6).Value= p_det
    if p_price!="":
        tr_sh.Cells(r_nm,7).Value= p_price
    if pic_url!="":
        tr_sh.Cells(r_nm,8).Value=pic_url

    if tr_sh.Cells(r_nm,3).Value is not None:
        if sel_not=="in stock":
            # if abs(float(tr_sh.Cells(r_nm,9).Value)-float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm,3).Value),int(tr_sh.Cells(r_nm,16).Value))###価格変更
            #         tr_sh.Cells(r_nm,9).Value=NgrYD###価格変更時為替レート記入
            #     except:
            #         pass
                
            if tr_sh.Cells(r_nm,11).Value is not None:###販売停止中（timestumpが存在）であれば,販売停止timestumpを削除して在庫を１に
                
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,1)
                    tr_sh.Cells(r_nm,11).Value=""
                except:
                    pass
                
        elif sel_not=="out of stock":
            z_range=0
            try:
                z_range=time.time()-tr_sh.Cells(r_nm,11).Value
            except:
                pass
            
            if tr_sh.Cells(r_nm,11).Value is not None and zkt<z_range:###販売停止中で（timestumpが存在して）タイマー経過していれば
                #削除して商品IO削除、販売停止timestumpを削除
                try:
                    DlePrdct(tr_sh.Cells(r_nm,3).Value)
                    tr_sh.Cells(r_nm,3).Value =""
                    tr_sh.Cells(r_nm,11).Value =""
                except:
                    pass
                
            elif tr_sh.Cells(r_nm,11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,0)
                    tr_sh.Cells(r_nm,11).Value =time.time()
                except:
                    pass

def get_det_21(prd_url,r_nm):
    global driver
    global tr_sh
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)   
    #㉑firekids
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    
    try:
        #document.getElementsByClassName("product__title")
        p_name=driver.find_elements(By.CLASS_NAME,"product__title")[0].text
    except:
        pass   
    
    try:
        #document.getElementsByClassName("product-form__submit")
        tex_yn=""
        tex_yn=driver.find_elements(By.CLASS_NAME,"product-form__submit")[0].text
    except:
        pass    
    if "カートに追加する" in tex_yn:
        sel_not="in stock"
    else:
        sel_not="out of stock"
        
    try:
        #document.getElementsByClassName("custom-description-table")
        elm_det=driver.find_elements(By.CLASS_NAME,"custom-description-table")[0].find_elements(By.TAG_NAME,"tr")
        for n in range(len(elm_det)):
            p_det= p_det+elm_det[n].find_elements(By.TAG_NAME,"th")[0].text+" : "+elm_det[n].find_elements(By.TAG_NAME,"td")[0].text+"<br>"
        p_det= p_det.replace('<br>', '<?>')
    except:
        pass
    try:
        #document.getElementsByClassName("product__description")
        p_det= p_det+driver.find_elements(By.CLASS_NAME,"product__description")[0].text
    except:
        pass   
    
    try:
        #document.getElementsByClassName("price__regular")
        #https://firekids.jp/collections/all/products/sinn-%E3%83%9F%E3%83%AA%E3%82%BF%E3%83%AA%E3%83%BC%E3%82%AF%E3%83%AD%E3%83%8E-%E3%83%AC%E3%83%9E%E3%83%8B%E3%82%A25100
        p_price=driver.find_elements(By.CLASS_NAME,"price__regular")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('¥', '')
        p_price=p_price.replace('通常価格', '')
        p_price=p_price.replace('円', '')
        p_price=int(p_price)
    except:
        pass
    try:
        #document.getElementsByClassName("thumbnail-list")[0].getElementsByTagName("img")
        elm_pic=driver.find_elements(By.CLASS_NAME,"thumbnail-list")[0].find_elements(By.TAG_NAME,"img")
        for m in range(len( elm_pic)):
            
            pic_url=pic_url+elm_pic[m].get_attribute("src")+"|"
    except:
        pass
    
    if p_name=="":
        sel_not="out of stock"
    
    if p_name!="":
        tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    if p_det!="":
        tr_sh.Cells(r_nm,6).Value= p_det
    if p_price!="":
        tr_sh.Cells(r_nm,7).Value= p_price
    if pic_url!="":
        tr_sh.Cells(r_nm,8).Value=pic_url

    if tr_sh.Cells(r_nm,3).Value is not None:
        if sel_not=="in stock":
            # if abs(float(tr_sh.Cells(r_nm,9).Value)-float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm,3).Value),int(tr_sh.Cells(r_nm,16).Value))###価格変更
            #         tr_sh.Cells(r_nm,9).Value=NgrYD###価格変更時為替レート記入
            #     except:
            #         pass
                
            if tr_sh.Cells(r_nm,11).Value is not None:###販売停止中（timestumpが存在）であれば,販売停止timestumpを削除して在庫を１に
                
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,1)
                    tr_sh.Cells(r_nm,11).Value=""
                except:
                    pass
                
        elif sel_not=="out of stock":
            z_range=0
            try:
                z_range=time.time()-tr_sh.Cells(r_nm,11).Value
            except:
                pass
            
            if tr_sh.Cells(r_nm,11).Value is not None and zkt<z_range:###販売停止中で（timestumpが存在して）タイマー経過していれば
                #削除して商品IO削除、販売停止timestumpを削除
                try:
                    DlePrdct(tr_sh.Cells(r_nm,3).Value)
                    tr_sh.Cells(r_nm,3).Value =""
                    tr_sh.Cells(r_nm,11).Value =""
                except:
                    pass
                
            elif tr_sh.Cells(r_nm,11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,0)
                    tr_sh.Cells(r_nm,11).Value =time.time()
                except:
                    pass

def get_det_22(prd_url,r_nm):
    global driver
    global tr_s
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)   
    #㉒bigmoon
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price="SOLD"
    pic_url=""
    
    try:
        #document.getElementsByClassName("c-product_detail__title")
        p_name=driver.find_elements(By.CLASS_NAME,"c-product_detail__title")[1].text
    except:
        pass   
    
    try:
        #document.getElementsByClassName("product-form__cart-submit")
        tex_yn=""
        tex_yn=driver.find_elements(By.CLASS_NAME,"price")[0].text
    except:
        pass    
    if "SOLD" in tex_yn:
        sel_not="out of stock"
    else:
        sel_not="in stock"


    #document.getElementsByClassName("c-product_detail__bnrs")
    try:
        if len(driver.find_elements(By.CLASS_NAME,"c-product_detail__bnrs"))!=0:
            sel_not="out of stock"
    except:
        pass
    
    try:
        #ocument.getElementsByClassName("c-product_detail__table")[1].getElementsByTagName("tr")
        elm_det=driver.find_elements(By.CLASS_NAME,"c-product_detail__table")[1].find_elements(By.TAG_NAME,"tr")
        for n in range(len(elm_det)):
            p_det= p_det+elm_det[n].find_elements(By.TAG_NAME,"th")[0].text+" : "+elm_det[n].find_elements(By.TAG_NAME,"td")[0].text+"<br>"
        p_det= p_det.replace('<br>', '<?>')
    except:
        pass
        
    
    try:
        #document.getElementsByClassName("price")
        p_price=driver.find_elements(By.CLASS_NAME,"price")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('¥', '')
        p_price=p_price.replace('(税込)', '')
        p_price=p_price.replace('円', '')
        p_price=int(p_price)
    except:
        pass
    try:
        #document.getElementsByClassName("c-product_detail-images__thumb")[0].getElementsByTagName("img")
        elm_pic=driver.find_elements(By.CLASS_NAME,"c-product_detail-images__thumb")[0].find_elements(By.TAG_NAME,"img")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("src")+"|"
    except:
        pass
    
    if p_name=="":
        sel_not="out of stock"
    
    if p_name!="":
        tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    if p_det!="":
        tr_sh.Cells(r_nm,6).Value= p_det
    if p_price!="":
        tr_sh.Cells(r_nm,7).Value= p_price
    if pic_url!="":
        tr_sh.Cells(r_nm,8).Value=pic_url

    if tr_sh.Cells(r_nm,3).Value is not None:
        if sel_not=="in stock":
            # if abs(float(tr_sh.Cells(r_nm,9).Value)-float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm,3).Value),int(tr_sh.Cells(r_nm,16).Value))###価格変更
            #         tr_sh.Cells(r_nm,9).Value=NgrYD###価格変更時為替レート記入
            #     except:
            #         pass
                
            if tr_sh.Cells(r_nm,11).Value is not None:###販売停止中（timestumpが存在）であれば,販売停止timestumpを削除して在庫を１に
                
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,1)
                    tr_sh.Cells(r_nm,11).Value=""
                except:
                    pass
                
        elif sel_not=="out of stock":
            z_range=0
            try:
                z_range=time.time()-tr_sh.Cells(r_nm,11).Value
            except:
                pass
            
            if tr_sh.Cells(r_nm,11).Value is not None and zkt<z_range:###販売停止中で（timestumpが存在して）タイマー経過していれば
                #削除して商品IO削除、販売停止timestumpを削除
                try:
                    DlePrdct(tr_sh.Cells(r_nm,3).Value)
                    tr_sh.Cells(r_nm,3).Value =""
                    tr_sh.Cells(r_nm,11).Value =""
                except:
                    pass
                
            elif tr_sh.Cells(r_nm,11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,0)
                    tr_sh.Cells(r_nm,11).Value =time.time()
                except:
                    pass
    
       

def get_det_23(prd_url,r_nm):
    global driver
    global tr_sh
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)   
    #㉓ticken
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not="SOLD OUT"
    p_det=""
    p_price=""
    pic_url=""
    
    try:
        #document.getElementsByTagName("h1")
        p_name=driver.find_elements(By.TAG_NAME,"h1")[0].text
    except:
        pass   
    
    try:
        #document.getElementsByClassName("mar_t_0")
        #document.getElementsByClassName("table-bordered")[0].getElementsByTagName("td")[0]
        tex_yn="SOLD OUT"
        tex_yn=driver.find_elements(By.CLASS_NAME,"table-bordered")[0].find_elements(By.TAG_NAME,"td")[0].text
    except:
        pass  
    if "SOLD OUT" in tex_yn:
        sel_not="out of stock"
    
    else:
        sel_not="in stock"
        
    #document.getElementsByClassName("btn-lg")
    try:
        ex_btn=""
        ex_btn=driver.find_elements(By.CLASS_NAME,"btn-lg")[0].text
        if "取り置き" in ex_btn:
            sel_not="out of stock"
    except:
        pass

    try:
        #document.getElementsByClassName("product-order-exp")
        p_det= driver.find_elements(By.CLASS_NAME,"product-order-exp")[0].text
        p_det= p_det.replace('<br>', '<?>')
    except:
        pass
        
    try:
        #document.getElementsByClassName("table-bordered")[0].getElementsByTagName("td")[0]
        p_price=driver.find_elements(By.CLASS_NAME,"table-bordered")[0].find_elements(By.TAG_NAME,"td")[0].text
        try:
            p_price_del=driver.find_elements(By.CLASS_NAME,"table-bordered")[0].find_elements(By.TAG_NAME,"td")[0].find_elements(By.TAG_NAME,"s")[0].text
            p_price=p_price.replace(p_price_del,"")
        except:
            pass
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('￥', '')
        p_price=p_price.replace('(税込)', '')
        p_price=p_price.replace('円', '')
        p_price=int(p_price)
    except:
        pass
    try:
        #document.getElementById("bx-pager").getElementsByTagName("img")
        elm_pic=driver.find_element_by_id("bx-pager").find_elements(By.TAG_NAME,"img")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("src")+"|"
    except:
        pass
    if p_name=="":
        sel_not="out of stock"

    if p_name!="":
        tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    if p_det!="":
        tr_sh.Cells(r_nm,6).Value= p_det
    if p_price!="":
        tr_sh.Cells(r_nm,7).Value= p_price
    if pic_url!="":
        tr_sh.Cells(r_nm,8).Value=pic_url

    if tr_sh.Cells(r_nm,3).Value is not None:
        if sel_not=="in stock":
            # if abs(float(tr_sh.Cells(r_nm,9).Value)-float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm,3).Value),int(tr_sh.Cells(r_nm,16).Value))###価格変更
            #         tr_sh.Cells(r_nm,9).Value=NgrYD###価格変更時為替レート記入
            #     except:
            #         pass
                
            if tr_sh.Cells(r_nm,11).Value is not None:###販売停止中（timestumpが存在）であれば,販売停止timestumpを削除して在庫を１に
                
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,1)
                    tr_sh.Cells(r_nm,11).Value=""
                except:
                    pass
                
        elif sel_not=="out of stock":
            z_range=0
            try:
                z_range=time.time()-tr_sh.Cells(r_nm,11).Value
            except:
                pass
            
            if tr_sh.Cells(r_nm,11).Value is not None and zkt<z_range:###販売停止中で（timestumpが存在して）タイマー経過していれば
                #削除して商品IO削除、販売停止timestumpを削除
                try:
                    DlePrdct(tr_sh.Cells(r_nm,3).Value)
                    tr_sh.Cells(r_nm,3).Value =""
                    tr_sh.Cells(r_nm,11).Value =""
                except:
                    pass
                
            elif tr_sh.Cells(r_nm,11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,0)
                    tr_sh.Cells(r_nm,11).Value =time.time()
                except:
                    pass
    

def get_det_24(prd_url,r_nm):
    global driver
    global tr_sh
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)   
    #㉔komehyo
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    
    try:
        #document.getElementsByClassName("p-product-name")
        p_name=driver.find_elements(By.CLASS_NAME,"p-product-name")[0].text
    except:
        pass   
    
    try:
        #document.getElementsByClassName("p-link--button__txt")
        tex_yn=""
        tex_yn=driver.find_elements(By.CLASS_NAME,"p-link--button__txt")[0].text
    except:
        pass    
    if "ショッピングカートに入れる" in tex_yn:
        sel_not="in stock"
    else:
        sel_not="out of stock"
        
    try:
        #document.getElementsByClassName("p-table__content")
        #p_det=driver.find_elements(By.CLASS_NAME,"p-table__content")[0].text
        
        elm_det=driver.find_elements(By.CLASS_NAME,"p-table__content")[0].find_elements(By.TAG_NAME,"tr")
        for n in range(len(elm_det)):
            try:
                p_det= p_det+elm_det[n].find_elements(By.TAG_NAME,"th")[0].text+" : "+elm_det[n].find_elements(By.TAG_NAME,"td")[0].text+"<br>"
            except:
                pass
        #p_det= p_det.replace('<br>', '<?>')
        
    except:
        pass
        
    try:
        #document.getElementsByClassName("p-txt--07")
        p_price=driver.find_elements(By.CLASS_NAME,"p-txt--07")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('￥', '')
        p_price=p_price.replace('(税込)', '')
        p_price=p_price.replace('円', '')
        p_price=int(p_price)
    except:
        pass
    try:
        #document.getElementsByClassName("p-showcase__thumbs")[0].getElementsByTagName("img")
        elm_pic=driver.find_elements(By.CLASS_NAME,"p-showcase__thumbs")[0].find_elements(By.TAG_NAME,"img")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("src")+"|"
    except:
        pass
    if p_name=="":
        sel_not="out of stock"

    if p_name!="":
        tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    if p_det!="":
        tr_sh.Cells(r_nm,6).Value= p_det
    if p_price!="":
        tr_sh.Cells(r_nm,7).Value= p_price
    if pic_url!="":
        tr_sh.Cells(r_nm,8).Value=pic_url

    if tr_sh.Cells(r_nm,3).Value is not None:
        if sel_not=="in stock":
            # if abs(float(tr_sh.Cells(r_nm,9).Value)-float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm,3).Value),int(tr_sh.Cells(r_nm,16).Value))###価格変更
            #         tr_sh.Cells(r_nm,9).Value=NgrYD###価格変更時為替レート記入
            #     except:
            #         pass
                
            if tr_sh.Cells(r_nm,11).Value is not None:###販売停止中（timestumpが存在）であれば,販売停止timestumpを削除して在庫を１に
                
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,1)
                    tr_sh.Cells(r_nm,11).Value=""
                except:
                    pass
                
        elif sel_not=="out of stock":
            z_range=0
            try:
                z_range=time.time()-tr_sh.Cells(r_nm,11).Value
            except:
                pass
            
            if tr_sh.Cells(r_nm,11).Value is not None and zkt<z_range:###販売停止中で（timestumpが存在して）タイマー経過していれば
                #削除して商品IO削除、販売停止timestumpを削除
                try:
                    DlePrdct(tr_sh.Cells(r_nm,3).Value)
                    tr_sh.Cells(r_nm,3).Value =""
                    tr_sh.Cells(r_nm,11).Value =""
                except:
                    pass
                
            elif tr_sh.Cells(r_nm,11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm,3).Value,0)
                    tr_sh.Cells(r_nm,11).Value =time.time()
                except:
                    pass

#https://www.bigmoon-kyoto.com/   
def get_det_25 (prd_url,r_nm):

    print(f"Accessing URL: {prd_url}")  # In ra URL

    global driver
    global tr_sh
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)   

    time.sleep(3)
    p_id = ""
    p_name = ""
    sel_not = ""
    p_det = ""
    p_price = ""
    pic_url = ""
    tex_yn = ""  # Gán giá trị mặc định cho tex_yn

    # Log tên sản phẩm
    try:
        p_name = driver.find_elements(By.CLASS_NAME,"sale_single_tit")[0].text
    except Exception as e:
        print(f"Error while fetching product name: {e}")
    print(f"Product Name: {p_name}")  # Log tên sản phẩm

    # Log trạng thái sản phẩm
    try:
        tex_yn = driver.find_elements(By.CLASS_NAME,"p-link--button__txt")[0].text
    except Exception as e:
        print(f"Error while fetching stock status text: {e}")
    print(f"Stock Status Text: {tex_yn}")  # Log trạng thái sản phẩm

    if "ショッピングカートに入れる" in tex_yn:
        sel_not = "in stock"
        print("Status: in stock")  # Log trạng thái "in stock"
    else:
        sel_not = "out of stock"
        print("Status: out of stock")  # Log trạng thái "out of stock"
        
    # Log chi tiết sản phẩm
    try:
        elm_det = driver.find_elements(By.CLASS_NAME,"p-table__content")[0].find_elements(By.TAG_NAME,"tr")
        for n in range(len(elm_det)):
            try:
                p_det = p_det + elm_det[n].find_elements(By.TAG_NAME,"th")[0].text + " : " + elm_det[n].find_elements(By.TAG_NAME,"td")[0].text + "<br>"
            except:
                pass
    except Exception as e:
        print(f"Error while fetching product details: {e}")
    print(f"Product Details: {p_det}")  # Log chi tiết sản phẩm
        
    # Log giá sản phẩm
    try:
        p_price = driver.find_elements(By.CLASS_NAME,"sale_price_b")[0].text
        p_price = p_price.replace(',', '')
        p_price = p_price.replace('￥', '')
        p_price = p_price.replace('(税込)', '')
        p_price = p_price.replace('円', '')
        p_price = int(p_price)
    except Exception as e:
        print(f"Error while fetching product price: {e}")
    print(f"Product Price: {p_price}")  # Log giá sản phẩm

    # Log hình ảnh sản phẩm
    try:
        elm_pic = driver.find_elements(By.CLASS_NAME,"slick-list")[0].find_elements(By.TAG_NAME,"img")
        pic_url = ""
        for m in range(len(elm_pic)):
            pic_url = pic_url + elm_pic[m].get_attribute("src") + "|"
    except Exception as e:
        print(f"Error while fetching product images: {e}")
    print(f"Product Images: {pic_url}")  # Log đường dẫn hình ảnh sản phẩm

    if p_name == "":
        sel_not = "out of stock"

    if p_name != "":
        tr_sh.Cells(r_nm, 4).Value = p_name
    tr_sh.Cells(r_nm, 5).Value = sel_not
    if p_det != "":
        tr_sh.Cells(r_nm, 6).Value = p_det
    if p_price != "":
        tr_sh.Cells(r_nm, 7).Value = p_price
    if pic_url != "":
        tr_sh.Cells(r_nm, 8).Value = pic_url

    if tr_sh.Cells(r_nm, 3).Value is not None:
        if sel_not == "in stock":
            # if abs(float(tr_sh.Cells(r_nm, 9).Value) - float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm, 3).Value), int(tr_sh.Cells(r_nm, 16).Value))  # 価格変更
            #         tr_sh.Cells(r_nm, 9).Value = NgrYD  # 価格変更時為替レート記入
            #     except Exception as e:
            #         print(f"Error while changing price: {e}")

            if tr_sh.Cells(r_nm, 11).Value is not None:  # 販売停止中（timestumpが存在）
                try:
                    ChngQuantity(tr_sh.Cells(r_nm, 3).Value, 1)
                    tr_sh.Cells(r_nm, 11).Value = ""
                except Exception as e:
                    print(f"Error while changing quantity: {e}")

        elif sel_not == "out of stock":
            z_range = 0
            try:
                z_range = time.time() - tr_sh.Cells(r_nm, 11).Value
            except Exception as e:
                print(f"Error while calculating time range: {e}")

            if tr_sh.Cells(r_nm, 11).Value is not None and zkt < z_range:  # 販売停止中で（timestumpが存在して）タイマー経過していれば
                try:
                    DlePrdct(tr_sh.Cells(r_nm, 3).Value)
                    tr_sh.Cells(r_nm, 3).Value = ""
                    tr_sh.Cells(r_nm, 11).Value = ""
                except Exception as e:
                    print(f"Error while deleting product: {e}")

            elif tr_sh.Cells(r_nm, 11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm, 3).Value, 0)
                    tr_sh.Cells(r_nm, 11).Value = time.time()
                except Exception as e:
                    print(f"Error while changing quantity to zero: {e}")

#https://www.rodeodrive.co.jp/
def get_det_26 (prd_url,r_nm):

    print(f"Accessing URL: {prd_url}")  # In ra URL

    global driver
    global tr_sh
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)   

    time.sleep(3)
    p_id = ""
    p_name = ""
    sel_not = ""
    p_det = ""
    p_price = ""
    pic_url = ""
    tex_yn = ""  # Gán giá trị mặc định cho tex_yn

    # Log tên sản phẩm
    try:
        p_name = driver.find_elements(By.CLASS_NAME,"goodsdetail_goods_name")[0].text
    except Exception as e:
        print(f"Error while fetching product name: {e}")
    print(f"Product Name: {p_name}")  # Log tên sản phẩm

    # Log trạng thái sản phẩm
    try:
        tex_yn = driver.find_elements(By.CLASS_NAME,"p-link--button__txt")[0].text
    except Exception as e:
        print(f"Error while fetching stock status text: {e}")
    print(f"Stock Status Text: {tex_yn}")  # Log trạng thái sản phẩm

    if "ショッピングカートに入れる" in tex_yn:
        sel_not = "in stock"
        print("Status: in stock")  # Log trạng thái "in stock"
    else:
        sel_not = "out of stock"
        print("Status: out of stock")  # Log trạng thái "out of stock"
        
    # Log chi tiết sản phẩm
    try:
        elm_det = driver.find_elements(By.CLASS_NAME,"p-table__content")[0].find_elements(By.TAG_NAME,"tr")
        for n in range(len(elm_det)):
            try:
                p_det = p_det + elm_det[n].find_elements(By.TAG_NAME,"th")[0].text + " : " + elm_det[n].find_elements(By.TAG_NAME,"td")[0].text + "<br>"
            except:
                pass
    except Exception as e:
        print(f"Error while fetching product details: {e}")
    print(f"Product Details: {p_det}")  # Log chi tiết sản phẩm
        
    # Log giá sản phẩm
    try:
        p_price = driver.find_elements(By.CLASS_NAME,"js-enhanced-ecommerce-goods-price")[0].text
        p_price = p_price.replace(',', '')
        p_price = p_price.replace('￥', '')
        p_price = p_price.replace('(税込)', '')
        p_price = p_price.replace('円', '')
        p_price = int(p_price)
    except Exception as e:
        print(f"Error while fetching product price: {e}")
    print(f"Product Price: {p_price}")  # Log giá sản phẩm

    # Log hình ảnh sản phẩm
    try:
        elm_pic = driver.find_elements(By.CLASS_NAME,"splide__track")[0].find_elements(By.TAG_NAME,"img")
        pic_url = ""
        for m in range(len(elm_pic)):
            pic_url = pic_url + elm_pic[m].get_attribute("src") + "|"
    except Exception as e:
        print(f"Error while fetching product images: {e}")
    print(f"Product Images: {pic_url}")  # Log đường dẫn hình ảnh sản phẩm

    if p_name == "":
        sel_not = "out of stock"

    if p_name != "":
        tr_sh.Cells(r_nm, 4).Value = p_name
    tr_sh.Cells(r_nm, 5).Value = sel_not
    if p_det != "":
        tr_sh.Cells(r_nm, 6).Value = p_det
    if p_price != "":
        tr_sh.Cells(r_nm, 7).Value = p_price
    if pic_url != "":
        tr_sh.Cells(r_nm, 8).Value = pic_url

    if tr_sh.Cells(r_nm, 3).Value is not None:
        if sel_not == "in stock":
            # if abs(float(tr_sh.Cells(r_nm, 9).Value) - float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm, 3).Value), int(tr_sh.Cells(r_nm, 16).Value))  # 価格変更
            #         tr_sh.Cells(r_nm, 9).Value = NgrYD  # 価格変更時為替レート記入
            #     except Exception as e:
            #         print(f"Error while changing price: {e}")

            if tr_sh.Cells(r_nm, 11).Value is not None:  # 販売停止中（timestumpが存在）
                try:
                    ChngQuantity(tr_sh.Cells(r_nm, 3).Value, 1)
                    tr_sh.Cells(r_nm, 11).Value = ""
                except Exception as e:
                    print(f"Error while changing quantity: {e}")

        elif sel_not == "out of stock":
            z_range = 0
            try:
                z_range = time.time() - tr_sh.Cells(r_nm, 11).Value
            except Exception as e:
                print(f"Error while calculating time range: {e}")

            if tr_sh.Cells(r_nm, 11).Value is not None and zkt < z_range:  # 販売停止中で（timestumpが存在して）タイマー経過していれば
                try:
                    DlePrdct(tr_sh.Cells(r_nm, 3).Value)
                    tr_sh.Cells(r_nm, 3).Value = ""
                    tr_sh.Cells(r_nm, 11).Value = ""
                except Exception as e:
                    print(f"Error while deleting product: {e}")

            elif tr_sh.Cells(r_nm, 11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm, 3).Value, 0)
                    tr_sh.Cells(r_nm, 11).Value = time.time()
                except Exception as e:
                    print(f"Error while changing quantity to zero: {e}")

#https://www.treasure-f.com/
def get_det_27 (prd_url,r_nm):

    print(f"Accessing URL: {prd_url}")  # In ra URL

    global driver
    global tr_sh
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)   

    time.sleep(3)
    p_id = ""
    p_name = ""
    sel_not = ""
    p_det = ""
    p_price = ""
    pic_url = ""
    tex_yn = ""  # Gán giá trị mặc định cho tex_yn

    # Log tên sản phẩm
    try:
        p_name = driver.find_elements(By.CLASS_NAME,"title")[0].text
    except Exception as e:
        print(f"Error while fetching product name: {e}")
    print(f"Product Name: {p_name}")  # Log tên sản phẩm

    # Log trạng thái sản phẩm
    try:
        tex_yn = driver.find_elements(By.CLASS_NAME,"p-link--button__txt")[0].text
    except Exception as e:
        print(f"Error while fetching stock status text: {e}")
    print(f"Stock Status Text: {tex_yn}")  # Log trạng thái sản phẩm

    if "ショッピングカートに入れる" in tex_yn:
        sel_not = "in stock"
        print("Status: in stock")  # Log trạng thái "in stock"
    else:
        sel_not = "out of stock"
        print("Status: out of stock")  # Log trạng thái "out of stock"
        
    # Log chi tiết sản phẩm
    try:
        elm_det = driver.find_elements(By.CLASS_NAME,"p-table__content")[0].find_elements(By.TAG_NAME,"tr")
        for n in range(len(elm_det)):
            try:
                p_det = p_det + elm_det[n].find_elements(By.TAG_NAME,"th")[0].text + " : " + elm_det[n].find_elements(By.TAG_NAME,"td")[0].text + "<br>"
            except:
                pass
    except Exception as e:
        print(f"Error while fetching product details: {e}")
    print(f"Product Details: {p_det}")  # Log chi tiết sản phẩm
        
    # Log giá sản phẩm
    try:
        p_price = driver.find_elements(By.CLASS_NAME,"js-enhanced-ecommerce-goods-price")[0].text
        p_price = p_price.replace(',', '')
        p_price = p_price.replace('￥', '')
        p_price = p_price.replace('(税込)', '')
        p_price = p_price.replace('円', '')
        p_price = int(p_price)
    except Exception as e:
        print(f"Error while fetching product price: {e}")
    print(f"Product Price: {p_price}")  # Log giá sản phẩm

    # Log hình ảnh sản phẩm
    try:
        elm_pic = driver.find_elements(By.CLASS_NAME,"ql-editor")[0].find_elements(By.TAG_NAME,"img")
        pic_url = ""
        for m in range(len(elm_pic)):
            pic_url = pic_url + elm_pic[m].get_attribute("src") + "|"
    except Exception as e:
        print(f"Error while fetching product images: {e}")
    print(f"Product Images: {pic_url}")  # Log đường dẫn hình ảnh sản phẩm

    if p_name == "":
        sel_not = "out of stock"

    if p_name != "":
        tr_sh.Cells(r_nm, 4).Value = p_name
    tr_sh.Cells(r_nm, 5).Value = sel_not
    if p_det != "":
        tr_sh.Cells(r_nm, 6).Value = p_det
    if p_price != "":
        tr_sh.Cells(r_nm, 7).Value = p_price
    if pic_url != "":
        tr_sh.Cells(r_nm, 8).Value = pic_url

    if tr_sh.Cells(r_nm, 3).Value is not None:
        if sel_not == "in stock":
            # if abs(float(tr_sh.Cells(r_nm, 9).Value) - float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm, 3).Value), int(tr_sh.Cells(r_nm, 16).Value))  # 価格変更
            #         tr_sh.Cells(r_nm, 9).Value = NgrYD  # 価格変更時為替レート記入
            #     except Exception as e:
            #         print(f"Error while changing price: {e}")

            if tr_sh.Cells(r_nm, 11).Value is not None:  # 販売停止中（timestumpが存在）
                try:
                    ChngQuantity(tr_sh.Cells(r_nm, 3).Value, 1)
                    tr_sh.Cells(r_nm, 11).Value = ""
                except Exception as e:
                    print(f"Error while changing quantity: {e}")

        elif sel_not == "out of stock":
            z_range = 0
            try:
                z_range = time.time() - tr_sh.Cells(r_nm, 11).Value
            except Exception as e:
                print(f"Error while calculating time range: {e}")

            if tr_sh.Cells(r_nm, 11).Value is not None and zkt < z_range:  # 販売停止中で（timestumpが存在して）タイマー経過していれば
                try:
                    DlePrdct(tr_sh.Cells(r_nm, 3).Value)
                    tr_sh.Cells(r_nm, 3).Value = ""
                    tr_sh.Cells(r_nm, 11).Value = ""
                except Exception as e:
                    print(f"Error while deleting product: {e}")

            elif tr_sh.Cells(r_nm, 11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm, 3).Value, 0)
                    tr_sh.Cells(r_nm, 11).Value = time.time()
                except Exception as e:
                    print(f"Error while changing quantity to zero: {e}")

#https://shop.the-ticken.com/
def get_det_28 (prd_url,r_nm):

    print(f"Accessing URL: {prd_url}")  # In ra URL

    global driver
    global tr_sh
    global tim_sh
    global RateYD
    global rangYD
    global NgrYD
    driver.get(prd_url)   

    time.sleep(3)
    p_id = ""
    p_name = ""
    sel_not = ""
    p_det = ""
    p_price = ""
    pic_url = ""
    tex_yn = ""  # Gán giá trị mặc định cho tex_yn

    # Log tên sản phẩm
    try:
        p_name = driver.find_elements(By.CLASS_NAME,"title")[0].text
    except Exception as e:
        print(f"Error while fetching product name: {e}")
    print(f"Product Name: {p_name}")  # Log tên sản phẩm

    # Log trạng thái sản phẩm
    try:
        tex_yn = driver.find_elements(By.CLASS_NAME,"p-link--button__txt")[0].text
    except Exception as e:
        print(f"Error while fetching stock status text: {e}")
    print(f"Stock Status Text: {tex_yn}")  # Log trạng thái sản phẩm

    if "ショッピングカートに入れる" in tex_yn:
        sel_not = "in stock"
        print("Status: in stock")  # Log trạng thái "in stock"
    else:
        sel_not = "out of stock"
        print("Status: out of stock")  # Log trạng thái "out of stock"
        
    # Log chi tiết sản phẩm
    try:
        elm_det = driver.find_elements(By.CLASS_NAME,"p-table__content")[0].find_elements(By.TAG_NAME,"tr")
        for n in range(len(elm_det)):
            try:
                p_det = p_det + elm_det[n].find_elements(By.TAG_NAME,"th")[0].text + " : " + elm_det[n].find_elements(By.TAG_NAME,"td")[0].text + "<br>"
            except:
                pass
    except Exception as e:
        print(f"Error while fetching product details: {e}")
    print(f"Product Details: {p_det}")  # Log chi tiết sản phẩm
        
    # Log giá sản phẩm
    try:
        p_price = driver.find_elements(By.CLASS_NAME,"js-enhanced-ecommerce-goods-price")[0].text
        p_price = p_price.replace(',', '')
        p_price = p_price.replace('￥', '')
        p_price = p_price.replace('(税込)', '')
        p_price = p_price.replace('円', '')
        p_price = int(p_price)
    except Exception as e:
        print(f"Error while fetching product price: {e}")
    print(f"Product Price: {p_price}")  # Log giá sản phẩm

    # Log hình ảnh sản phẩm
    try:
        elm_pic = driver.find_elements(By.CLASS_NAME,"ql-editor")[0].find_elements(By.TAG_NAME,"img")
        pic_url = ""
        for m in range(len(elm_pic)):
            pic_url = pic_url + elm_pic[m].get_attribute("src") + "|"
    except Exception as e:
        print(f"Error while fetching product images: {e}")
    print(f"Product Images: {pic_url}")  # Log đường dẫn hình ảnh sản phẩm

    if p_name == "":
        sel_not = "out of stock"

    if p_name != "":
        tr_sh.Cells(r_nm, 4).Value = p_name
    tr_sh.Cells(r_nm, 5).Value = sel_not
    if p_det != "":
        tr_sh.Cells(r_nm, 6).Value = p_det
    if p_price != "":
        tr_sh.Cells(r_nm, 7).Value = p_price
    if pic_url != "":
        tr_sh.Cells(r_nm, 8).Value = pic_url

    if tr_sh.Cells(r_nm, 3).Value is not None:
        if sel_not == "in stock":
            # if abs(float(tr_sh.Cells(r_nm, 9).Value) - float(NgrYD)) > rangYD:
            #     try:
            #         ChngPrice(int(tr_sh.Cells(r_nm, 3).Value), int(tr_sh.Cells(r_nm, 16).Value))  # 価格変更
            #         tr_sh.Cells(r_nm, 9).Value = NgrYD  # 価格変更時為替レート記入
            #     except Exception as e:
            #         print(f"Error while changing price: {e}")

            if tr_sh.Cells(r_nm, 11).Value is not None:  # 販売停止中（timestumpが存在）
                try:
                    ChngQuantity(tr_sh.Cells(r_nm, 3).Value, 1)
                    tr_sh.Cells(r_nm, 11).Value = ""
                except Exception as e:
                    print(f"Error while changing quantity: {e}")

        elif sel_not == "out of stock":
            z_range = 0
            try:
                z_range = time.time() - tr_sh.Cells(r_nm, 11).Value
            except Exception as e:
                print(f"Error while calculating time range: {e}")

            if tr_sh.Cells(r_nm, 11).Value is not None and zkt < z_range:  # 販売停止中で（timestumpが存在して）タイマー経過していれば
                try:
                    DlePrdct(tr_sh.Cells(r_nm, 3).Value)
                    tr_sh.Cells(r_nm, 3).Value = ""
                    tr_sh.Cells(r_nm, 11).Value = ""
                except Exception as e:
                    print(f"Error while deleting product: {e}")

            elif tr_sh.Cells(r_nm, 11).Value is None:
                try:
                    ChngQuantity(tr_sh.Cells(r_nm, 3).Value, 0)
                    tr_sh.Cells(r_nm, 11).Value = time.time()
                except Exception as e:
                    print(f"Error while changing quantity to zero: {e}")


def get_table_and_go():
    global tim_sh
    global RateYD
    global NgrYD
    global rangYD
    global JpnTaxRate
    global eBayChrgRate
    global zkt
    xl = win32com.client.GetObject(Class="Excel.Application")
    tim_sh= xl.Worksheets("time_table")
    tim_sh.Cells(6,3).Value='=WEBSERVICE("https://api.excelapi.org/currency/rate?pair=usd-jpy")'
    time.sleep(3)
    RateYD=tim_sh.Cells(8,3).Value
    NgrYD=float(tim_sh.Cells(7,3).Value)+float(tim_sh.Cells(8,3).Value)
    rangYD=tim_sh.Cells(10,3).Value
    JpnTaxRate=tim_sh.Cells(9,3).Value
    eBayChrgRate=tim_sh.Cells(8,6).Value
    zkt=tim_sh.Cells(11,3).Value
    zkt= zkt*86400
    global tr_sh


    # tr_sh= xl.Worksheets("the-ticken")###
    # for i in range(10000):
    #     #print(i)
    #     if tr_sh.Cells(i+3,2).Value!=None:
            
    #         get_det_28(tr_sh.Cells(i+3,2).Value,i+3)
    #     else:
    #         break 

    tr_sh= xl.Worksheets("treasure-f")###
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_27(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break 

    tr_sh= xl.Worksheets("rodeodrive")###
    logging.info("vao check rodeodrive")
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_26(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break 

    tr_sh= xl.Worksheets("bigmoon-kyoto")###
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_25(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break 
    
    
    tr_sh= xl.Worksheets("the gold")###
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            try:
                get_det_1(tr_sh.Cells(i+3,2).Value,i+3)
            except:
                pass
        else:
            break                                
    
    
    tr_sh= xl.Worksheets("sweet road")###
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            try:
                get_det_3(tr_sh.Cells(i+3,2).Value,i+3)
            except:
                pass
        else:
            break
    
    tr_sh= xl.Worksheets("takayamasititen")### 
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            try:
                get_det_5(tr_sh.Cells(i+3,2).Value,i+3)
            except:
                pass
        else:
            break

    tr_sh= xl.Worksheets("kanteikyoku")###                                      
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_7(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break

    tr_sh= xl.Worksheets("elady")                                          
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_9(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break

    tr_sh= xl.Worksheets("daikokuya")###                                        
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_11(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break
    
    tr_sh= xl.Worksheets("yafuoku")###
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_15(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break
    
    tr_sh= xl.Worksheets("firekids")###                                        
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_21(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break
    
    tr_sh= xl.Worksheets("bigmoon")###                                       
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_22(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break
          
    tr_sh= xl.Worksheets("ticken")###
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_23(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break

    tr_sh= xl.Worksheets("komehyo")###
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_24(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break 

    
def clse():
    try:
        driver.quit()
        sys.exit()
    except:
        pass
    base.destroy()
    sys.exit()
    

def get_table_main():
    logging.info("Nhan nut khoi chay")
    tim_sh= xl.Worksheets("time_table")
    for k in range(10000):
        get_table_and_go()
        upl=int(tim_sh.Cells(4,2).Value)
        print(upl)
        time.sleep(upl*60)   


bebe=""
rex=0
total_data=[]
base=tk.Tk()
iconfile = '008.ico'
base.iconbitmap(iconfile)
base.geometry("350x100")
base.configure(bg='chocolate1')
base.title("eBay_11ShopsMainVer3.03")
my_font = font.Font(base,family="Arial",size=15,weight="bold")


c = tk.Canvas( base, width= 400, height=200,bg='chocolate1' )
c.pack()

txt5 = tk.Entry(bg='gray15',insertbackground="snow",fg="snow")
txt5.insert(0, 3)
txt5.place(x=400, y=20,width=40,height=20)
euu=int(txt5.get())

button2=tk.Button(base,text="インターバル運用 START",width=18,height=1,command=get_table_main,bg='light steel blue').place(x=10,y=15)
button1=tk.Button(base,text="CLOSE",width=12,height=1,command=clse,bg='light steel blue').place(x=200,y=15)

base.mainloop()