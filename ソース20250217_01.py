## eBay_15ShopsVer_1.05
from selenium import webdriver
options = webdriver.ChromeOptions()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
import zipfile
import urllib.request                            # ライブラリを取り込む
import os
import time
import datetime
import pandas as pd
import winsound
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

import win32com.client
xl = win32com.client.GetObject(Class="Excel.Application")
#driver.quit()
driver = webdriver.Chrome(executable_path='ccm/chromedriver',options=options)

def dl_pro(pr_id):
    try:
        api = Trading(appid="", devid="", certid="", token="",config_file=None)
        api.execute('EndFixedPriceItem', { "EndingReason":"LostOrBroken","ItemID":pr_id})
    except ConnectionError as e:
        print(e)

def fin():
    sys.exit()

def get_det_1(prd_url,r_nm):
    global driver
    global tr_sｈ
    driver.get(prd_url)#①
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    try:
        p_id=driver.find_elements_by_class_name("fs-c-productNumber__number")[0].text
    except:
        pass
    try:  
        p_name=driver.find_elements_by_class_name("fs-c-productNameHeading__name")[0].text
    except:
        pass
   
    try:
        tex_yn=""
        tex_yn=driver.find_elements_by_class_name("fs-c-productNotice--outOfStock")[0].text
        #print(tex_yn)
    except:
        pass    
    if "在庫がございません" in tex_yn:
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass
    else:
        sel_not="in stock"

    try:  
        elm_det=driver.find_elements_by_tag_name("table")[0].find_elements_by_tag_name("tr")
    except:
        pass
    try:
        for n in range(len(elm_det)-1):
            #print("TD")
            #print(elm_det[n+1].find_elements_by_tag_name("td")[0].text)
            p_det= p_det+elm_det[n+1].find_elements_by_tag_name("td")[0].text+" : "+elm_det[n+1].find_elements_by_tag_name("td")[1].text+"<br>"
    except:
        pass
    try:
        p_price=driver.find_elements_by_class_name("fs-c-price__value")[0].text
        p_price=p_price.replace(',', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements_by_class_name("fs-c-productThumbnail")[0].find_elements_by_tag_name("img")
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
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
    
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url


def get_det_2(prd_url,r_nm):
    global driver
    global tr_sｈ
    driver.get(prd_url)
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    
    try:
        p_name=driver.find_element_by_id("content_tit").text
    except:
        pass
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements_by_class_name("buy_btn")[0].find_elements_by_tag_name("img")[0].get_attribute("title")
    except:
        pass
    
    if tex_yn=="SOLDOUT":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass
    else:
        sel_not="in stock"
        
    try:
        elm_det=driver.find_element_by_id("item").find_elements_by_tag_name("tr")
        for n in range(len(elm_det)):
            p_det= p_det+elm_det[n].find_elements_by_tag_name("td")[0].text+" : "+elm_det[n].find_elements_by_tag_name("td")[1].text+"<br>"

        p_det= p_det+"商品説明:<br>"+driver.find_elements_by_class_name("desc_box")[1].text+"<br><br>"

        p_det= p_det+"商品状態:<br>"+driver.find_elements_by_class_name("desc_box")[3].text

    except:
        pass
    try:    
        p_price=driver.find_element_by_id("price").find_elements_by_class_name("right")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('円', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements_by_class_name("detail_img")[0].find_elements_by_tag_name("img")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("src")+"|"
        pic_url=pic_url.replace('width=88&height=66', 'width=800&height=600')
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass
    
    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
    
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri

    
    
    
    
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url
        
    
def get_det_3(prd_url,r_nm):
    global driver
    global tr_sｈ
    driver.get(prd_url)
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    try:
        p_id=driver.find_elements_by_class_name("")[0].text
    except:
        pass
    try:
        p_name=driver.find_elements_by_class_name("product-name")[0].text
    except:
        pass
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements_by_class_name("detail-sold-out")[0].text
    except:
        pass
    
    if tex_yn=="SOLD OUT":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass
    else:
        sel_not="in stock"
        
    try:
        p_det=driver.find_elements_by_class_name("description-text")[0].text
    except:
        pass
    try:    
        p_price=driver.find_elements_by_class_name("product-price-block")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('￥', '')
        p_price=p_price.replace('（税込）', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements_by_class_name("grid-multi-image")[0].find_elements_by_tag_name("img")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("src")+"|"
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url
    
        
def get_det_4(prd_url,r_nm):
    global driver
    global tr_sｈ
    driver.get(prd_url)
    #driver.get("https://housekihiroba.jp/shop/g/g608695001/")#④
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    
    try:
        p_id=driver.find_elements_by_class_name("goodsspec_")[0].find_elements_by_class_name("redline02")[0].text
    except:
        pass
    
    try:
         p_name=p_name+driver.find_elements_by_class_name("brand_name_")[0].text+" "
    except:
        pass
    try:
        p_name=p_name+driver.find_elements_by_class_name("goodsspec_")[0].find_elements_by_class_name("goods_name_")[0].text
    except:
        pass
    
    try:
        tex_yn=""
        tex_yn=driver.find_element_by_id("goods_stock").find_elements_by_tag_name("td")[0].text
    except:
        pass    
    if "在庫有り" in tex_yn:
        sel_not="in stock"
    else:
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    try:
        elm_det=driver.find_elements_by_class_name("formdetail_")[0].find_elements_by_tag_name("tr")
        for n in range(len(elm_det)):
            p_det= p_det+elm_det[n].find_elements_by_tag_name("th")[0].text+" : "+elm_det[n].find_elements_by_tag_name("td")[0].text+"<br>"
    except:
        pass
    try:
        p_price=driver.find_element_by_id("goods_price").find_elements_by_class_name("price_")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('円', '')
        p_price=int(p_price)
        
    except:
        pass
    try:
        elm_pic=driver.find_elements_by_class_name("etc_goodsimg_")[0].find_elements_by_tag_name("a")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("href")+"|"    
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
        
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url
    
def get_det_5(prd_url,r_nm):
    global driver
    global tr_sｈ
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
        p_id=driver.find_element_by_id("spec_code_").find_elements_by_tag_name("dd")[0].text
    except:
        pass
    try:
        p_name=driver.find_elements_by_class_name("goods_name_")[0].text
    except:
        pass
        #document.getElementById("spec_stock_")
    
    zaik_an="aa"
    try:
        zaik_an=driver.find_elements_by_class_name("mainspec_")[0].find_elements_by_class_name("stock_")[0].text
    except:
        pass
    
    if "在庫あり" in zaik_an:
        sel_not="in stock"
    else:
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    try:
        elm_det=driver.find_elements_by_tag_name("table")[1].find_elements_by_tag_name("tr")
        for n in range(len(elm_det)):
            p_det= p_det+elm_det[n].find_elements_by_tag_name("td")[0].text+" : "+elm_det[n].find_elements_by_tag_name("td")[1].text+"<br>"
    except:
        pass
    try:
        p_price=driver.find_elements_by_class_name("price_")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('円', '')
        p_price=int(p_price)

    except:
        pass
    try:
        elm_pic=driver.find_elements_by_class_name("goodsimg_")[0].find_elements_by_tag_name("a")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("href")+"|"           
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
        
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url
    
    
def get_det_6(prd_url,r_nm):
    global driver
    global tr_sｈ
    driver.get(prd_url)
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    
    try:
        p_id=driver.find_elements_by_class_name("sku__value")[0].text
    except:
        pass
    try:
        p_name=driver.find_elements_by_class_name("title")[0].text
    except:
        pass
    
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements_by_class_name("product-unavailable")[0].text
    except:
        pass    
    if "完売" in tex_yn:
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass
    else:
        sel_not="in stock"
    
    
    
    try:
        elm_det=driver.find_elements_by_class_name("itemdetail")[0].find_elements_by_tag_name("tr")
        for n in range(len(elm_det)):
            p_det= p_det+elm_det[n].find_elements_by_tag_name("th")[0].text+" : "+elm_det[n].find_elements_by_tag_name("td")[0].text+"<br>"
    except:
        pass
    try:
        p_price=driver.find_elements_by_class_name("current-price")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('¥', '')
        p_price=int(p_price)
    except:
        pass   
    try:
        elm_pic=driver.find_elements_by_class_name("thumbnails")[0].find_elements_by_tag_name("a")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("href")+"|"     
    except:
        pass
    
    if p_name=="":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    tr_sh.Cells(r_nm,9).Value=""
    
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
        
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url
    
def get_det_7(prd_url,r_nm):
    global driver
    global tr_sｈ
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
        p_id=driver.find_elements_by_class_name("")[0].text
    except:
        pass
    try:
        p_name=driver.find_elements_by_class_name("ec-headingTitle")[0].text
    except:
        pass   
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements_by_class_name("add-cart")[0].text
    except:
        pass    
    if "カートに追加" in tex_yn:
        sel_not="in stock"
    else:
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass
        
    try:
        elm_det=driver.find_elements_by_class_name("ec-productRole__description")[0].find_elements_by_tag_name("tr")
        for n in range(len(elm_det)-1):
            if elm_det[n+1].find_elements_by_tag_name("th")[0].text=="管理番号":
                p_id=elm_det[n+1].find_elements_by_tag_name("td")[0].text
            p_det= p_det+elm_det[n+1].find_elements_by_tag_name("th")[0].text+" : "+elm_det[n+1].find_elements_by_tag_name("td")[0].text+"<br>"
    except:
        pass
    try:
        p_price=driver.find_elements_by_class_name("ec-price__price")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('￥', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements_by_class_name("item_nav")[0].find_elements_by_tag_name("img")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("src")+"|"
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
    
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url
    
    
def get_det_8(prd_url,r_nm):
    global driver
    global tr_sｈ
    driver.get(prd_url)   
    #⑧パーパス
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    
    try:
        p_name=driver.find_elements_by_class_name("item-detail-title")[0].text
    except:
        pass   
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements_by_class_name("cart-btn")[0].text
    except:
        pass    
    if "カートへ入れる" in tex_yn:
        sel_not="in stock"
    else:
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass
        
    try:
        elm_det=driver.find_element_by_id("accordion").find_elements_by_tag_name("h4")
        for n in range(len(elm_det)-1):
            p_det= p_det+elm_det[n].text+" :<br> "
            
            if elm_det[n].text=="Rank":
                p_det= p_det+driver.find_element_by_id("accordion").find_elements_by_class_name("r_p")[0].get_attribute('innerHTML')+"<br><br>"
            else:
                p_det= p_det+driver.find_element_by_id("accordion").find_elements_by_class_name("accordion-content1")[n].get_attribute('innerHTML')+"<br><br>"
    except:
        pass
    try:
        p_price=driver.find_elements_by_class_name("item-price-text")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('￥', '')
        p_price=p_price.replace('(税込)', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements_by_class_name("slides")[1].find_elements_by_tag_name("img")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("src")+"|"
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
    
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url


def get_det_9(prd_url,r_nm):
    global driver
    global tr_sｈ
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
        p_name=driver.find_elements_by_class_name("detail__name")[0].text
    except:
        pass   
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements_by_class_name("detail__btn_list--cart")[0].text
    except:
        pass    
    if "カートに入れる" in tex_yn:
        sel_not="in stock"
    else:
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass
        
    try:
        elm_det=driver.find_element_by_id("info1_content").find_elements_by_class_name("y")
        for n in range(len(elm_det)):
            p_det=p_det+elm_det[n].text+": "
            p_det= p_det+driver.find_element_by_id("info1_content").find_elements_by_class_name("x")[n].text+"<br>"
    except:
        pass
    
    try:
        driver.find_elements_by_class_name("pc_view_label")[1].click()
        time.sleep(1)
        elm_det=driver.find_element_by_id("info2_content").find_elements_by_class_name("y")
        for n in range(len(elm_det)):
            p_det=p_det+elm_det[n].text+": "
            p_det= p_det+driver.find_element_by_id("info2_content").find_elements_by_class_name("x")[n].text+"<br>"
    except:
        pass

    try:
        p_price=driver.find_elements_by_class_name("price")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('¥', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements_by_class_name("product-image-thumbs")[0].find_elements_by_tag_name("img")
        for m in range(len( elm_pic)):
            pic_url=pic_url+(elm_pic[m].get_attribute("src")).replace('/250', '/1000')+"|"
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
    
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url
    
    
def get_det_10(prd_url,r_nm):
    global driver
    global tr_sｈ
    driver.get(prd_url)   
    #⑩質ウエダ
    time.sleep(3)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    
    try:
        p_name=driver.find_elements_by_tag_name("h1")[0].text
    except:
        pass   
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements_by_class_name("cartBtn")[0].text
    except:
        pass    
    if "かごに追加" in tex_yn:
        sel_not="in stock"
    else:
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass
        
    try:
        elm_det=driver.find_elements_by_class_name("spec")[1].find_elements_by_tag_name("tr")
        for n in range(len(elm_det)):
            p_det= p_det+elm_det[n].find_elements_by_tag_name("th")[0].text+" : "+elm_det[n].find_elements_by_tag_name("td")[0].text+"<br>"
    except:
        pass
    try:
        elm_det=driver.find_elements_by_class_name("spec")[0].find_elements_by_tag_name("tr")
        for n in range(len(elm_det)):
            p_det= p_det+elm_det[n].find_elements_by_tag_name("th")[0].text+" : "+elm_det[n].find_elements_by_tag_name("td")[0].text+"<br>"
    except:
        pass
    
    
    try:
        p_price=driver.find_elements_by_class_name("price")[1].find_elements_by_tag_name("small")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('￥', '')
        p_price=p_price.replace('(税込)', '')
        p_price=p_price.replace('円', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_element_by_id("itemslider-pager").find_elements_by_tag_name("img")
        for m in range(len( elm_pic)):
            
            pic_url=pic_url+elm_pic[m].get_attribute("src")+"|"
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
    
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url
    

def get_det_11(prd_url,r_nm):
    global driver
    global tr_sｈ
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
        p_name=driver.find_elements_by_tag_name("h1")[0].text
    except:
        pass   
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements_by_class_name("btn-primary")[1].text
    except:
        pass    
    if "カートに入れる" in tex_yn:
        sel_not="in stock"
    else:
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    try:
        driver.find_elements_by_class_name("mat-expansion-panel-header-title")[0].click()
        time.sleep(1)
        elm_det=driver.find_elements_by_class_name("pt-4")[1].find_elements_by_tag_name("div")
        for n in range(int(len(elm_det)/3)):
            p_det= p_det+elm_det[3*n+1].text+" : "+elm_det[3*n+2].text+"<br>"
    except:
        pass
    try:
        p_det= p_det+driver.find_elements_by_class_name("pt-2")[1].get_attribute('innerHTML')
    except:
        pass
    
    
    try:
        p_price=driver.find_elements_by_class_name("my-2")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('¥', '')
        p_price=p_price.replace(' ', '')
        p_price=p_price.replace('(税込)', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements_by_class_name("swiper-container")[1].find_elements_by_tag_name("img")
        for m in range(len( elm_pic)):
            pic_url=(pic_url+elm_pic[m].get_attribute("src")+"|").replace('-thumb', '')
            
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
    
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url
    
def get_det_12(prd_url,r_nm):
    global driver
    global tr_sｈ
    driver.get(prd_url)
    #⑫ALLU
    time.sleep(8)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    
    try:
        p_name=driver.find_elements_by_tag_name("h1")[0].text
    except:
        pass   
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements_by_class_name("button--large")[0].text
    except:
        pass    
    if "カートへ入れる" in tex_yn:
        sel_not="in stock"
    else:
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass
        
    try:
        elm_det=driver.find_elements_by_class_name("alc-section-product-description")[0].find_elements_by_class_name("product__description-title")
        for n in range(len(elm_det)):
            p_det= p_det+elm_det[n].text+" : "+driver.find_elements_by_class_name("alc-section-product-description")[0].find_elements_by_class_name("product__description-values")[n].text+"<br>"
    except:
        pass
    try:
        p_price=driver.find_elements_by_class_name("price")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('¥', '')
        p_price=p_price.replace('（税込）', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements_by_class_name("owl-stage")[0].find_elements_by_tag_name("a")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("href")+"|"
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
    
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url
    
def get_det_13(prd_url,r_nm):
    global driver
    global tr_sh
    driver.get(prd_url)
    #⑬HOUBIDOU
    #time.sleep(3)
    time.sleep(9)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    
    try:
        p_name=driver.find_elements_by_tag_name("h1")[0].text
    except:
        pass   
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements_by_class_name("shopify-payment-button__button")[0].text
    except:
        pass    
    if "今すぐ購入する" in tex_yn:
        sel_not="in stock"
    else:
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass
        
    try:
        elm_det=driver.find_elements_by_class_name("table-wrapper")[1].find_elements_by_tag_name("tr")
        p_det= p_det + elm_det[0].text
        for n in range(len(elm_det)):
            p_det= p_det+elm_det[n].find_elements_by_tag_name("td")[0].text+" : "+elm_det[n].find_elements_by_tag_name("td")[1].text+"<br>"
    except:
        pass
    try:
        p_price=driver.find_elements_by_class_name("price--large")[0].text
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('¥', '')
        p_price=p_price.replace('セール価格', '')
        p_price=p_price.replace('(税込)', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements_by_class_name("product__thumbnail-scroll-shadow")[0].find_elements_by_tag_name("img")
        for m in range(len( elm_pic)):
            pic_url=pic_url+(elm_pic[m].get_attribute("src")).replace('_288x', '_600x')+"|"
    except:
        pass
    if p_name=="":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
    
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url
    
    
def get_det_14(prd_url,r_nm):
    global driver
    global tr_sh
    driver.get(prd_url)
    #⑭rehello
    time.sleep(8)
    p_id=""
    p_name=""
    sel_not=""
    p_det=""
    p_price=""
    pic_url=""
    
    try:
        p_name=driver.find_elements_by_class_name("ProductHeading__title")[0].text
    except:
        pass   
    
    try:
        tex_yn=""
        tex_yn=driver.find_elements_by_class_name("ShoppingMenu__soldOut")[0].text
    except:
        pass    
    if "SOLD OUT" in tex_yn:
        sel_not="out of stock"
    else:
        sel_not="in stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass
        
    try:
        p_det= p_det+driver.find_elements_by_tag_name("tbody")[4].text
    except:
        pass
    try:
        elm_det=driver.find_elements_by_tag_name("tbody")[8].find_elements_by_tag_name("tr")

        for n in range(len(elm_det)-1):

            u_te=elm_det[n+1].find_elements_by_tag_name("td")[0].text
            if "採寸イメージ" in u_te:
                break
            
            p_det= p_det+elm_det[n+1].find_elements_by_tag_name("td")[0].text+" : "+elm_det[n+1].find_elements_by_tag_name("td")[1].text+"<br>"


    except:
        pass
    try:
        elm_det=driver.find_elements_by_tag_name("tbody")[12].find_elements_by_tag_name("tr")
        for n in range(len(elm_det)):
            p_det= p_det+elm_det[n].find_elements_by_tag_name("td")[0].text+" : "+elm_det[n].find_elements_by_tag_name("td")[1].text+"<br>"
    except:
        pass
    try:
        p_det= p_det+driver.find_elements_by_tag_name("tbody")[0].find_elements_by_tag_name("b")[2].text+" : "+driver.find_elements_by_tag_name("tbody")[0].find_elements_by_tag_name("b")[3].text+"<br>"
    except:
        pass
    
    try:
        p_det= p_det+(driver.find_elements_by_tag_name("tbody")[4].text).replace('ダメージ詳細', 'ダメージ詳細 : ')+"<br>"
    except:
        pass

 
    try:
        p_price=driver.find_elements_by_class_name("ProductDetail")[0].find_elements_by_class_name("money")[0].text
        p_price=driver.find_elements_by_class_name("ProductDetail")[0].find_elements_by_class_name("money")[1].text
    except:
        pass
    try:
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('¥', '')
        p_price=p_price.replace('セール価格', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements_by_class_name("ProductDetailSlide__thumbnail")[0].find_elements_by_tag_name("img")
        for m in range(len( elm_pic)):
            pic_url=pic_url+elm_pic[m].get_attribute("src")+"|"
    except:
        pass
    
    try:
        pic_url=pic_url+driver.find_elements_by_tag_name("tbody")[8].find_elements_by_tag_name("img").get_attribute("src")+"|"
    except:
        pass
    
    if p_name=="":
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
    
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url


    

def get_det_15(prd_url,r_nm):
    global driver
    global tr_sh
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
        p_name=driver.find_elements_by_class_name("ProductTitle__text")[0].text###OK
    except:
        pass   

    try:
        tex_yn=""
        tex_yn=driver.find_elements_by_class_name("ClosedHeader__tag")[0].text
    except:
        pass    
    if "このオークションは終了" in tex_yn:
        sel_not="out of stock"
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass
    
    else:
        sel_not="in stock"
        
        

    
    try:
        if p_det=="":
            #document.getElementsByClassName("ProductExplanation__commentBody")
            #p_det=driver.find_elements_by_class_name("ProductExplanation__commentBody")[0].get_attribute('outerHTML')#OKKKKK
            p_det=driver.find_elements_by_class_name("ProductExplanation__commentBody")[0].get_attribute('outerText')
            #p_det=driver.find_elements_by_class_name("ProductExplanation__commentBody")[0].find_elements_by_tag_name("div")[0].get_attribute('outerHTML')#A1
            #p_det=driver.find_elements_by_class_name("ProductExplanation__commentBody")[0].find_elements_by_tag_name("div")[0].get_attribute('outerText')
            
            #document.getElementsByClassName("ProductExplanation__commentArea")
            #p_det=driver.find_elements_by_class_name("ProductExplanation__commentArea")[0].get_attribute('innerHTML')
    except:
        pass
    
    try:
        if p_det=="":
            elm_det=driver.find_elements_by_tag_name("tbody")[1].find_elements_by_tag_name("tr")
            for n in range(len(elm_det)):
                try:
                    p_det= p_det+elm_det[n].find_elements_by_tag_name("td")[0].text+" : "+elm_det[n].find_elements_by_tag_name("td")[1].text+"<br>"
                except:
                    pass
    except:
        pass
    
    try:
        if p_det=="":
            p_det=driver.find_elements_by_tag_name("tbody")[1].get_attribute('innerHTML')
    except:
        pass
    
    try:
        p_price=driver.find_elements_by_class_name("Price__value")[0].text
    except:
        pass
    try:
        p_price.find('（')
        p_price= p_price[: p_price.find('（')]
        p_price=p_price.replace(',', '')
        p_price=p_price.replace('¥', '')
        p_price=p_price.replace('円', '')
        p_price=int(p_price)
    except:
        pass
    try:
        elm_pic=driver.find_elements_by_class_name("ProductImage__thumbnail")
        for m in range(len( elm_pic)):
            pind_url=elm_pic[m].find_elements_by_tag_name("img")[0].get_attribute("src")
            pind_url=pind_url[:pind_url.find('.jpg')+4]
            pic_url=pic_url+ pind_url+"|"
    except:
        pass
   
    
    if p_name=="":
        sel_not="out of stock"
    
    if sel_not=="out of stock":    
        try:
            ppid=tr_sh.Cells(r_nm,3).Value
            ppid=int(ppid)
            dl_pro(ppid)
        except:
            pass

    tr_sh.Cells(r_nm,4).Value=p_name
    tr_sh.Cells(r_nm,5).Value=sel_not
    tr_sh.Cells(r_nm,6).Value= p_det
    
    tr_sh.Cells(r_nm,9).Value=""
    ag_pri=""
    try:
        ag_pri=tr_sh.Cells(r_nm,7).Value
        ag_pri=int(ag_pri)
    except:
        pass    
    if ag_pri!=p_price and ag_pri!="" and p_price!="":
        tr_sh.Cells(r_nm,9).Value=ag_pri
    
    tr_sh.Cells(r_nm,7).Value= p_price
    tr_sh.Cells(r_nm,8).Value=pic_url

    
def get_table_and_go():
    #global xl
    global tr_sh
    xl = win32com.client.GetObject(Class="Excel.Application")
    
    tr_sh= xl.Worksheets("the gold")
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            try:
                get_det_1(tr_sh.Cells(i+3,2).Value,i+3)
            except:
                pass
        else:
            break
                                                
  
    tr_sh= xl.Worksheets("lips")        
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            try:
                get_det_2(tr_sh.Cells(i+3,2).Value,i+3)
            except:
                pass
        else:
            break
    
    
    tr_sh= xl.Worksheets("sweet road")        
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            try:
                get_det_3(tr_sh.Cells(i+3,2).Value,i+3)
            except:
                pass
        else:
            break

    tr_sh= xl.Worksheets("housekihiroba") 
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            try:
                get_det_4(tr_sh.Cells(i+3,2).Value,i+3)
            except:
                pass
        else:
            break
   
    tr_sh= xl.Worksheets("takayamasititen")         
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            try:
                get_det_5(tr_sh.Cells(i+3,2).Value,i+3)
            except:
                pass
        else:
            break
      
    tr_sh= xl.Worksheets("baiseru") 
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            try:
                get_det_6(tr_sh.Cells(i+3,2).Value,i+3)
            except:
                pass
        else:
            break
   
    tr_sh= xl.Worksheets("kanteikyoku")                                          
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_7(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break
     
    tr_sh= xl.Worksheets("purpose")                                          
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_8(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break
    
    tr_sh= xl.Worksheets("elady")                                          
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_9(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break


    tr_sh= xl.Worksheets("ueda")                                          
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_10(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break

    tr_sh= xl.Worksheets("daikokuya")                                          
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_11(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break
 
    tr_sh= xl.Worksheets("allu")                                          
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_12(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break
            
          
    tr_sh= xl.Worksheets("houbidou")                                          
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_13(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break

    tr_sh= xl.Worksheets("rehello")                                          
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_14(tr_sh.Cells(i+3,2).Value,i+3)
        else:
            break

    tr_sh= xl.Worksheets("yafuoku")                                          
    for i in range(10000):
        #print(i)
        if tr_sh.Cells(i+3,2).Value!=None:
            
            get_det_15(tr_sh.Cells(i+3,2).Value,i+3)
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
    
            
#t1 = threading.Thread(target=get_table_and_go)

#def get_table_and_go_th():
    #t1.start()
    
    
def get_table_main():
    tim_sh= xl.Worksheets("time_table")
    for k in range(10000):
        get_table_and_go()
        upl=int(tim_sh.Cells(4,2).Value)
        print(upl)
        time.sleep(upl*60)   
    
def get_table_time():
    tim_sh= xl.Worksheets("time_table")
    for k in range(10000):
        uph=str(tim_sh.Cells(4,5).Value)
        print(uph)
        upm=str(tim_sh.Cells(4,6).Value)
        print(upm)
        un=datetime.datetime.now()
        h_st=un.strftime('%H')
        m_st=un.strftime('%M')
        if h_st==uph and m_st==upm:
            get_table_and_go()
            break
        time.sleep(15) 
    

bebe=""
rex=0
total_data=[]
base=tk.Tk()
iconfile = '008.ico'
base.iconbitmap(iconfile)
base.geometry("350x100")
base.configure(bg='chocolate1')
base.title("eBay_15ShopsMainVer 1.05")
my_font = font.Font(base,family="Arial",size=15,weight="bold")


c = tk.Canvas( base, width= 400, height=200,bg='chocolate1' )
c.pack()

button2=tk.Button(base,text="インターバル運用 START",width=18,height=1,command=get_table_main,bg='light steel blue').place(x=10,y=15)
button3=tk.Button(base,text="時限運用 START",width=18,height=1,command=get_table_time,bg='light steel blue').place(x=10,y=55)
button1=tk.Button(base,text="CLOSE",width=12,height=1,command=clse,bg='light steel blue').place(x=200,y=15)

base.mainloop()