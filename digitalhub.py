from email.mime import image
import os
import sys
import json
from time import sleep
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
# from datetime import datetime
import chromedriver_autoinstaller
from models.product import Product
from models.variant import Variant
from models.metafields import Metafields
import glob
import requests
# import pandas as pd

from openpyxl import Workbook
from openpyxl.drawing.image import Image as Imag
from openpyxl.utils import get_column_letter
from PIL import Image
# from natsort 

class Digitalhub_Scraper:
    def __init__(self, DEBUG: bool, result_filename: str) -> None:
        self.DEBUG = DEBUG
        self.result_filename = result_filename
        self.chrome_options = Options()
        self.chrome_options.add_argument('--disable-infobars')
        self.chrome_options.add_argument("--start-maximized")
        self.chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        self.args = ["hide_console", ]
        self.browser = webdriver.Chrome(options=self.chrome_options, service_args=self.args)
        self.data = []
        pass

    def controller(self, brands: list[dict], url: str, username: str, password: str):
        try:
            self.browser.get(url)
            self.wait_until_browsing()

            if self.login(username, password):
                self.browser.get('https://digitalhub.marcolin.com/shop')
                self.wait_until_browsing()

                if self.wait_until_element_found(20, 'xpath', '//div[@id="mCSB_1_container"]'):
                    print('Scraping products for')
                    for brand in brands:

                        brand_urls = self.get_brand_urls(brand)

                        for brand_url in brand_urls:
                            self.open_new_tab(brand_url[0])
                            self.wait_until_browsing()

                            if self.wait_until_element_found(30, 'xpath', '//div[@class="row mt-4 list grid-divider"]/div'):
                                glasses_type = brand_url[1]
                                print(f'Brand: {brand["brand"]} | Type: {glasses_type}')
                                total_products_found = str(self.browser.find_element(By.XPATH, '//div[@class="row mt-4 results"]/div').text).strip().split(' ')[0]
                                print(f'Products found: {str(total_products_found)}')

                                scraped_products = 0
                                while True:
                                    products_data = self.get_all_products_from_page()
                                    scraped_products += len(products_data)

                                    for index, product_data in enumerate(products_data):
                                        number, name, gender, url = '', '', '', ''
                                        number = str(product_data['number']).strip()
                                        name = str(product_data['name']).strip()
                                        gender = str(product_data['gender']).strip()
                                        url = str(product_data['url']).strip()
                                        
                                        self.open_new_tab(url)
                                        self.wait_until_browsing()

                                        if self.wait_until_element_found(25, 'xpath', '//div[@class="col-7 details"]/p/span'):
                                            try:
                                                divs = self.browser.find_element(By.XPATH, '//div[@class="slick-list draggable"]').find_elements(By.XPATH, './/div[@class="slick-track"]/div[@style="width: 128px;"]')
                                                for i in range(0, len(divs)):
                                                    product = Product()
                                                    # product.url = url
                                                    product.brand = brand["brand"]
                                                    product.number = str(number).strip().upper()
                                                    product.name = str(name).strip().upper()
                                                    
                                                    metafields = Metafields()
                                                    metafields.for_who = str(gender).strip().title()
                                                    if metafields.for_who == 'Male': metafields.for_who = 'Men'
                                                    elif metafields.for_who == 'Female': metafields.for_who = 'Women'
                                                    else: metafields.for_who = ''
                                                    product.type = glasses_type

                                                    # print(product.number, metafields.for_who, product.name)
                                                
                                                    try:
                                                        if i == 0: self.move_to_first_variant()

                                                        self.select_variant_image(divs[i])
                                                        self.click_to_make_price_visible()


                                                        try:
                                                            if not product.number and not product.name:
                                                                number_name = str(self.browser.find_element(By.CSS_SELECTOR, 'p.model').text).strip()
                                                                if not product.number:
                                                                    product.number = str(number_name.split(' ')[0]).strip()
                                                                if not product.name and len(number_name) > 1:
                                                                    product.name = str(number_name.split(' ')[-1]).strip()
                                                        except Exception as e:
                                                            if self.DEBUG: print(f'Exception in getting number and name: {str(e)}')
                                                            else: pass

                                                        try:
                                                            for inner_div in self.browser.find_elements(By.XPATH, '//div[@class="details mt-4"]/div[@class="row"]/div'):
                                                                for p in inner_div.find_elements(By.TAG_NAME, 'p'):
                                                                    value = str(p.text).strip()
                                                                    if str('Material').strip().lower() in str(value).lower(): 
                                                                        metafields.frame_material = str(p.find_element(By.TAG_NAME, 'span').text).strip()
                                                                    elif str('Shape').strip().lower() in str(value).lower(): 
                                                                        metafields.frame_shape = str(p.find_element(By.TAG_NAME, 'span').text).strip()
                                                                    elif not metafields.for_who and str('Gender').strip().lower() in str(value).lower(): 
                                                                        metafields.for_who = str(p.find_element(By.TAG_NAME, 'span').text).strip()
                                                        except Exception as e:
                                                            if self.DEBUG: print(f'Exception in getting frame material and shape: {str(e)}')
                                                            else: pass

                                                        try:
                                                            for p in self.browser.find_elements(By.XPATH, '//div[@class="col-7 details"]/p'):
                                                                value = str(p.text).strip()
                                                                if str('Color').strip().lower() in str(value).lower():
                                                                    new_value = str(p.find_element(By.TAG_NAME, 'span').text).strip()
                                                                    # print(i, new_value)
                                                                    product.frame_code = str(new_value).split('-')[0].strip().upper()
                                                                    product.frame_color = str(new_value).split('-')[-1].strip().split(' / ')[0].strip().upper()
                                                                    product.lens_color = str(new_value).split('-')[-1].strip().split(' / ')[-1].strip().upper()
                                                                elif str('Lens type').strip().lower() in str(value).lower():
                                                                    metafields.lens_technology = str(p.find_element(By.TAG_NAME, 'span').text).strip()
                                                        except Exception as e:
                                                            if self.DEBUG: print(f'Exception in getting color code: {str(e)}')
                                                            else: pass

                                                        try:
                                                            sizes, availabilities, prices, gtin = [], [], [], []
                                                            size_titles, wholesale_prices, listing_prices,  availabilities, gtin, sizes = self.get_size_price_status()
                                                            
                                                            if len(size_titles) != len(availabilities):
                                                                print(url, sizes, availabilities)
                                                            else:
                                                                for x in range(0, len(size_titles)):
                                                                    variant = Variant()
                                                                    variant.position = (x + 1)
                                                                    variant.title = size_titles[x]
                                                                    variant.sku = f'{product.number} {product.frame_code} {variant.title}'
                                                                    if availabilities[x] == 'Active': variant.inventory_quantity = 1
                                                                    else: variant.inventory_quantity = 0
                                                                    variant.found_status = 1
                                                                    variant.wholesale_price = wholesale_prices[x]
                                                                    variant.listing_price = listing_prices[x]
                                                                    if len(gtin) > x:
                                                                        variant.barcode_or_gtin = str(gtin[x]).strip()
                                                                    else: variant.barcode_or_gtin = ''
                                                                    variant.size = sizes[x]
                                                                    product.variants = variant
                                                        except Exception as e:
                                                            if self.DEBUG: 
                                                                print(f'Exception in getting sizes, price and availabilities: {e}')
                                                                print(size_titles, prices, availabilities, gtin, sizes)
                                                            else: pass

                                                        try:
                                                            for variant in product.variants:
                                                                if str(variant.size).strip():
                                                                    metafields.product_size += f'{str(variant.size).strip()}, '

                                                                if str(variant.barcode_or_gtin).strip():
                                                                    metafields.gtin1 += f'{str(variant.barcode_or_gtin).strip()}, '

                                                            if str(metafields.product_size).strip()[-1] == ',': metafields.product_size = str(metafields.product_size).strip()[0:-1]
                                                            if str(metafields.gtin1).strip()[-1] == ',': metafields.gtin1 = str(metafields.gtin1).strip()[0:-1]
                                                        except: pass
                                                        
                                                        try: metafields.img_url = str(self.browser.find_element(By.XPATH, '//img[@class="ngxImageZoomThumbnail"]').get_attribute('src')).strip()
                                                        except: pass

                                                        # try:
                                                        #     if metafields.img_url:
                                                        #         src = ''
                                                        #         for _ in range(0, 10):
                                                        #             try:
                                                        #                 self.browser.find_element(By.CSS_SELECTOR, 'svg-icon[src$="360.svg"]').click()
                                                        #                 break
                                                        #             except: pass
                                                        #         for _ in range(0, 20):
                                                        #             try:
                                                        #                 src = self.browser.find_element(By.XPATH, '//div[@id="image-rotator"]/div[@class="window"]/div/img').get_attribute('src')
                                                        #                 break
                                                        #             except: sleep(0.2)
                                                        #         if src:
                                                        #             metafields.img_360_urls = src
                                                        #             for i in range(2, 9):
                                                        #                 metafields.img_360_urls = str(src).replace('_01.jpg', f'_0{i}.jpg')
                                                        #         sleep(0.3)
                                                        #         self.browser.find_element(By.CSS_SELECTOR, 'svg-icon[class="circle-cross-icon"]').click()
                                                        # except: pass    

                                                        # try:
                                                        #     for j in range(1, 9):
                                                        #         img_url_360 = f'https://nfseu.marcolin.com/Immagini/360/{str(brand["code"]).upper()}/{str(product.number).upper()}_{str(product.frame_code).upper()}/{str(product.number).upper()}_{str(product.frame_code).upper()}_0{j}.jpg'
                                                        #         metafields.img_360_urls.append(img_url_360)
                                                        # except Exception as e:
                                                        #     if self.DEBUG: print(f'Exception in getting 360 image urls: {e}')
                                                        #     else: pass
                                                        product.url = f'https://digitalhub.marcolin.com/shop/product-detail?idLine={str(brand["code"]).upper()}&idCode={product.number.upper().replace("/", "-")}&prod={product.number.upper().replace("/", "-")}{product.variants[0].title}{product.frame_code.upper().replace("/", "-")}'
                                                        product.metafields = metafields
                                                        
                                                    except: pass
                                                    self.data.append(product)
                                                    self.save_to_json(self.data)
                                                    
                                            except: pass
                                        self.close_last_tab() 
                                    
                                    if self.is_next_page(): self.move_to_next_page()
                                    else: break                           
                                print(f'Products scraped: {str(scraped_products)}')

                            self.close_last_tab()

            else: print(f'Failed to login \nURL: {self.URL}\nUsername: {str(username)}\nPassword: {str(password)}')
        except Exception as e:
            if self.DEBUG: print(f'Exception in scraper controller: {e}')
            else: pass
        finally: 
            self.browser.quit()

    def wait_until_browsing(self) -> None:
        while True:
            try:
                state = self.browser.execute_script('return document.readyState; ')
                if 'complete' == state: break
                else: sleep(0.2)
            except: pass

    def login(self, username: str, password: str) -> bool:
        login_flag = False
        try:
            if self.wait_until_element_found(20, 'xpath', '//input[@id="user-name"]'):
                self.browser.find_element(By.XPATH, '//input[@id="user-name"]').send_keys(username)
                self.browser.find_element(By.XPATH, '//input[@id="password"]').send_keys(password)
                try:
                    button = WebDriverWait(self.browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//button[@type="submit"]')))
                    button.click()

                    WebDriverWait(self.browser, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class*="welcome-msg my-5"] > h3')))
                    login_flag = True
                except Exception as e: print(e)
        except Exception as e:
            if self.DEBUG: print(f'Exception in login: {str(e)}')
            else: pass
        finally: return login_flag

    def wait_until_element_found(self, wait_value: int, type: str, value: str) -> bool:
        flag = False
        try:
            if type == 'id':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.ID, value)))
                flag = True
            elif type == 'xpath':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.XPATH, value)))
                flag = True
            elif type == 'css_selector':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.CSS_SELECTOR, value)))
                flag = True
            elif type == 'class_name':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.CLASS_NAME, value)))
                flag = True
            elif type == 'tag_name':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.TAG_NAME, value)))
                flag = True
        except: pass
        finally: return flag

    def get_brand_urls(self, brand: dict) -> str:
        brand_urls = []
        try:
            div_tags = self.browser.find_element(By.XPATH, '//div[@id="mCSB_1_container"]').find_elements(By.XPATH, './/div[@class="brand-box col-2"]')
            for div_tag in div_tags:
                if bool(brand['glasses_type']['sunglasses']):
                    href = div_tag.find_element(By.XPATH, ".//a[contains(text(), 'Sun')]").get_attribute('href')
                    if f'codeLine1={str(brand["code"]).strip().upper()}' in href:
                        brand_urls.append([f'{href}&limit=80', 'Sunglasses'])
                if bool(brand['glasses_type']['eyeglasses']):
                    href = div_tag.find_element(By.XPATH, ".//a[contains(text(), 'Optical')]").get_attribute('href')
                    if f'codeLine1={str(brand["code"]).strip().upper()}' in href:
                        brand_urls.append([f'{href}&limit=80', 'Eyeglasses'])
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_brand_url: {str(e)}')
            else: pass
        finally: return brand_urls

    def open_new_tab(self, url: str) -> None:
        # open category in new tab
        self.browser.execute_script('window.open("'+str(url)+'","_blank");')
        self.browser.switch_to.window(self.browser.window_handles[len(self.browser.window_handles) - 1])
        self.wait_until_browsing()
    
    def close_last_tab(self) -> None:
        self.browser.close()
        self.browser.switch_to.window(self.browser.window_handles[len(self.browser.window_handles) - 1])

    def get_all_products_from_page(self) -> list[dict]:
        products_on_page = []
        try:
            for _ in range(0, 30):
                products_on_page = []
                try:
                    for div_tag in self.browser.find_elements(By.XPATH, '//div[@class="row mt-4 list grid-divider"]/div'): 
                        ActionChains(self.browser).move_to_element(div_tag).perform()
                        product_url, product_number, product_brand, product_gender = '', '', '', ''
                        sizes = []

                        product_url = div_tag.find_element(By.TAG_NAME, 'a').get_attribute('href')
                        text = str(div_tag.find_element(By.XPATH, './/p[@class="model-name"]').text).strip()
                        product_number = str(text.split(' ')[0]).strip()
                        product_name = str(text).replace(product_number, '').strip()
                        product_brand = str(div_tag.find_element(By.XPATH, '//div[@class="line-name d-flex justify-content-between"]/p').text).strip()
                        # if str('WEB').strip().lower() == str(product_brand).strip().lower():
                        #     try:
                        #         if str(product_number[:2]).strip().lower() != str('WB').strip().lower():
                        #             new_number = quote(product_number)
                        #             new_url = f'https://digitalhub.marcolin.com/shop/products?searchText={new_number}'
                        #             new_brand_name = self.search_for_brand_name(new_url, product_number)
                        #             if new_brand_name and str(new_brand_name).strip().lower() != str(product_brand).strip().lower(): product_brand = new_brand_name
                        #     except Exception as e: 
                        #         if self.DEBUG: print(f'Exception as in getting new brand name: {str(e)}')
                        #         else: pass

                        try: product_gender = str(div_tag.find_element(By.XPATH, './/div[@class="info"]/p[contains(text(), "Gender")]').text).replace('Gender:', '').strip()
                        except: pass
                        
                        try:
                            size_value = str(div_tag.find_element(By.XPATH, './/div[@class="info"]/p[contains(text(), "Size")]').text).replace('Size:', '').strip()
                            
                            if ',' in size_value:
                                for value in size_value.split(','):
                                    sizes.append(str(value).strip())
                            else: sizes.append(size_value)
                        except: pass
                        
                        json_data = {
                            'number': product_number,
                            'name': product_name,
                            'brand': product_brand,
                            'gender': product_gender,
                            'url': product_url,
                            'sizes': sizes
                        }
                        if json_data not in products_on_page: products_on_page.append(json_data)
                    break
                except: sleep(0.3)
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_all_products_from_page: {str(e)}')
            else: pass
        finally: return products_on_page

    def is_next_page(self) -> bool:
        next_page_flag = False
        try:
            next_span_style = self.browser.find_element(By.XPATH, '//span[@class="next"]').get_attribute('style')
            if ': hidden;' not in next_span_style: next_page_flag = True
        except Exception as e:
            if self.DEBUG: print(f'Exception in is_next_page: {str(e)}')
            else: pass
        finally: return next_page_flag

    def move_to_next_page(self) -> None:
        try:
            current_page_number = str(self.browser.find_element(By.XPATH, '//span[@class="current"]').text).strip()
            next_page_span = self.browser.find_element(By.XPATH, '//span[@class="next"]')
            # ActionChains(self.browser).move_to_element(next_page_span).perform()
            ActionChains(self.browser).move_to_element(next_page_span).click().perform()
            self.wait_for_next_page_to_load(current_page_number)
        except Exception as e:
            if self.DEBUG: print(f'Exception in move_to_next_page: {str(e)}')
            else: pass
    
    def wait_for_next_page_to_load(self, current_page_number: str) -> None:
        for _ in range(0, 100):
            try:
                next_page_number = str(self.browser.find_element(By.XPATH, '//span[@class="current"]').text).strip()
                if int(next_page_number) > int(current_page_number): 
                    for _ in range(0, 30):
                        try:
                            for div_tag in self.browser.find_elements(By.XPATH, '//div[@class="row mt-4 list grid-divider"]/div'):
                                div_tag.find_element(By.XPATH, './/p[@class="model-name"]').text
                            break
                        except: sleep(0.3)
                    break
            except: sleep(0.3)
     
    def move_to_first_variant(self) -> None:
        while True:
            try:
                elements = self.browser.find_elements(By.XPATH, '//span[@class="n-arrow prev slick-arrow"]')
                if len(elements) == 2:
                    ActionChains(self.browser).move_to_element(elements[0]).click().perform()
                    sleep(0.3)
                else: break
            except: pass
    
    def select_variant_image(self, divs: list) -> None:
        for _ in range(0, 5):
            try:
                ActionChains(self.browser).move_to_element(divs).click().perform()
                sleep(0.3)
                break
            except:
                try:
                    elements = self.browser.find_elements(By.XPATH, '//span[@class="n-arrow next slick-arrow"]')
                    if len(elements) == 2:
                        ActionChains(self.browser).move_to_element(elements[0]).click().perform()
                        sleep(0.3)
                except Exception as e: 
                    if self.DEBUG: print(str(e))
                    else: pass

    def click_to_make_price_visible(self) -> None:
        try:
            if not self.is_price_visible():
                element = self.browser.find_element(By.XPATH, '//svg-icon[@class="eye-icon"]')
                ActionChains(self.browser).move_to_element(element).click().perform()
                sleep(0.4)

                for li in self.browser.find_elements(By.XPATH, '//ul[@aria-labelledby="basic-link"]/li'):
                    if str('Cost and SRP').strip().lower() in str(li.text).strip().lower():
                        ActionChains(self.browser).move_to_element(li).click().perform()
                        self.wait_until_price_is_shown()
        except Exception as e:
            if self.DEBUG: print(f'Exception in click_to_make_price_visible: {str(e)}')
            else: pass

    def is_price_visible(self) -> bool:
        try:
            for td in self.browser.find_elements(By.XPATH, '//table[@class="table table-borderless"]/tbody/tr/td'):
                if str('Suggested Retail Price').strip().lower() in str(td.text).strip().lower(): return True
            return False
        except: return False

    def wait_until_price_is_shown(self) -> bool:
        flag = False
        for _ in range(0, 30):
            try:
                tds_label = self.browser.find_elements(By.XPATH, '//table[@class="table table-borderless"]/tbody/tr/td')
                for i in range(0, len(tds_label)):
                    if str('Suggested Retail Price').strip().lower() in str(tds_label[i].text).strip().lower(): 
                        flag = True
                        break
                if flag: break
            except: sleep(0.3)
            finally: return flag

    def get_size_price_status(self):
        size_titles, wholesale_prices, listing_prices, availability, gtin = [], [], [], [], []
        sizes = []
        caliber, rod, bridge = '', '', ''
        try:
            trs = self.browser.find_elements(By.XPATH, '//table[@class="table table-borderless"]/tbody/tr')
            
            for j in range(1, len(trs)):
                tr = trs[j].find_element(By.XPATH, '//table[@class="table table-borderless inner-table"]/tr')
                
                # tds_value = tr.find_elements_by_tag_name('td')
                try:
                    value = str(tr.find_elements(By.XPATH, ".//td[contains(text(), '€')]")[0].text).replace('€', '').strip()
                    value = f"{str(value[0:-3]).strip().replace(',', '').strip()}.00"
                    wholesale_prices.append(value)
                except Exception as e: 
                    if self.DEBUG: print(f'Exception in wholesale_price: {str(e)}')
                    else: pass

                try:
                    value = str(tr.find_elements(By.XPATH, ".//td[contains(text(), '€')]")[1].text).replace('€', '').strip()
                    value = f"{str(value[0:-3]).strip().replace(',', '').strip()}.00"
                    listing_prices.append(value)
                except Exception as e: 
                    if self.DEBUG: print(f'Exception in listing_prices: {str(e)}')
                    else: pass
                
            for tr in self.browser.find_elements(By.XPATH, '//table[@class="table table-borderless inner-table"]/tr'):
                tds_value = tr.find_elements(By.TAG_NAME, 'td')
                try:
                    caliber, rod, bridge = str(tds_value[0].text).strip(), str(tds_value[1].text).strip(), str(tds_value[2].text).strip()
                    # print(len(tds_value), f'{caliber}-{rod}-{bridge}')
                    size_titles.append(caliber)
                    sizes.append(f'{caliber}-{rod}-{bridge}')
                    # if not metafields.product_size: metafields.product_size =f'{caliber}-{rod}-{bridge}'
                    caliber, rod, bridge = '', '', ''
                except Exception as e: 
                    if self.DEBUG: print(f'Exception in product size: {str(e)}')
                    else: pass

                
                try:
                    span_tag_class = tr.find_element(By.CSS_SELECTOR, 'span[class^="availability"]').get_attribute('class')
                    # if str('a-0').strip().lower() in  str(span_tag_class).strip().lower(): availability.append('Draft')
                    # if str('a-1').strip().lower() in  str(span_tag_class).strip().lower(): availability.append('Draft')
                    if str('a-2').strip().lower() in  str(span_tag_class).strip().lower(): availability.append('Active')
                    else: availability.append('Draft')
                    # elif str('a-3').strip().lower() in  str(span_tag_class).strip().lower(): availability.append('Draft')
                except: 
                    availability.append('Not Available')

            
            # for j in range(1, len(trs)):
            try:
                for button in self.browser.find_elements(By.XPATH, '//svg-icon[@class="arrow-icon"]'):
                    ActionChains(self.browser).move_to_element(button).perform()
                    button.click()
                    sleep(0.5)
                    for _ in range(0, 20):
                        try:
                            tags = self.browser.find_elements(By.XPATH, '//table[@class="table table-borderless inner-table open-shadow"]/tr[@class="d-flex drawer"]')
                            for tag in tags:    
                                g = str(tag.find_element(By.CSS_SELECTOR, 'div[class$="ean-detail"] > p > span').text).strip()
                                if g: 
                                    if g not in gtin: gtin.append(g)
                                else: gtin.append('')
                            # close_element = self.browser.find_element(By.XPATH, '//table[@class="table table-borderless inner-table open-shadow"]/tr[@class="d-flex"]').find_element(By.XPATH, '//svg-icon[@class="arrow-icon"]')
                            # ActionChains(self.browser).move_to_element(close_element).perform()
                            # close_element.click()
                            # sleep(0.3)
                            break
                        except Exception as e:pass
            except Exception as e: pass
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_size_price_status: {str(e)}')
            else: pass
        finally: return size_titles, wholesale_prices, listing_prices, availability, gtin, sizes

    def save_to_json(self, products: list[Product]):
        try:
            json_products = []
            for product in products:
                json_varinats = []
                for index, variant in enumerate(product.variants):
                    json_varinat = {
                        'position': (index + 1), 
                        'title': variant.title, 
                        'sku': variant.sku, 
                        'inventory_quantity': variant.inventory_quantity,
                        'found_status': variant.found_status,
                        'wholesale_price': variant.wholesale_price,
                        'listing_price': variant.listing_price, 
                        'barcode_or_gtin': variant.barcode_or_gtin,
                        'size': variant.size,
                        'weight': variant.weight
                    }
                    json_varinats.append(json_varinat)
                json_product = {
                    'brand': product.brand, 
                    'number': product.number, 
                    'name': product.name, 
                    'frame_code': product.frame_code, 
                    'frame_color': product.frame_color, 
                    'lens_code': product.lens_code, 
                    'lens_color': product.lens_color, 
                    'status': product.status, 
                    'type': product.type, 
                    'url': product.url, 
                    'metafields': [
                        { 'key': 'for_who', 'value': product.metafields.for_who },
                        { 'key': 'product_size', 'value': product.metafields.product_size }, 
                        { 'key': 'lens_material', 'value': product.metafields.lens_material }, 
                        { 'key': 'lens_technology', 'value': product.metafields.lens_technology }, 
                        { 'key': 'frame_material', 'value': product.metafields.frame_material }, 
                        { 'key': 'frame_shape', 'value': product.metafields.frame_shape },
                        { 'key': 'gtin1', 'value': product.metafields.gtin1 }, 
                        { 'key': 'img_url', 'value': product.metafields.img_url }
                    ],
                    'variants': json_varinats
                }
                json_products.append(json_product)
            
           
            with open(self.result_filename, 'w') as f: json.dump(json_products, f)
            
        except Exception as e:
            if self.DEBUG: print(f'Exception in save_to_json: {e}')
            else: pass


def read_data_from_json_file(DEBUG, result_filename: str):
    data = []
    try:
        files = glob.glob(result_filename)
        if files:
            f = open(files[-1])
            json_data = json.loads(f.read())
            products = []

            for json_d in json_data:
                number, frame_code, brand, img_url, frame_color, lens_color = '', '', '', '', '', ''
                # product = Product()
                brand = json_d['brand']
                number = str(json_d['number']).strip().upper()
                if '/' in number: number = number.replace('/', '-').strip()
                # product.name = str(json_d['name']).strip().upper()
                frame_code = str(json_d['frame_code']).strip().upper()
                if '/' in frame_code: frame_code = frame_code.replace('/', '-').strip()
                frame_color = str(json_d['frame_color']).strip().title()
                # lens_code = str(json_d['lens_code']).strip().upper()
                lens_color = str(json_d['lens_color']).strip().title()
                # product.status = str(json_d['status']).strip().lower()
                # product.type = str(json_d['type']).strip().title()
                # product.url = str(json_d['url']).strip()
                # metafields = Metafields()
                
                for json_metafiels in json_d['metafields']:
                    # if json_metafiels['key'] == 'for_who':metafields.for_who = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'product_size':metafields.product_size = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'activity':metafields.activity = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'lens_material':metafields.lens_material = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'graduabile':metafields.graduabile = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'interest':metafields.interest = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'lens_technology':metafields.lens_technology = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'frame_material':metafields.frame_material = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'frame_shape':metafields.frame_shape = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'gtin1':metafields.gtin1 = str(json_metafiels['value']).strip().title()
                    if json_metafiels['key'] == 'img_url':img_url = str(json_metafiels['value']).strip()
                    # elif json_metafiels['key'] == 'img_360_urls':
                    #     value = str(json_metafiels['value']).strip()
                    #     if '[' in value: value = str(value).replace('[', '').strip()
                    #     if ']' in value: value = str(value).replace(']', '').strip()
                    #     if "'" in value: value = str(value).replace("'", '').strip()
                    #     for v in value.split(','):
                    #         metafields.img_360_urls = str(v).strip()
                # product.metafields = metafields
                for json_variant in json_d['variants']:
                    sku, price = '', ''
                    # variant = Variant()
                    # variant.position = json_variant['position']
                    # variant.title = str(json_variant['title']).strip()
                    sku = str(json_variant['sku']).strip().upper()
                    if '/' in sku: sku = sku.replace('/', '-').strip()
                    # variant.inventory_quantity = json_variant['inventory_quantity']
                    # variant.found_status = json_variant['found_status']
                    wholesale_price = str(json_variant['wholesale_price']).strip()
                    listing_price = str(json_variant['listing_price']).strip()
                    # variant.barcode_or_gtin = str(json_variant['barcode_or_gtin']).strip()
                    # variant.size = str(json_variant['size']).strip()
                    # variant.weight = str(json_variant['weight']).strip()
                    # product.variants = variant

                    image_attachment = download_image(img_url)
                    if image_attachment:
                        with open(f'Images/{sku}.jpg', 'wb') as f: f.write(image_attachment)
                        crop_downloaded_image(f'Images/{sku}.jpg')
                    data.append([number, frame_code, frame_color, lens_color, brand, sku, wholesale_price, listing_price])
    except Exception as e:
        if DEBUG: print(f'Exception in read_data_from_json_file: {e}')
        else: pass
    finally: return data

def download_image(url):
    image_attachment = ''
    try:
        headers = {
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'accept-Encoding': 'gzip, deflate, br',
            'accept-Language': 'en-US,en;q=0.9',
            'cache-Control': 'max-age=0',
            'sec-ch-ua': '"Google Chrome";v="95", "Chromium";v="95", ";Not A Brand";v="99"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'document',
            'sec-fetch-mode': 'navigate',
            'sec-fetch-site': 'none',
            'Sec-Fetch-User': '?1',
            'upgrade-insecure-requests': '1',
        }
        counter = 0
        while True:
            try:
                response = requests.get(url=url, headers=headers, timeout=20)
                # print(response.status_code)
                if response.status_code == 200:
                    # image_attachment = base64.b64encode(response.content)
                    image_attachment = response.content
                    break
                else: print(f'{response.status_code} found for downloading image')
            except: sleep(0.3)
            counter += 1
            if counter == 10: break
    except Exception as e: print(f'Exception in download_image: {str(e)}')
    finally: return image_attachment

def crop_downloaded_image(filename):
    try:
        im = Image.open(filename)
        width, height = im.size   # Get dimensions
        new_width = 1680
        new_height = 1020
        if width > new_width and height > new_height:
            left = (width - new_width)/2
            top = (height - new_height)/2
            right = (width + new_width)/2
            bottom = (height + new_height)/2
            im = im.crop((left, top, right, bottom))
            im.save(filename)
    except Exception as e: print(f'Exception in crop_downloaded_image: {e}')

def saving_picture_in_excel(data: list):
    workbook = Workbook()
    worksheet = workbook.active

    worksheet.cell(row=1, column=1, value='Model Code')
    worksheet.cell(row=1, column=2, value='Lens Code')
    worksheet.cell(row=1, column=3, value='Color Frame')
    worksheet.cell(row=1, column=4, value='Color Lens')
    worksheet.cell(row=1, column=5, value='Brand')
    worksheet.cell(row=1, column=6, value='SKU')
    worksheet.cell(row=1, column=7, value='Wholesale Price')
    worksheet.cell(row=1, column=8, value='Listing Price')
    worksheet.cell(row=1, column=9, value="Image")

    for index, d in enumerate(data):
        new_index = index + 2

        worksheet.cell(row=new_index, column=1, value=d[0])
        worksheet.cell(row=new_index, column=2, value=d[1])
        worksheet.cell(row=new_index, column=3, value=d[2])
        worksheet.cell(row=new_index, column=4, value=d[3])
        worksheet.cell(row=new_index, column=5, value=d[4])
        worksheet.cell(row=new_index, column=6, value=d[5])
        worksheet.cell(row=new_index, column=7, value=d[6])
        worksheet.cell(row=new_index, column=8, value=d[7])

        image = f'Images/{d[-3]}.jpg'
        if os.path.exists(image):
            im = Image.open(image)
            width, height = im.size
            worksheet.row_dimensions[new_index].height = height
            worksheet.add_image(Imag(image), anchor='I'+str(new_index))
            # col_letter = get_column_letter(9)
            # worksheet.column_dimensions[col_letter].width = width
        # print(index, image)

    workbook.save('Digitalhub Results.xlsx')

DEBUG = True
try:
    pathofpyfolder = os.path.realpath(sys.argv[0])
    # get path of Exe folder
    path = pathofpyfolder.replace(pathofpyfolder.split('\\')[-1], '')
    # download chromedriver.exe with same version and get its path
    if os.path.exists('chromedriver.exe'): os.remove('chromedriver.exe')
    if os.path.exists('Digitalhub Results.xlsx'): os.remove('Digitalhub Results.xlsx')

    chromedriver_autoinstaller.install(path)
    if '.exe' in pathofpyfolder.split('\\')[-1]: DEBUG = False
    
    f = open('Digitalhub start.json')
    json_data = json.loads(f.read())
    f.close()

    brands = json_data['brands']

    
    f = open('requirements/digitalhub.json')
    data = json.loads(f.read())
    f.close()
    url = data['url']
    username = data['username']
    password = data['password']
    
    result_filename = 'requirements/Digitalhub Results.json'
    Digitalhub_Scraper(DEBUG, result_filename).controller(brands, url, username, password)
    
    for filename in glob.glob('Images/*'): os.remove(filename)
    data = read_data_from_json_file(DEBUG, result_filename)
    os.remove(result_filename)

    saving_picture_in_excel(data)

except Exception as e:
    if DEBUG: print('Exception: '+str(e))
    else: pass
