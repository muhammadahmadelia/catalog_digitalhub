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
from models.store import Store
from models.brand import Brand
from models.product import Product
from models.variant import Variant
from models.metafields import Metafields
import glob
import requests
from datetime import datetime
import threading

from openpyxl import Workbook
from openpyxl.drawing.image import Image as Imag
from openpyxl.utils import get_column_letter
from PIL import Image

from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager


class myScrapingThread(threading.Thread):
    def __init__(self, threadID: int, name: str, obj, username: str, brand: str, brand_code: str, product_number: str, glasses_type: str, headers: dict) -> None:
        threading.Thread.__init__(self)
        self.threadID = threadID
        self.name = name
        self.username = username
        self.brand = brand
        self.brand_code = brand_code
        self.product_number = product_number
        self.glasses_type = glasses_type
        self.headers = headers
        self.obj = obj
        self.status = 'in progress'
        pass

    def run(self):
        self.obj.scrape_product(self.username, self.brand, self.brand_code, self.product_number, self.glasses_type, self.headers)
        self.status = 'completed'

    def active_threads(self):
        return threading.activeCount()

class Digitalhub_Scraper:
    def __init__(self, DEBUG: bool, result_filename: str, logs_filename: str) -> None:
        self.DEBUG = DEBUG
        self.data = []
        self.result_filename = result_filename
        self.logs_filename = logs_filename
        self.thread_list = []
        self.thread_counter = 0
        self.chrome_options = Options()
        self.chrome_options.add_argument('--disable-infobars')
        self.chrome_options.add_argument("--start-maximized")
        self.chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        self.args = ["hide_console", ]
        # self.browser = webdriver.Chrome(options=self.chrome_options, service_args=self.args)
        self.browser = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=self.chrome_options)
        pass

    def controller(self, store: Store, brands_with_types: list[dict]) -> None:
        try:
            cookies, fs_token = '', ''

            self.browser.get(store.link)
            self.wait_until_browsing()

            if self.login(store.username, store.password):
                self.browser.get('https://digitalhub.marcolin.com/shop')
                self.wait_until_browsing()

                if self.wait_until_element_found(20, 'xpath', '//div[@id="mCSB_1_container"]'):
                    
                    for brand_with_type in brands_with_types:
                        brand: str = brand_with_type['brand']
                        brand_code: str = str(brand_with_type['code']).strip()
                        print(f'Brand: {brand}')

                        for glasses_type in brand_with_type['glasses_type']:

                            brand_url = self.get_brand_url(brand, brand_code, glasses_type)
                            self.open_new_tab(brand_url)
                            self.wait_until_browsing()
                            start_time = datetime.now()

                            if self.wait_until_element_found(90, 'xpath', '//div[@class="row mt-4 list grid-divider"]/div'):
                                total_products = self.get_total_products()
                                scraped_products = 0
                                
                                print(f'Type: {glasses_type} | Total products: {total_products}')
                                print(f'Start Time: {start_time.strftime("%A, %d %b %Y %I:%M:%S %p")}')

                                self.printProgressBar(scraped_products, total_products, prefix = 'Progress:', suffix = 'Complete', length = 50)
                                while True:

                                    for product_data in self.get_all_products_from_page():
                                        product_number = str(product_data['number']).strip().upper()
                                        product_url = str(product_data['url']).strip()
                                        

                                        if not cookies: cookies = self.get_cookies()
                                        if not fs_token: fs_token = self.get_fs_token()
                                        headers = self.get_headers(fs_token, cookies, product_url)

                                        # self.scrape_product(store.username, brand, product_number, glasses_type, headers)
                                        self.create_thread(store.username, brand, brand_code, product_number, glasses_type, headers)
                                        if self.thread_counter >= 50: 
                                            self.wait_for_thread_list_to_complete()
                                            self.save_to_json(self.data)
                                        scraped_products += 1

                                        self.printProgressBar(scraped_products, total_products, prefix = 'Progress:', suffix = 'Complete', length = 50)
                                    
                                    if self.is_next_page(): self.move_to_next_page()
                                    else: break

                            self.wait_for_thread_list_to_complete()
                            self.save_to_json(self.data)
                            end_time = datetime.now()
                            
                            print(f'End Time: {end_time.strftime("%A, %d %b %Y %I:%M:%S %p")}')
                            print('Duration: {}\n'.format(end_time - start_time))
                            
                            self.close_last_tab()

            else: print(f'Failed to login \nURL: {store.link}\nUsername: {str(store.username)}\nPassword: {str(store.password)}')
        except Exception as e:
            if self.DEBUG: print(f'Exception in Digitalhub_Scraper controller: {e}')
            self.print_logs(f'Exception in Digitalhub_Scraper controller: {e}')
        finally: 
            self.browser.quit()
            self.wait_for_thread_list_to_complete()
            self.save_to_json(self.data)

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
                except Exception as e: 
                    self.print_logs(str(e))
                    if self.DEBUG: print(str(e))
                    else: pass
        except Exception as e:
            self.print_logs(f'Exception in login: {str(e)}')
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

    def get_brand_url(self, brand: str, brand_code: str, glasses_type: str) -> str:
        brand_url = ''
        try:
            div_tags = self.browser.find_element(By.XPATH, '//div[@id="mCSB_1_container"]').find_elements(By.XPATH, './/div[@class="brand-box col-2"]')
            xpath_glasses_type = ''
            if glasses_type == 'Sunglasses':
                xpath_glasses_type = ".//a[contains(text(), 'Sun')]"
            elif glasses_type == 'Eyeglasses':
                xpath_glasses_type = ".//a[contains(text(), 'Optical')]"
            for div_tag in div_tags:
                href = div_tag.find_element(By.XPATH, xpath_glasses_type).get_attribute('href')
                if f'codeLine1={str(brand_code).strip().upper()}' in href:
                    brand_url = f'{href}&limit=80'

        except Exception as e:
            self.print_logs(f'Exception in get_brand_url: {str(e)}')
            if self.DEBUG: print(f'Exception in get_brand_url: {str(e)}')
            else: pass
        finally: return brand_url

    def open_new_tab(self, url: str) -> None:
        # open category in new tab
        self.browser.execute_script('window.open("'+str(url)+'","_blank");')
        self.browser.switch_to.window(self.browser.window_handles[len(self.browser.window_handles) - 1])
        self.wait_until_browsing()
    
    def close_last_tab(self) -> None:
        self.browser.close()
        self.browser.switch_to.window(self.browser.window_handles[len(self.browser.window_handles) - 1])
    
    def is_next_page(self) -> bool:
        next_page_flag = False
        try:
            next_span_style = self.browser.find_element(By.XPATH, '//span[@class="next"]').get_attribute('style')
            if ': hidden;' not in next_span_style: next_page_flag = True
        except Exception as e:
            self.print_logs(f'Exception in is_next_page: {str(e)}')
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
            self.print_logs(f'Exception in move_to_next_page: {str(e)}')
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
 
    def get_total_products(self) -> int:
        total_products = 0
        try:
            total_products = int(str(self.browser.find_element(By.XPATH, '//div[@class="row mt-4 results"]/div').text).strip().split(' ')[0])
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_total_products: {e}')
            self.print_logs(f'Exception in get_total_products: {e}')
        finally: return total_products

    def get_all_products_from_page(self) -> list[dict]:
        products_on_page = []
        try:
            for _ in range(0, 30):
                products_on_page = []
                try:
                    for div_tag in self.browser.find_elements(By.XPATH, '//div[@class="row mt-4 list grid-divider"]/div'): 
                        ActionChains(self.browser).move_to_element(div_tag).perform()
                        product_url, product_number = '', ''

                        product_url = div_tag.find_element(By.TAG_NAME, 'a').get_attribute('href')
                        text = str(div_tag.find_element(By.XPATH, './/p[@class="model-name"]').text).strip()
                        product_number = str(text.split(' ')[0]).strip()
                        
                        json_data = {
                            'number': product_number,
                            'url': product_url
                        }
                        if json_data not in products_on_page: products_on_page.append(json_data)
                    break
                except: sleep(0.3)
        except Exception as e:
            self.print_logs(f'Exception in get_all_products_from_page: {str(e)}')
            if self.DEBUG: print(f'Exception in get_all_products_from_page: {str(e)}')
            else: pass
        finally: return products_on_page

    def scrape_product(self, username: str, brand: str, brand_code: str, product_number: str, glasses_type: str, headers: dict) -> None:
        try:
            url = f'https://digitalhub.marcolin.com/api/model?codeSalesOrg=IA01&soldCode={username}&shipCode=&idLine={str(brand_code).upper()}&idCode={product_number}&spareParts=null'

            response = self.make_request(url, headers)
            if response.status_code == 200:
                json_data = json.loads(response.text)
                product_name = str(json_data['data']['name']).strip().replace(str(product_number).strip().upper(), '').strip()
                frame_codes = []

                for json_product in json_data['data']['products']:
                    if str(json_product['colorCode']).strip().upper() not in frame_codes:
                        product = Product()

                        product.url = str(headers['Referer']).strip().split('&prod=')[0] + f'&prod={json_product["idCode"]}'
                        product.brand = brand
                        product.number = str(json_product['codLevel1']).strip().upper().replace('-', '/')
                        product.name = product_name
                        product.frame_code = str(json_product['colorCode']).strip().upper()
                        frame_codes.append(product.frame_code)
                        
                        colorDescription = str(json_product['colorDescription']).strip().split(' - ')[-1].strip().split(' / ')
                        if len(colorDescription) == 1: product.frame_color = colorDescription[0].strip().title()
                        elif len(colorDescription) == 2:
                            product.frame_color = colorDescription[0].strip().title()
                            product.lens_color = colorDescription[1].strip().title()
                        
                        product.type = glasses_type
                        product.status = 'active'
                        barcodes, sizes = [], []
                        
                        for json_product2 in json_data['data']['products']:
                            if str(json_product2['colorCode']).strip().upper() == product.frame_code:
                                variant = Variant()
                                variant.position = len(product.variants) + 1
                                if 'sizeDescription' in json_product2:
                                    variant.title = str(json_product2['sizeDescription']).strip()
                                variant.sku = f'{product.number} {product.frame_code} {variant.title}'
                                
                                if 'aux' in json_product2:
                                    if 'availabilityColor' in json_product2['aux']:
                                        if json_product2['aux']['availabilityColor'] == 2: variant.inventory_quantity = 1
                                        else: variant.inventory_quantity = 0
                                    else: variant.inventory_quantity = 0
                                else: variant.inventory_quantity = 0
                                variant.found_status = 1
                                if 'price' in json_product2:
                                    variant.wholesale_price = format(int(json_product2['price']), '.2f')
                                if 'publicPrice' in json_product2:
                                    variant.listing_price = format(int(json_product2['publicPrice']), '.2f')
                                if 'barcode' in json_product2:
                                    variant.barcode_or_gtin = str(json_product2['barcode']).strip()
                                if 'aux' in json_product2:
                                    if 'rodLength' in json_product2['aux'] and 'noseLength' in json_product2['aux']:
                                        variant.size = f'{variant.title}-{json_product2["aux"]["rodLength"]}-{json_product2["aux"]["noseLength"]}'
                                product.variants = variant
                                barcodes.append(variant.barcode_or_gtin)
                                sizes.append(variant.size)

                        
                        metafields = Metafields()

                        try:
                            metafields.for_who = str(json_product['aux']['genderDesc']).strip().title()
                            if metafields.for_who == 'Male': metafields.for_who = 'Men'
                            elif metafields.for_who == 'Female': metafields.for_who = 'Women'
                        except: pass
                        
                        try: metafields.product_size = ', '.join(sizes)
                        except: pass
                        try: metafields.lens_technology = str(json_product['aux']['typeLensesDesc']).strip().title()
                        except: pass
                        try: metafields.frame_material = str(json_product['aux']['productGroupDesc']).strip().title()
                        except: pass
                        try: metafields.frame_shape = str(json_product['aux']['formDesc']).strip().title()
                        except: pass
                        try: metafields.gtin1 = ', '.join(barcodes)
                        except: pass
                        
                        try: metafields.img_url = str(json_product['image']).strip().replace('\/', '\\')
                        except: pass
                        # try: 
                        #     for image360 in json_product['images360']:
                        #         metafields.img_360_urls = str(image360).strip().replace('\/', '/')
                        # except: pass

                        product.metafields = metafields

                        self.data.append(product)
            else: self.print_logs(f'{response.status_code} for {url}')
        except Exception as e:
            if self.DEBUG: print(f'Exception in scrape_product_data: {e}')
            self.print_logs(f'Exception in scrape_product_data: {e}')

    def get_fs_token(self) -> str:
        fs_token = ''
        try:
            fs_token = self.browser.execute_script('return window.localStorage.getItem(arguments[0]);', 'fs_token')
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_fs_token: {e}')
            self.print_logs(f'Exception in get_fs_token: {e}')
        finally: return fs_token

    def get_cookies(self) -> str:
        cookies = ''
        try:
            browser_cookies = self.browser.get_cookies()
            for browser_cookie in browser_cookies:
                if browser_cookie["name"] == 'php-console-server':
                    cookies = f'{browser_cookie["name"]}={browser_cookie["value"]}; _gat_UA-153573784-1=1; {cookies}'
                else: cookies = f'{browser_cookie["name"]}={browser_cookie["value"]}; {cookies}'
            cookies = cookies.strip()[:-1]
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_cookies: {e}')
            self.print_logs(f'Exception in get_cookies: {e}')
        finally: return cookies

    def get_headers(self, fs_token: str, cookies: str, referer_url: str) -> dict:
        return {
            'Accept': 'application/json, text/plain, */*',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'en-US,en;q=0.9',
            'Authorization': f'Bearer {fs_token}',
            'Connection': 'keep-alive',
            'Cookie': cookies,
            'Host': 'digitalhub.marcolin.com',
            'Referer': referer_url,
            'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'
        }
    
    def make_request(self, url, headers):
        response = ''
        for _ in range(0, 10):
            try:
                response = requests.get(url=url, headers=headers)
                if response.status_code == 200: break
            except: sleep(0.2)
        return response

    def save_to_json(self, products: list[Product]) -> None:
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
                        'listing_price': variant.listing_price, 
                        'wholesale_price': variant.wholesale_price,
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
                        { 'key': 'img_url', 'value': product.metafields.img_url },
                        { 'key': 'img_360_urls', 'value': product.metafields.img_360_urls }
                    ],
                    'variants': json_varinats
                }
                json_products.append(json_product)
            
           
            with open(self.result_filename, 'w') as f: json.dump(json_products, f)
            
        except Exception as e:
            if self.DEBUG: print(f'Exception in save_to_json: {e}')
            self.print_logs(f'Exception in save_to_json: {e}')
    
    # print logs to the log file
    def print_logs(self, log: str) -> None:
        try:
            with open(self.logs_filename, 'a') as f:
                f.write(f'\n{log}')
        except: pass

    def printProgressBar(self, iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r") -> None:
        """
        Call in a loop to create terminal progress bar
        @params:
            iteration   - Required  : current iteration (Int)
            total       - Required  : total iterations (Int)
            prefix      - Optional  : prefix string (Str)
            suffix      - Optional  : suffix string (Str)
            decimals    - Optional  : positive number of decimals in percent complete (Int)
            length      - Optional  : character length of bar (Int)
            fill        - Optional  : bar fill character (Str)
            printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
        """
        percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
        filledLength = int(length * iteration // total)
        bar = fill * filledLength + '-' * (length - filledLength)
        print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
        # Print New Line on Complete
        if iteration == total: 
            print()

    def create_thread(self, username: str, brand: str, brand_code: str, product_number: str, glasses_type: str, headers: dict) -> None:
        thread_name = "Thread-"+str(self.thread_counter)
        self.thread_list.append(myScrapingThread(self.thread_counter, thread_name, self, username, brand, brand_code, product_number, glasses_type, headers))
        self.thread_list[self.thread_counter].start()
        self.thread_counter += 1

    def is_thread_list_complted(self) -> bool:
        for obj in self.thread_list:
            if obj.status == "in progress":
                return False
        return True

    def wait_for_thread_list_to_complete(self) -> None:
        while True:
            result = self.is_thread_list_complted()
            if result: 
                self.thread_counter = 0
                self.thread_list.clear()
                break
            else: sleep(1)


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
    # if os.path.exists('chromedriver.exe'): os.remove('chromedriver.exe')
    if os.path.exists('Digitalhub Results.xlsx'): os.remove('Digitalhub Results.xlsx')

    # chromedriver_autoinstaller.install(path)
    if '.exe' in pathofpyfolder.split('\\')[-1]: DEBUG = False
    
    f = open('Digitalhub start.json')
    json_data = json.loads(f.read())
    f.close()

    brands = json_data['brands']

    
    f = open('requirements/digitalhub.json')
    data = json.loads(f.read())
    f.close()

    store = Store()
    store.link = data['url']
    store.username = data['username']
    store.password = data['password']
    store.login_flag = True

    result_filename = 'requirements/Digitalhub Results.json'

    if not os.path.exists('Logs'): os.makedirs('Logs')

    log_files = glob.glob('Logs/*.txt')
    if len(log_files) > 5:
        oldest_file = min(log_files, key=os.path.getctime)
        os.remove(oldest_file)
        log_files = glob.glob('Logs/*.txt')

    scrape_time = datetime.now().strftime('%d-%m-%Y %H-%M-%S')
    logs_filename = f'Logs/Logs {scrape_time}.txt'
    
    Digitalhub_Scraper(DEBUG, result_filename, logs_filename).controller(store, brands)
    
    for filename in glob.glob('Images/*'): os.remove(filename)
    data = read_data_from_json_file(DEBUG, result_filename)
    os.remove(result_filename)

    saving_picture_in_excel(data)
except Exception as e:
    if DEBUG: print('Exception: '+str(e))
    else: pass
