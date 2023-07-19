from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.select import Select
import undetected_chromedriver as uc
import time
import os
import re
from datetime import datetime
import pandas as pd
import warnings
import sys
import xlsxwriter
from multiprocessing import freeze_support
import shutil
warnings.filterwarnings('ignore')

def initialize_bot():

    # Setting up chrome driver for the bot
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--headless')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    # configuring the driver
    driver = webdriver.Chrome(options=chrome_options, service=chrome_service)
    ver = int(driver.capabilities['chrome']['chromedriverVersion'].split('.')[0])
    driver.quit()
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--lang=en")
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument('--headless=new')
    chrome_options.page_load_strategy = 'normal'
    driver = uc.Chrome(version_main = ver, options=chrome_options) 
    driver.set_window_size(1920, 1080)
    driver.maximize_window()
    driver.set_page_load_timeout(200)

    return driver

def scrape_JHC(driver, output1, page, cat, settings):

    print('-'*75)
    print(f'Scraping products Links from: {page}')
    stamp = datetime.now().strftime("%d_%m_%Y")
    prod_limit = settings['Product Limit']

    # getting the products list
    links = []
    driver.get(page)

    # handling lazy loading
    try:
        total_height = driver.execute_script("return document.body.scrollHeight")
        height = total_height/30
        new_height = 0
        for _ in range(30):
            prev_hight = new_height
            new_height += height   
            driver.execute_script(f"window.scrollTo({prev_hight}, {new_height})")
            time.sleep(0.1)
    except:
        pass

    nprods = 0
    done = False
    # checking if the link is for a single product
    try:
        wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h5[class='product-code']")))
        links.append(page)
    except:
        # scraping products urls 
        while True:
            if done: break
            try:
                prods = wait(driver, 4).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class='product-name']")))    
            except:
                print('No products are available')
                return

            for prod in prods:
                try:
                    link = wait(prod, 4).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).get_attribute('href') 
                    links.append(link)
                    nprods += 1
                    if nprods == prod_limit:
                        done = True
                        break
                except:
                    pass

            # moving to the next page
            try:
                url = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[aria-label='Next']"))).get_attribute('href')
                driver.get(url)
                time.sleep(3)
            except:
                break


    # scraping Products details
    print('-'*75)
    print('Scraping Products Details...')
    print('-'*75)

    n = len(links)
    data = pd.DataFrame()
    #comments = pd.DataFrame()
    for i, link in enumerate(links):
        try:
            # loading the chinese version of the site
            if '/en/' in link and '/zh/' not in link:
                link = link.replace('/en/', '/zh/')
            try:
                driver.get(link)   
            except:
                print(f'Warning: Failed to load the url: {link}')
                continue
  
            print(f'Scraping the details of product {i+1}\{n}')
            details = {}

            details['Store'] = 'JHC'

            # Chinese name
            name = ''
            try:
                name = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h2[class='product-name']"))).get_attribute('textContent').strip()
            except Exception as err:
                print(f'Warning: Failed to scrape the Chinese product name from: {link}')
                print(str(err))
            
            details['Product Name (Chinese)'] = name
  
            # loading the english version of the site
            driver.get(link.replace('/zh/', '/en/')) 

            # handling lazy loading
            try:
                total_height = driver.execute_script("return document.body.scrollHeight")
                height = total_height/30
                new_height = 0
                for _ in range(30):
                    prev_hight = new_height
                    new_height += height   
                    driver.execute_script(f"window.scrollTo({prev_hight}, {new_height})")
                    time.sleep(0.1)
            except:
                pass
            
            # English name
            name_en = ''
            try:
                name_en = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h2[class='product-name']"))).get_attribute('textContent').strip()
            except Exception as err:
                print(f'Warning: Failed to scrape the English product name from: {link}')
                print(str(err))       
                
            details['Product Name (English)'] = name_en
                                
            # Product ID
            prod_id = ''             
            try:
                prod_id = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h5[class='product-code']"))).get_attribute('textContent').strip()
            except:
                continue  

            details['Product ID'] = prod_id 
            details['Link'] = driver.current_url 
 
            # product image
            img = ''             
            try:
                div = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class*='product-image']")))
                try:
                    wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class*='promotag']")))
                    img = wait(div, 4).until(EC.presence_of_all_elements_located((By.TAG_NAME, "img")))[-1].get_attribute('src')
                except:
                    img = wait(div, 4).until(EC.presence_of_element_located((By.TAG_NAME, "img"))).get_attribute('src')
            except:
                continue 
                
            details['Image Link'] = img   

            # Product outline
            outline = ''             
            try:
                outline = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='desc']"))).get_attribute('textContent').strip()
                outline = ' '.join(outline.split()).split('Wish List')[0].split('心水貨品')[0].replace('Size：', '\nSize：').replace('尺寸：', '\n尺寸：').replace('沙發床：', '\n沙發床：')
            except:
                pass 
            
            details['Product Outline'] = outline

            # product price
            price = ''
            try:
                price = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[class='product-price']"))).get_attribute('textContent').split('$')[-1].strip().replace(',', '')
            except:
                pass

            details['Price (HKD)'] = price  

            # product description
            des = ''             
            try:
                div = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='tab-container']")))
                des = wait(div, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='tab-panel active']"))).get_attribute('textContent').strip()
                imgs = wait(div, 4).until(EC.presence_of_all_elements_located((By.TAG_NAME, "img")))
                for img in imgs:
                    try:
                        des += '\n' + img.get_attribute('src')
                    except:
                        pass
            except:
                pass 
             
            details['Product Description'] = des.strip('\n')   

            # other info
            info = ''             
            try:
                div = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class*='category-list']")))
                title = wait(div, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='block-title']"))).get_attribute('textContent').strip()
                info = title + ': '
                tags = wait(div, 4).until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))
                for tag in tags:
                    try:
                        text = tag.get_attribute('textContent').strip()
                        info += text + '; '
                    except:
                        pass
                info = info.strip('; ')
            except:
                pass 
             
            details['Other Information'] = info
            details['Product Type'] = cat.strip()
            details['Colour'] = ''  

            # product availability
            avail = 'In stock'             
            try:
                wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='box-qty']")))
            except:
                avail = 'Out of stock'
             
            details['Availability'] = avail   
            details['Extraction Date'] = stamp

            # appending the output to the datafame       
            data = pd.concat([data, pd.DataFrame([details.copy()])], ignore_index=True)
            
        except Exception as err:
            print(f'Warning: the below error occurred while scraping the product link: {link}')
            print(str(err))
           
    # output to excel
    if data.shape[0] > 0:
        data['Extraction Date'] = pd.to_datetime(data['Extraction Date'],  errors='coerce', format="%d_%m_%Y")
        data['Extraction Date'] = data['Extraction Date'].dt.date   
        df1 = pd.read_excel(output1)
        try:
            df1['Extraction Date'] = df1['Extraction Date'].dt.date  
        except:
            pass
        df1 = pd.concat([df1, data], ignore_index=True)
        df1 = df1.drop_duplicates()
        writer = pd.ExcelWriter(output1, date_format='d/m/yyyy')
        df1.to_excel(writer, index=False)
        writer.close()    
               
def get_inputs():
 
    print('-'*75)
    print('Processing The Settings Sheet ...')
    # assuming the inputs to be in the same script directory
    path = os.path.join(os.getcwd(), 'JHC_settings.xlsx')

    if not os.path.isfile(path):
        print('Error: Missing the settings file "JHC_settings.xlsx"')
        input('Press any key to exit')
        sys.exit(1)
    try:
        urls = []
        df = pd.read_excel(path)
        cols  = df.columns
        for col in cols:
            df[col] = df[col].astype(str)

        settings = {}
        inds = df.index
        for ind in inds:
            row = df.iloc[ind]
            link, link_type, status = '', '', ''
            for col in cols:
                if row[col] == 'nan': continue
                elif col == 'Category Link':
                    link = row[col]
                elif col == 'Scrape':
                    status = row[col]                
                elif col == 'Type':
                    link_type = row[col]                
                else:
                    settings[col] = row[col]

            if link != '' and status != '' and link_type != '':
                try:
                    status = int(float(status))
                    urls.append((link, status, link_type))
                except:
                    urls.append((link, 0, link_type))
    except:
        print('Error: Failed to process the settings sheet')
        input('Press any key to exit')
        sys.exit(1)

    # checking the settings dictionary
    keys = ["Product Limit"]
    for key in keys:
        if key not in settings.keys():
            print(f"Warning: the setting '{key}' is not present in the settings file")
            settings[key] = 0
        try:
            settings[key] = int(float(settings[key]))
        except:
            input(f"Error: Incorrect value for '{key}', values must be numeric only, press an key to exit.")
            sys.exit(1)

    return urls, settings

def initialize_output():

    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")
    path = os.path.join(os.getcwd(), 'Scraped_Data', stamp)
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path)

    file1 = f'JHC_{stamp}.xlsx'
    #file2 = f'JHC_Comments_{stamp}.xlsx'

    # Windws and Linux slashes
    output1 = os.path.join(path, file1)
    #output2 = os.path.join(path, file2)

    # Create an new Excel file and add a worksheet.
    workbook1 = xlsxwriter.Workbook(output1)
    workbook1.add_worksheet()
    workbook1.close()     
    #workbook1 = xlsxwriter.Workbook(output2)
    #workbook1.add_worksheet()
    #workbook1.close()    

    return output1#, output2

def main():

    print('Initializing The Bot ...')
    freeze_support()
    start = time.time()
    output1 = initialize_output()
    urls, settings = get_inputs()
    try:
        driver = initialize_bot()
    except Exception as err:
        print('Failed to initialize the Chrome driver due to the following error:\n')
        print(str(err))
        print('-'*75)
        input('Press any key to exit.')
        sys.exit()

    for url in urls:
        if url[1] == 0: continue
        link = url[0]
        cat = url[2]
        try:
            scrape_JHC(driver, output1, link, cat, settings)
        except Exception as err: 
            print(f'Warning: the below error occurred:\n {err}')
            driver.quit()
            time.sleep(5)
            driver = initialize_bot()

    driver.quit()
    print('-'*75)
    elapsed_time = round(((time.time() - start)/60), 4)
    hrs = round(elapsed_time/60, 4)
    input(f'Process is completed in {elapsed_time} mins ({hrs} hours), Press any key to exit.')
    sys.exit()

if __name__ == '__main__':

    main()