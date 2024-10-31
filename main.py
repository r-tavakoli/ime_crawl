from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
from datetime import datetime
from persiantools.jdatetime import JalaliDate 
import yaml
from loguru import logger


def read_config_data(file_name):
    # reading file configurations
    with open(file_name, 'r') as fr:
        config_info = yaml.load(fr, Loader=yaml.SafeLoader)

    return config_info

def scrape(config_info: dict):
    # selenium configurations
    service = Service()
    options = webdriver.ChromeOptions()
    options.add_experimental_option('prefs', {'download.default_directory' : config_info['output_path']})
    options.add_argument('--start-maximized')
    if not config_info['show_browser']:
        options.add_argument("--headless=new")
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.80 Safari/537.36 Edg/98.0.1108.62'
    options.add_argument(f'user-agent={user_agent}')
    driver = webdriver.Chrome(service=service, options=options)   

    # sending request to website
    driver.get('https://www.ime.co.ir/offer-stat.html')

    WebDriverWait(driver, 30).until(
        EC.presence_of_all_elements_located((By.ID, 'FillGrid'))
    )
    sleep(5)

    # page container
    page_container = driver.find_element(By.ID, 'PageContainer')

    # selecting all columns
    columns_drop_down = page_container.find_element(By.CLASS_NAME, 'keep-open')
    columns_drop_down.find_element(By.TAG_NAME, 'button').click()
    columns_selection = columns_drop_down.find_elements(By.TAG_NAME, 'label')
    action_chain = ActionChains(driver)
    for c in columns_selection:
            cb = c.find_element(By.TAG_NAME, 'input')
            if not cb.is_selected():
                action_chain.move_to_element(cb).perform()
                sleep(2)
                cb.click()    

    # fill the date input boxes
    # by default from yesterday to today
    date_range_container = page_container.find_element(By.ID, 'DatesHolder')
    from_date_input_box = date_range_container.find_element(By.ID, 'ctl05_ReportsHeaderControl_FromDate')
    to_date_input_box = date_range_container.find_element(By.ID, 'ctl05_ReportsHeaderControl_ToDate')


    last_read_year, last_read_month, last_read_day = config_info['last_read_date'].split('/')

    from_date_input_box.clear()
    from_date = JalaliDate(int(last_read_year), int(last_read_month), int(last_read_day))
    from_date_input_box.send_keys(from_date.strftime('%Y/%m/%d'))

    to_date_input_box.clear()
    to_date = JalaliDate.today()
    to_date_input_box.send_keys(to_date.strftime('%Y/%m/%d'))


    # get data with clicking "show button"
    show_button = page_container.find_element(By.ID, 'FillGrid')
    show_button.click()
    time_to_wait = (to_date - from_date).days * 5 if to_date != from_date else 5
    sleep(time_to_wait)


    # choosing the export method
    export_list = page_container.find_element(By.CLASS_NAME, 'export')
    export_list.find_element(By.TAG_NAME, 'button').click()
    export_modes = export_list.find_elements(By.TAG_NAME, 'a')

    for e in export_modes:
        if e.text == config_info['export_method']:
            e.click()


    sleep(time_to_wait)

    # close the browser
    driver.quit()


    # update config file
    config_info['last_read_date'] = to_date.strftime('%Y/%m/%d')

    with open('config.yaml', 'w') as fw:
        yaml.dump(config_info, fw)


if __name__=='__main__':
    logger.add('logs\errors.log', format="{time} {name} {level} {message}", level='ERROR', backtrace=True, diagnose=True)
    try:
        config_dict = read_config_data(file_name='config.yaml')
        scrape(config_info=config_dict)
    except Exception as e:
        logger.exception(e)