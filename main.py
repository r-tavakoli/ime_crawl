from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
import datetime
from persiantools.jdatetime import JalaliDate 
import yaml
from loguru import logger

class BaseScrape:

    def __init__(self, config_file_path):
        self.config_file_path = config_file_path
        self.read_configuration()
        service = Service()
        options = webdriver.ChromeOptions()
        options.add_experimental_option('prefs', {'download.default_directory' : self.config_info['output_path']})
        options.add_argument('--start-maximized')
        options.add_experimental_option("detach", True)
        if not self.config_info['show_browser']:
            options.add_argument("--headless=new")
        user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.80 Safari/537.36 Edg/98.0.1108.62'
        options.add_argument(f'user-agent={user_agent}')
        self.driver = webdriver.Chrome(service=service, options=options)  
        self.driver.implicitly_wait(10)

    def read_configuration(self):
        with open(self.config_file_path, 'r') as fr:
            self.config_info = yaml.load(fr, Loader=yaml.SafeLoader)
    
    def write_configuration(self):
        with open(self.config_file_path, 'w') as fw:
            yaml.dump(self.config_info, fw)
    

    def quit(self):
        self.driver.quit()


class Scrape(BaseScrape):
    def __init__(self, config_file_path):
        super().__init__(config_file_path=config_file_path)
        self.last_read_year, self.last_read_month, self.last_read_day = self.config_info['last_read_date'].split('/')
        self.is_incremental = 0 if self.config_info['first_time_read']['state'] == 'active' else 1
        self.full_scrape_from_year, self.full_scrape_from_month, self.full_scrape_from_day = self.config_info['first_time_read']['from_date'].split('/')
        self.full_scrape_to_year, self.full_scrape_to_month, self.full_scrape_to_day = self.config_info['first_time_read']['to_date'].split('/')
        self.full_scrape_steps = self.config_info['first_time_read']['step_days']
        self.selected_export_mode = self.config_info['export_method']
        
    def start(self):
        self.get_preparation_elements()
        if self.is_incremental==1:
            self.scrape_incremental()
        else:
            self.scrape_full()

        self.quit()

    def get_preparation_elements(self):
        # request
        self.driver.get('https://www.ime.co.ir/offer-stat.html')
        WebDriverWait(self.driver, 30).until(
            EC.presence_of_all_elements_located((By.ID, 'FillGrid'))
        )
        sleep(5)

        # page container
        self.page_container = self.driver.find_element(By.ID, 'PageContainer')

        # selecting all columns
        columns_drop_down = self.page_container.find_element(By.CLASS_NAME, 'keep-open')
        columns_drop_down.find_element(By.TAG_NAME, 'button').click()
        columns_selection = columns_drop_down.find_elements(By.TAG_NAME, 'label')
        action_chain = ActionChains(self.driver)
        for c in columns_selection:
                cb = c.find_element(By.TAG_NAME, 'input')
                if not cb.is_selected():
                    action_chain.move_to_element(cb).perform()
                    sleep(2)
                    cb.click()   

        # date range elements
        # by default from yesterday to today
        date_range_container = self.page_container.find_element(By.ID, 'DatesHolder')
        self.from_date_input_box = date_range_container.find_element(By.ID, 'ctl05_ReportsHeaderControl_FromDate')
        self.to_date_input_box = date_range_container.find_element(By.ID, 'ctl05_ReportsHeaderControl_ToDate')    


        # show button
        self.show_button = self.page_container.find_element(By.ID, 'FillGrid')

        # output mode
        self.export_list = self.page_container.find_element(By.CLASS_NAME, 'export')

    def scrape_full(self):
        '''
            it gets data from_date until to_date if state is "active"
            to minimum the wait and error we will get data in steps by default it is 10 days
        '''
        from_date = JalaliDate(int(self.full_scrape_from_year), int(self.full_scrape_from_month), int(self.full_scrape_from_day))
        end_date = JalaliDate(int(self.full_scrape_to_year), int(self.full_scrape_to_month), int(self.full_scrape_to_day))

        while from_date<=end_date:
            # date range
            if (end_date - from_date).days <= self.full_scrape_steps:
                to_date = end_date
            else:
                to_date = from_date + datetime.timedelta(days=self.full_scrape_steps)

            sleep(3)
            self.from_date_input_box.clear()
            self.to_date_input_box.clear()
            self.from_date_input_box.send_keys(from_date.strftime('%Y/%m/%d'))
            self.to_date_input_box.send_keys(to_date.strftime('%Y/%m/%d'))


            # get data with clicking "show button"
            self.show_button.click()
            time_to_wait = (to_date - from_date).days * 10 if to_date != from_date else 10
            sleep(time_to_wait)
            WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable((By.ID, 'FillGrid'))
            )


            # choosing the export method
            export_list = self.page_container.find_element(By.CLASS_NAME, 'export')
            export_list.find_element(By.TAG_NAME, 'button').click()
            export_modes = export_list.find_elements(By.TAG_NAME, 'a')
            for e in export_modes:
                if e.text == self.selected_export_mode:
                    e.click()

            sleep(time_to_wait)

            # update config file
            self.config_info['last_read_date'] = to_date.strftime('%Y/%m/%d')
            self.config_info['first_time_read']['from_date'] = to_date.strftime('%Y/%m/%d')

            # write data
            self.write_configuration()

            # next date range
            from_date = from_date + datetime.timedelta(days=self.full_scrape_steps)
        
    def scrape_incremental(self):
        '''
            it gets data from last_update_date until today (persian date).
            last_update_date of config file will be updated after execution
        '''
        # date range
        from_date = JalaliDate(int(self.last_read_year), int(self.last_read_month), int(self.last_read_day))
        to_date = JalaliDate.today()


        self.from_date_input_box.clear()
        self.to_date_input_box.clear()
        self.from_date_input_box.send_keys(from_date.strftime('%Y/%m/%d'))
        self.to_date_input_box.send_keys(to_date.strftime('%Y/%m/%d'))


        # get data with clicking "show button"
        self.show_button.click()
        time_to_wait = (to_date - from_date).days * 10 if to_date != from_date else 10
        WebDriverWait(self.driver, 30).until(
            EC.element_to_be_clickable((By.ID, 'FillGrid'))
        )       
        sleep(time_to_wait)


        # choosing the export method
        export_list = self.page_container.find_element(By.CLASS_NAME, 'export')
        export_list.find_element(By.TAG_NAME, 'button').click()
        export_modes = export_list.find_elements(By.TAG_NAME, 'a')
        for e in export_modes:
            if e.text == self.selected_export_mode:
                e.click()

        sleep(time_to_wait)


        # update config file
        self.config_info['last_read_date'] = to_date.strftime('%Y/%m/%d')

        # write data
        self.write_configuration()


if __name__=='__main__':
    logger.add('logs\log_file.log', format="{time} {name} {level} {message}", level='ERROR', backtrace=True, diagnose=True)
    try:
        scrape = Scrape(config_file_path='config.yaml')
        scrape.start()
    except Exception as e:
        logger.exception(e)
        scrape.quit()