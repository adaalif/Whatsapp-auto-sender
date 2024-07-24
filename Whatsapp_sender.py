from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class MessageSender:
    def __init__(self):
        self.driver_path = r"C:\chromedriver-win64\chromedriver.exe"
        self.user_data_dir = r'C:\chromedriver-win64\profile'
        self.driver = None

    def setup_driver(self):
        chrome_options = Options()
        chrome_options.add_argument(f"user-data-dir={self.user_data_dir}")
        service = Service(self.driver_path)
        self.driver = webdriver.Chrome(service=service, options=chrome_options)

    def send_message(self, phone_number, message):
        if not self.driver:
            self.setup_driver()
        url = f'https://web.whatsapp.com/send?phone={phone_number}'
        self.driver.get(url)
        try:
            # Wait for the input box to be ready
            input_box = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
            )
            # Split the message by newlines and send each part with Shift+Enter except the last part
            parts = message.split('\n')
            for part in parts[:-1]:
                input_box.send_keys(part)
                input_box.send_keys(Keys.SHIFT + Keys.ENTER)
            input_box.send_keys(parts[-1])
            input_box.send_keys(Keys.ENTER)
            time.sleep(5)
            return True
        except Exception as e:
            print(f"Failed to send message: {e}")
            return False

    def close_driver(self):
        if self.driver:
            self.driver.quit()