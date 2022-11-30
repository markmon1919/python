__author__ = 'Mark Mon Monteros'

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from datetime import datetime
import time

class PunchAutomation():

	def __init__(self):
		self.website = 'https://yourwebsite.com/login'
		self.browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
		self.now = datetime.now()
		self.current_time = self.now.strftime("%H:%M:%S")
		self.start_shift = '14:55:00' # 2:55 PM
		self.end_shift = '00:00:00' # 12:00 AM

		self.launch()

	def launch(self):
		self.browser.get(self.website)
		self.browser.maximize_window()

		self.auth()

	def auth(self):
		self.username = self.browser.find_element(By.ID, "userName")
		self.username.send_keys("myusernamehere")
		self.password = self.browser.find_element(By.ID, "password")
		self.password.send_keys("mypasswordhere")
		self.login = self.browser.find_element(By.XPATH, value="//button[@data-testid='button']")
		self.login.click()

		if self.current_time == self.start_shift:
			self.punch_in()

		if self.current_time == self.end_shift:
			self.punch_out()
		
	def punch_in(self):
		self.clock_in = self.browser.find_element(By.XPATH, value="//button[@id='GA-clockin-button-topbar']")
		self.clock_in.click()
		print("\n[*] - CLOCKING IN @ " + self.current_time)

		self.exit()

	def punch_out(self):
		self.clock_out = self.browser.find_element(By.XPATH, value="//button[@id='GA-clockout-button-topbar']")
		self.clock_out.click()
		print("\n[*] - CLOCKING OUT @ " + self.current_time)
		self.exit()

	def exit(self):
		time.sleep(10) #comment this for timeout testing only
		self.browser.quit()

if __name__ == '__main__':
	print('\nPUNNCH AUTOMATION')
	print('\nCreated by: ' + __author__)

	PunchAutomation()

	print('\n\nDONE...!!!\n')


	# ADD ENTRIES TO CRONJOB
	# 55 2 * * 1-5  /usr/local/bin/python3 ~/Projects/Emapta/punch_automation.py # run “At 02:55 PM on every day-of-week from Monday through Friday.”
	# 00 00 * * 1-5 /usr/local/bin/python3 ~/Projects/Emapta/punch_automation.py # run “At 12:00 AM on every day-of-week from Monday through Friday.”

	# ELEMENTS IN WEBSITE
	# <input autocomplete="off" data-testid="userName" id="userName" type="text" placeholder="Enter your account username" required="" maxlength="50" value="223915.mmonteros">
	# <input autocomplete="off" data-testid="password" id="password" type="password" placeholder="Enter your password" required="" maxlength="257" value="m1B`FN^Q">
	# <button data-testid="button">Login</button>
	# <button id="GA-clockin-button-topbar" class="button button-clockin"><div class="icon"></div><span>Clock In</span></button>
	# <button id="GA-clockout-button-topbar" class="button button-clockout"><div class="icon"></div><span>Clock Out</span></button>

