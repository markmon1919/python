__author__ = 'Mark Mon Monteros'

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import time

class PunchAutomation():

	def __init__(self):
		self.website = 'https://yourwebsite.com/login'
		self.browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
		self.now = datetime.now()
		print(self.now)
		self.current_time = self.now.strftime("%H:%M")
		self.start_shift = '14:55' # 2:55 PM
		self.end_shift = '00:00' # 12:00 AM

		self.launch()

	def launch(self):
		self.browser.get(self.website)
		self.browser.maximize_window()

		self.auth()

	def auth(self):
		self.username = self.browser.find_element(By.ID, "userName")
		self.username.send_keys('my_username')
		self.password = self.browser.find_element(By.ID, "password")
		self.password.send_keys('my_password')
		self.login = self.browser.find_element(By.XPATH, value="//button[@data-testid='button']")
		self.login.click()

		try:
			self.check = WebDriverWait(self.browser,10).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='top-toggle']")))
			print("\nLogin Success!")

			if self.current_time == self.start_shift:
				self.punch_in()
			elif self.current_time == self.end_shift:
				self.punch_out()
			else:
				print("\nCan't punch incorrect shift schedule.")
		except:
			print("\nLogin Failed! Invalid username or password")

		self.exit()

	def punch_in(self):
		try:
			self.clock_in = WebDriverWait(self.browser,3).until(EC.element_to_be_clickable((By.XPATH, "//button[@class='button button-clockin']")))
			self.clock_in.click()
			print("\n[*] - CLOCKING IN @ " + self.current_time)
		except:
			print("\n\nERROR during Clock-In!")
			print("\nCurrent state is Clock-In...")

	def punch_out(self):
		try:
			self.clock_out = WebDriverWait(self.browser,3).until(EC.element_to_be_clickable((By.XPATH, "//button[@class='button button-clockout']")))
			self.clock_out.click()
			print("\n[*] - CLOCKING OUT @ " + self.current_time)
		except:
			print("\n\nERROR during Clock-Out!")
			print("\nCurrent state is Clock-Out...")

	def exit(self):
		# time.sleep(5) #comment this for timeout testing only
		self.browser.quit()

if __name__ == '__main__':
	print('\nPUNCH AUTOMATION')
	print('\nCreated by: ' + __author__)

	PunchAutomation()

	print('\n\nDONE...!!!\n')


	# ADD ENTRIES TO CRONJOB
	# 55 14 * * 1-5  /usr/local/bin/python3 ~/Projects/Emapta/punch_automation.py # run “At 02:55 PM on every day-of-week from Monday through Friday.”
	# 00 00 * * 2-6 /usr/local/bin/python3 ~/Projects/Emapta/punch_automation.py # run “At 12:00 AM on every day-of-week from Tuesday through Saturday.”
