__author__ = 'Mark Mon Monteros'

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

class ICAutomationApprover():

	def __init__(self):
		self.website = 'https://yourwebsite.com'
		self.browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

		self.launch()

	def launch(self):
		self.browser.get(self.website)
		self.browser.maximize_window()

		self.auth()

	def auth(self):
		self.login = self.browser.find_element(By.XPATH, value="//a[@href='/users/sign_in']")
		self.login.click()
		self.username = self.browser.find_element(By.ID, 'user_email')
		self.username.send_keys('myuser@email.com')
		self.password = self.browser.find_element(By.ID, 'user_password')
		self.password.send_keys('mypassword')
		self.sign_in = self.browser.find_element(By.XPATH, value="//input[@type='submit']")
		self.sign_in.click()

		self.approve()
		
	def approve(self):
		self.admin = self.browser.find_element(By.XPATH, value="//a[@href='/admin']")
		self.admin.click()
		self.server_pool = self.browser.find_element(By.XPATH, value="//a[@href='/admin/server_pool']") 
		self.server_pool.click()
		self.time_req = self.browser.find_element(By.XPATH, value="//a[@href='/admin/server_pool?tab=time_requests']")
		self.time_req.click()

		try:
			self.count = self.browser.find_element(By.XPATH, value="//span[@class='request_badge']")
			print("\nTotal requests: " + self.count)
		except:
			self.count = 'undefined'
			print("\nNo requests. Exiting...")

		if self.count != 'undefined':
			for _ in range(self.count):
				self.approve = WebDriverWait(self.browser,3).until(EC.element_to_be_clickable((By.XPATH, "//img[@alt='Accept'][@src='/assets/accept-46faf18ebe19e34487dea3f39bd917aded869b2fedba4b2b13e239406f9f23de.png']")))
				self.approve.click()
		
		self.exit()

	def exit(self):
		time.sleep(10) #comment this for timeout testing only
		self.browser.quit()

if __name__ == '__main__':
  	print('\nIC AUTOMATION APPROVER')
	print('\nCreated by: ' + __author__)

	ICAutomationApprover()

	print('\n\nDONE...!!!\n')
  
