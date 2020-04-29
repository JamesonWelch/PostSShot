from selenium import webdriver
import pickle
import os

driver = webdriver.Chrome(os.path.join(os.getcwd(), 'chromedriver'))
driver.get('https://www.gmail.com')
pickle.dump(driver.get_cookies(), open('cookies.pkl', 'wb'))

