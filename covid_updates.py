from selenium import webdriver
from time import sleep
import re
from datetime import datetime
import smtplib


options = webdriver.ChromeOptions()

driver_path = 'C:/Users/ejmooney/plugins/chromedriver'
driver = webdriver.Chrome(executable_path=driver_path)

class Coronavirus():
    def __init__(self):
        self.driver = webdriver.Chrome(executable_path=driver_path)

    def get_data(self):
        self.driver.get('https://www.worldometers.info/coronavirus/')
        table = self.driver.find_element_by_xpath("//*[@id='main_table_countries_today']/tbody[1]")
        country_element = table.find_element_by_xpath("//td[contains(., 'USA')]")
        row = country_element.find_element_by_xpath('./..')
        data = row.text.split(' ')
        total_cases = data[1]
        new_cases = data[2]
        total_deaths = data[3]
        new_deaths = data[4]
        active_cases = data[6]
        total_recovered = data[5]
        serious_critical = data[7]

        send_mail(country_element.text, total_cases, new_cases, total_deaths, new_deaths, active_cases, total_recovered, serious_critical)

        self.driver.close()
        self.driver.quit()


def send_mail(country_element, total_cases, new_cases, total_deaths, new_deaths, active_cases,
                total_recovered, serious_critical):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login('mooney.ethan@gmail.com', 'Arsenal#12020')
    subject = 'Coronavirus stats today'
    body = 'Today in ' + country_element + '\
    \nThere is new data on coronavirus:\
    \nTotal cases: ' + total_cases + '\
    \nNew Cases: ' + new_cases + '\
    \nTotal Deaths: ' + total_deaths + '\
    \nNew Deaths: ' + new_deaths + '\
    \nActive cases: ' + active_cases + '\
    \nTotal recovered: ' + total_recovered + '\
    \nSerious, critical cases: ' + serious_critical + '\
    \n Check the link:  https://www.worldometers.info/coronavirus/'

    msg = f'Subject: {subject}\n\n{body}'
    server.sendmail('Coronavirus', 'mooney.ethan@gmail.com', msg)

    print ("email sent!")
    server.quit()

bot = Coronavirus()
bot.get_data()