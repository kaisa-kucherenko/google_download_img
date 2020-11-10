import os
import shutil
import requests
from time import sleep
from selenium import webdriver
from openpyxl import load_workbook

directory = os.getcwd()
fc_names_file = './Book1.xlsm'


def get_fc_names_lists(xlsm_file_name):
    work_book = load_workbook(xlsm_file_name)
    sheet = work_book['Sheet1']
    fc_names_list_clean = [sheet[f'A{num}'].value.strip() for num in range(1, 1000)
                           if sheet[f'A{num}'].value]
    fc_names_list_for_google_search = [f'FC+{fc_name.replace(" ", "+")}'
                                       for fc_name in fc_names_list_clean]
    work_book.close()
    return fc_names_list_clean, fc_names_list_for_google_search

def find_imgs_urls(fc_name, img_number=3):
    browser = webdriver.Firefox()
    url = f"http://www.google.com/search?q={fc_name}&tbm=isch&tbs=ift:png"
    browser.get(url)
    sleep(2)
    src_list = list()
    for num in range(1, img_number+1):
        img_url = browser.find_element_by_xpath(f'//div//div//div//div//div//div//div//div//div//div[{num}]//a[1]//div[1]//img[1]')
        img_url.click()
        sleep(5)
        src = browser.find_element_by_xpath('//body/div[2]/c-wiz/div[3]/div[2]/div[3]/div/div/div[3]/div[2]/c-wiz/div[1]/div[1]/div/div[2]/a/img').get_attribute("src")
        src_list.append(src)
    browser.quit()
    return src_list
