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


