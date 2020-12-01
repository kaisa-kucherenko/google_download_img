import os
import shutil
import requests
import argparse
import logging
from time import sleep
from selenium import webdriver
from openpyxl import load_workbook


def logger(logger_name):
    log = logging.getLogger(logger_name)
    log.setLevel('INFO')
    ch = logging.StreamHandler()
    ch.setLevel('INFO')
    formatter = logging.Formatter('[%(asctime)s %(levelname)s]: %(message)s')
    ch.setFormatter(formatter)
    log.addHandler(ch)
    return log


log = logger('img_download')


def get_search_query_lists(xlsm_file_name, sheet_name):
    work_book = load_workbook(xlsm_file_name)
    sheet = work_book[sheet_name]
    search_query_list_clean = [sheet[f'A{num}'].value.strip() for num
                               in range(1, 1000)
                               if sheet[f'A{num}'].value]
    work_book.close()
    return search_query_list_clean


def find_imgs_urls(browser, search_query, img_number, fc_club):
    if fc_club:
        if '/' in search_query:
            search_query = f'{search_query.replace("/", "-").replace(" ", "+")}+fc'
        elif '&' in search_query:
            search_query = f'{search_query.replace("&", "and").replace(" ", "+")}+fc'
        else:
            search_query = f'{search_query.replace(" ", "+")}+fc'
    else:
        search_query = search_query.replace(" ", "+")
    log.info(f'Search query {search_query}')
    url = f"http://www.google.com/search?q={search_query}&tbm=isch&tbs=ift:png"
    browser.get(url)
    log.info(f'Get {search_query} in browser')
    sleep(2)
    src_list = list()
    counter = 1
    while len(src_list) < img_number:
        if counter in (2, 39, 79, 119):
            browser.execute_script("window.scrollTo(0, "
                                   "document.body.scrollHeight || "
                                   "document.documentElement.scrollHeight);")
        img_url = browser.find_element_by_xpath(f'//div//div//div//div//div//'
                                                f'div//div//div//div//div[{counter}]'
                                                f'//a[1]//div[1]//img[1]')
        img_url.click()
        log.info(f'Looking src of {search_query} in web page')
        sleep(5)
        src = browser.find_element_by_xpath('//body/div[2]/c-wiz/div[3]/div[2]/'
                                            'div[3]/div/div/div[3]/div[2]/c-wiz'
                                            '/div[1]/div[1]/div/div[2]'
                                            '/a/img').get_attribute("src")
        if not src.startswith('data:'):
            log.info(f'I found src of {search_query}')
            src_list.append(src)
        counter += 1
    return src_list


def download_imgs(search_query, src_list):
    directory = os.getcwd()
    if '/' in search_query:
        search_query = search_query.replace("/", "-")
    file_name = f'{search_query}.png'
    for index, src in enumerate(src_list):
        if index == 0:
            image_path = os.path.join(directory, file_name)
        else:
            image_path = os.path.join(directory,
                                      f'{search_query}{"_"*index}.png')
        try:
            log.info(f"I'm trying to download a img {search_query}")
            response = requests.get(src, stream=True)
            with open(image_path, 'wb') as file:
                shutil.copyfileobj(response.raw, file)
            log.info("Downloaded, take the next one")
        except Exception:
            log.error(f'src {src} cant download. Skipped it')
            pass


def main(browser, xlsm_file, sheet_name, img_num, fc_club):
    search_query_clean = get_search_query_lists(xlsm_file, sheet_name)
    try:
        for query in search_query_clean:
            src_list = find_imgs_urls(browser, query,
                                      img_number=img_num,
                                      fc_club=fc_club)
            download_imgs(query, src_list)
    except Exception as e:
        log.error(f'Something go wrong. I cant download imgs {e}')
    finally:
        browser.quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Download png imgs from google')
    parser.add_argument('xlsm_file', type=str,
                        help='File with image search query')
    parser.add_argument('-sheet_name', type=str, default='Sheet1',
                        help='Sheet name in file (default Sheet1)')
    parser.add_argument('-img_num', type=int, default=4,
                        help='Number downloading imgs per query (default 4)')
    parser.add_argument('-fc_club', action='store_true', default=False,
                        help='Parse football club names (default False)')
    args = parser.parse_args()

    operating_system = os.name
    firefox_options = webdriver.FirefoxOptions()
    firefox_options.headless = True
    if operating_system == 'nt':
        firefox_options.binary = r'C:\Program Files\Mozilla Firefox\firefox.exe'
    browser = webdriver.Firefox(options=firefox_options)

    main(browser, args.xlsm_file, args.sheet_name,
         args.img_num, args.fc_club)

