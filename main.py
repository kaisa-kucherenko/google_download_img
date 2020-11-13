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


def find_imgs_urls(browser, search_query, img_number, not_fc_club):
    if not_fc_club:
        search_query = f'FC+{search_query.replace(" ", "+")}'
    else:
        search_query = search_query.replace(" ", "+")
    log.warning(f'searchquery {search_query}')
    url = f"""http://www.google.com/search?q={search_query}&tbm=isch&tbs=ift:png"""
    browser.get(url)
    log.info(f'Get {search_query} in browser')
    sleep(2)
    src_list = list()
    for num in range(1, img_number+1):
        img_url = browser.find_element_by_xpath(f'//div//div//div//div//div//div//div//div//div//div[{num}]//a[1]//div[1]//img[1]')
        img_url.click()
        log.info(f'Looking src of {search_query} in web page')
        sleep(5)
        src = browser.find_element_by_xpath('//body/div[2]/c-wiz/div[3]/div[2]/div[3]/div/div/div[3]/div[2]/c-wiz/div[1]/div[1]/div/div[2]/a/img').get_attribute("src")
        if not src.startswith('data:'):
            log.info(f'I found src of {search_query}')
            src_list.append(src)
    return src_list


def download_imgs(search_query, src_list):
    directory = os.getcwd()
    file_name = f'{search_query}.png'
    for index, src in enumerate(src_list):
        if index == 0:
            image_path = os.path.join(directory, file_name)
        else:
            image_path = os.path.join(directory, f'{search_query}_{index}.png')
        try:
            log.info(f"I'm trying to download a img {search_query}")
            response = requests.get(src, stream=True)
            with open(image_path, 'wb') as file:
                shutil.copyfileobj(response.raw, file)
            log.info("Downloaded, take the next one")
        except Exception:
            log.error(f'src {src} cant download. Skipped it')
            pass


def main():
    parser = argparse.ArgumentParser(description='Download png imgs from google')
    parser.add_argument('xlsm_file', type=str,
                        help='File with image search query')
    parser.add_argument('-sheet_name', type=str, default='Sheet1',
                        help='Sheet name in file (default Sheet1)')
    parser.add_argument('-img_num', type=int, default=4,
                        help='Number downloading imgs per query (default 4)')
    parser.add_argument('-not_fc_club', action='store_false', default=True,
                        help='Parse football club names (default True)')
    args = parser.parse_args()

    firefox_options = webdriver.FirefoxOptions()
    firefox_options.headless = True
    browser = webdriver.Firefox(options=firefox_options)

    search_query_clean = get_search_query_lists(args.xlsm_file, args.sheet_name)
    try:
        for query in search_query_clean:
            src_list = find_imgs_urls(browser, query,
                                      img_number=args.img_num,
                                      not_fc_club=args.not_fc_club)
            download_imgs(query, src_list)
    except Exception:
        log.error(f'Something go wrong. I cant download imgs')
    finally:
        browser.quit()


if __name__ == '__main__':
    main()

