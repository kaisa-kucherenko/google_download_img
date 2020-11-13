import os
import shutil
import requests
import argparse
from time import sleep
from selenium import webdriver
from openpyxl import load_workbook


def get_fc_names_lists(xlsm_file_name, sheet_name):
    work_book = load_workbook(xlsm_file_name)
    sheet = work_book[sheet_name]
    fc_names_list_clean = [sheet[f'A{num}'].value.strip() for num in range(1, 1000)
                           if sheet[f'A{num}'].value]
    work_book.close()
    return fc_names_list_clean


def find_imgs_urls(browser, fc_name, img_number, fc_club):
    if fc_club:
        fc_name = f'FC+{fc_name.replace(" ", "+")}'
    else:
        fc_name = fc_name.replace(" ", "+")
    url = f"""http://www.google.com/search?q={fc_name}&tbm=isch&tbs=ift:png"""
    browser.get(url)
    sleep(2)
    src_list = list()
    for num in range(1, img_number+1):
        img_url = browser.find_element_by_xpath(f'//div//div//div//div//div//div//div//div//div//div[{num}]//a[1]//div[1]//img[1]')
        img_url.click()
        sleep(5)
        src = browser.find_element_by_xpath('//body/div[2]/c-wiz/div[3]/div[2]/div[3]/div/div/div[3]/div[2]/c-wiz/div[1]/div[1]/div/div[2]/a/img').get_attribute("src")
        if not src.startswith('data:'):
            src_list.append(src)
    return src_list


def download_imgs(fc_name, src_list):
    directory = os.getcwd()
    file_name = f'{fc_name}.png'
    for index, src in enumerate(src_list):
        if index == 0:
            image_path = os.path.join(directory, file_name)
        else:
            image_path = os.path.join(directory, f'{fc_name}_{index}.png')
        try:
            response = requests.get(src, stream=True)
            with open(image_path, 'wb') as file:
                shutil.copyfileobj(response.raw, file)
        except Exception:
            print(f'src {src} cant download')
            pass


def main():
    parser = argparse.ArgumentParser(description='Download png imgs from google')
    parser.add_argument('xlsm_file', type=str,
                        help='File with image search query')
    parser.add_argument('-sheet_name', type=str, default='Sheet1',
                        help='Sheet name in file (default Sheet1)')
    parser.add_argument('-img_num', type=int, default=4,
                        help='Number downloading imgs per query (default 4)')
    parser.add_argument('-fc_club', type=bool, default=True,
                        help='Parse football club names (default True)')
    args = parser.parse_args()

    firefox_options = webdriver.FirefoxOptions()
    firefox_options.headless = True
    browser = webdriver.Firefox(options=firefox_options)

    fc_names_clean = get_fc_names_lists(args.xlsm_file, args.sheet_name)
    try:
        for fc_name in fc_names_clean:
            src_list = find_imgs_urls(browser, fc_name,
                                      img_number=args.img_num,
                                      fc_club=args.fc_club)
            download_imgs(fc_name, src_list)
    except Exception:
        print('Something go wrong')
    finally:
        browser.quit()


if __name__ == '__main__':
    main()

