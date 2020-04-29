
import os
from PIL import Image
import time
import xlrd
import webbrowser
import pyscreenshot as ImageGrab
import datetime
import pyautogui
import tldextract


prog_start = time.time()
now = datetime.datetime.now()

# MEIPATH for Mac:
# ROOT_DIR = os.path.join(sys._MEIPASS, "../../..")
# Path for Windows:
ROOT_DIR = os.path.join(os.getcwd(), "../../..")
# Sandbox ROOT_DIR-
# ROOT_DIR = os.path.join(os.getcwd(), '.')
SYSTEM_DIR = os.path.join(ROOT_DIR, "System Files")
SSHOTS_DIR = os.path.join(ROOT_DIR, "Screenshots")

def main():
    print('Importing Links Spreadsheet')
    wb = xlrd.open_workbook(os.path.join(ROOT_DIR, 'links.xlsx'))
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0,0)
    posts = []
    for i in range(sheet.nrows):
        if sheet.cell_value(i, 0) != 'Links':
            posts.append(sheet.cell_value(i, 0))
    print(posts)
    for post in posts:
        if 'twitter.com' in post:
            pform = "Twitter"
            twitterCrop(screenshot(post, pform))
        elif 'facebook.com' in post:
            pform = 'Facebook7'
            facebookCrop(screenshot(post, pform, delay=7))
        elif 'instagram.com' in post:
            pform = "Instagram"
            instagramCrop(screenshot(post, pform))
        elif 'linkedin.com' in post:
            pform = "LinkedIn"
            linkedInCrop(screenshot(post, pform))
        elif 'youtube.com' in post:
            pform = "Youtube"
            youtubeCrop(screenshot(post, pform, delay=5, zoom=0))
        else:
            pform = tldextract.extract(post)
            pform = pform.domain
            screenshot(post, pform, zoom=0)

def measure_crop():
    pass

def screenshot(post, pform, delay=4, zoom=3):
    def zoom_out():
        for _ in range(zoom):
            pyautogui.keyDown('ctrl')
            pyautogui.keyDown('-')
            pyautogui.keyUp('-')
            pyautogui.keyUp('ctrl')
    def zoom_in():
        for _ in range(zoom):
            pyautogui.keyDown('ctrl')
            pyautogui.keyDown('+')
            pyautogui.keyUp('+')
            pyautogui.keyUp('ctrl')
    def close_tab():
        time.sleep(2)
        pyautogui.keyDown('ctrl')
        pyautogui.keyDown('w')
        pyautogui.keyUp('w')
        pyautogui.keyUp('ctrl')
        time.sleep(2)
    print('Fetching URL source')
    webbrowser.open_new(post)
    
    print(post)
    print(f'Sleeping {delay} seconds')
    time.sleep(delay)
    print('Taking screenshot...')
    zoom_out()
    myScreenshot = pyautogui.screenshot()
    final = os.path.join(SSHOTS_DIR, f'{pform}_{now.strftime("%Y-%m-%d_%H-%M")}.png')
    myScreenshot.save(final)
    print(f'______ {final} _____')
    zoom_in()
    close_tab()
    return final

def instagramCrop(image):
    print(f'Cropping {image}')
    im = Image.open(image)
    x1 = 483 # over from upper left
    y1 = 198  # down for upper left
    x2 = 1411 # over from right / measure from left
    y2 = 888 # over from bottom / measure from top
    cropped = im.crop((x1, y1, x2, y2))
    print(f'Saving {image}')
    cropped.save(f"{image}")

def twitterCrop(image):
    print(f'Cropping {image}')
    im = Image.open(image)
    x1 = 543 # over from upper left
    y1 = 164  # down for upper left
    x2 = 1223 # over from right / measure from left
    y2 = 879 # over from bottom / measure from top
    cropped = im.crop((x1, y1, x2, y2))
    print(f'Saving {image}')
    cropped.save(f"{image}")

def youtubeCrop(image):
    print(f'Cropping {image}')
    im = Image.open(image)
    x1 = 97 # over from upper left
    y1 = 194  # down for upper left
    x2 = 1161 # over from right / measure from left
    y2 = 965 # over from bottom / measure from top
    cropped = im.crop((x1, y1, x2, y2))
    print(f'Saving {image}')
    cropped.save(f"{image}")

def facebookCrop(image):
    print(f'Cropping {image}')
    im = Image.open(image)
    x1 = 519 # over from upper left
    y1 = 226  # down for upper left
    x2 = 1214 # over from right / measure from left
    y2 = 919 # over from bottom / measure from top
    cropped = im.crop((x1, y1, x2, y2))
    print(f'Saving {image}')
    cropped.save(f"{image}")

def linkedInCrop(image):
    print(f'Cropping {image}')
    im = Image.open(image)
    x1 = 578 # over from upper left
    y1 = 217  # down for upper left
    x2 = 1212 # over from right / measure from left
    y2 = 1017 # over from bottom / measure from top
    cropped = im.crop((x1, y1, x2, y2))
    print(f'Saving {image}')
    cropped.save(f"{image}")


main()