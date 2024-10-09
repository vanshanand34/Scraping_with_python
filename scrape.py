import pandas as pd
import time

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as rt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException


def extract_urls(file):
    '''Extracting links from a given docx file'''
    links = []
    document = Document(file)
    rels = document.part.rels
    for rel in rels:
        if rels[rel].reltype == rt.HYPERLINK:
            links.append(rels[rel]._target)
    return links  # extracting the urls from a docx file


def fill_list(mylist, size):
    '''Function to fill lists with '-' character for making several lists to equal sizes so to convert them conveniently in the form of an excel file ( as when we try to convert python dictionary to excel format all the values of keys should have equal number of elements)'''
    for i in range(len(mylist) - 1, size):
        mylist.append("-")
    return mylist


def scrape(links):
    '''Function to take list of links as an argument and scrape the website of provided links and return the scraped data ( URLs , Text , Images ) in form of a python dictionary'''
    res_text = []
    res_links = []
    res_images = []
    res_websites = []
    for link in links:
        try:
            driver = webdriver.Chrome()
            urls = []
            images = []
            driver.get(link)
            time.sleep(5)
            text = (driver.find_element(By.TAG_NAME, 'body').text)
            urls = [a.get_attribute("href") for a in driver.find_elements(By.TAG_NAME, "a")]
            images = [image.get_attribute("src") for image in driver.find_elements(By.TAG_NAME, "img")]
        except StaleElementReferenceException:
            print("Got stale element reference exception")
            driver.quit()
            continue
        website = [link]
        maxsize = max(len(text), max(len(urls), len(images)))
        # maximum length among all the lists ( text , images , urls)
        # to fill list having text , images and urls to make all lists of equal size
        # in order to successfully convert data in the form of an excel file
        text = text.split('\n')  # to make text span several columns when converted in excel file form
        text = fill_list(text, maxsize)
        images = fill_list(images, maxsize)
        urls = fill_list(urls, maxsize)
        website = fill_list(website, maxsize)
        # making each list of equal size by filling with '-' char to each list
        res_websites.extend(website)
        res_text.extend(text)
        res_images.extend(images)
        res_links.extend(urls)
        driver.quit()
    return {'website_url': res_websites, 'text': res_text, 'links': res_links, 'images': res_images}


mylist = extract_urls('assign.docx')
data = scrape(mylist)
df = pd.DataFrame(data)
df.to_excel('output.xlsx')  # file in which the scraped data to be stored
