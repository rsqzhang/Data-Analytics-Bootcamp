from bs4 import BeautifulSoup
from splinter import Browser
import pandas as pd
import requests
import sys

!{sys.executable} -m pip install pymongo
import pymongo

def init_browser():
    executable_path = {'executable_path': 'chromedriver.exe'}
    browser = Browser('chrome', **executable_path, headless=False)

def scrape():
    browser = init_browser()
    scraped_data = {}

    # NASA Mars News
    news_url = 'https://mars.nasa.gov/news/'
    response = requests.get(news_url)
    soup = BeautifulSoup(response.text, 'html.parser')
    results = soup.find_all('div', class_='slide')[0]

    scraped_data['news_title'] = results.find('div', class_='content_title').a.text
    scraped_data['news_p'] = results.find('div', class_='rollover_description_inner').text

    # JPL Mars Space Images - Featured Image
    images_url = 'https://www.jpl.nasa.gov/spaceimages/?search=&category=Mars'
    browser.visit(images_url)
    browser.click_link_by_partial_text('FULL IMAGE')
    featured_image = browser.find_by_css('a.fancybox-expand')
    soup = BeautifulSoup(browser.html, 'html.parser')
    image_url = soup.find('img', class_='fancybox-image')['src']
    
    scraped_data['featured_image_url'] = f'https://www.jpl.nasa.gov{image_url}'

    # Mars Weather
    weather_url = 'https://twitter.com/marswxreport?lang=en'
    response = requests.get(weather_url)
    soup = BeautifulSoup(response.text, 'html.parser')
    tweet_text = soup.find_all('div', class_='tweet')[0]

    scraped_data['mars_weather'] = tweet_text.find('p', class_='TweetTextSize TweetTextSize--normal js-tweet-text tweet-text').text

    # Mars Facts
    facts_url = 'http://space-facts.com/mars/'
    facts_df = pd.read_html(facts_url)
    facts_transformed = facts_df[0]
    facts_transformed.columns = ['Parameter', 'Value']
    facts_transformed.set_index(['Parameter'])

    scraped_data['facts_html'] = facts_transformed.to_html()

    return scraped_data