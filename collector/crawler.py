# -*- coding: utf-8 -*-
from __future__ import print_function, division

import sys
reload(sys)
sys.setdefaultencoding("utf-8")

import random
import urllib
import re
import requests
import feedparser
import json
import mechanize
import cookielib
import urllib
import random
import time

from lxml import html
from lxml.etree import tostring
from datetime import datetime
from dateutil import parser as date_parser
from sqlalchemy.sql import desc
from bs4 import BeautifulSoup, UnicodeDammit
from sqlalchemy.exc import IntegrityError


from models import DBSession, Link, Apartment

import warnings
warnings.filterwarnings('ignore')

USER_AGENTS_FILE = '../user_agents.txt'
# http://<USERNAME>:<PASSWORD>@<IP-ADDR>:<PORT>
proxy = {"http": "http://pptp:jw6YhEcb@176.126.237.207"}

GMAP_API_KEY = 'AIzaSyA6Hzu7AUwTqTne_uYq4xX5EJJjhbt-B1k'
GMAP_REQUEST_URL = \
    'https://maps.googleapis.com/maps/api/geocode/json?address={address}+SE&key=AIzaSyA6Hzu7AUwTqTne_uYq4xX5EJJjhbt-B1k'


def LoadUserAgents(uafile=USER_AGENTS_FILE):
    """
    uafile : string
        path to text file of user agents, one per line
    """
    uas = []
    with open(uafile, 'rb') as uaf:
        for ua in uaf.readlines():
            if ua:
                uas.append(ua.strip()[1:-1-1])
    random.shuffle(uas)
    return uas

uas = LoadUserAgents()

ua = random.choice(uas)


def logged_in_browser(
    ua=random.choice(uas)
):
    """
    Returns a browser with logged in credentials. This will prevent multiple logging blocking
    from host website (merinfo).The logged in browser can be used for searching addresses multiple times.
    """
    login_url = 'http://www.merinfo.se/user/login'
    cj = cookielib.CookieJar()
    br = mechanize.Browser()
    br.set_handle_robots(False)
    br.set_handle_equiv(False)
    br.addheaders=[('User-agent', ua)]
    br.set_cookiejar(cj)

    br.open(login_url)
    br.select_form(nr=1)
    br.form['email'] = 'sd_prtn@yahoo.com'
    br.form['password'] = 'saeedpp'
    br.submit()
    return br


def crawl_name(
    where_,
    browser,
    proxy=None
):
    base_url = 'http://www.merinfo.se/search?'

    search_params = {
        'where': where_,
        'who': ''
    }
    url = base_url + urllib.urlencode(search_params)

    browser.open(url)
    page = browser.response().read()

    name_xpath = '//*[@id="result-list"]/div[1]/div/div[1]/a[1]/h2'
    tree = html.fromstring(UnicodeDammit(page).unicode_markup)
    name = tree.xpath(name_xpath)[0]
    soup = BeautifulSoup(tostring(name), 'lxml')
    return unicode(''.join(soup.findAll(text=True)))


def address_from_gmaps(
    address
):
    """
    Parameters
    ----------
    address : str
        e.g: 'Karlavagen 68 Stockholm'
    Returns
    -------
    address_fields : dict
    """
    gmap_api_url = GMAP_REQUEST_URL.format(address=address)

    address_fields = {
        'street_number': None,
        'route': None,
        'sublocality': None,
        'locality': None,
        'administrative_area_level_1': None,
        'postal_code': None,
    }
    map_api_to_apartment_filed_names = {
        'street_number': 'street_no',
        'route': 'street',
        'sublocality': 'locality',
        'locality': 'kommun',
        'formatted_address': 'formal_address',
        'postal_code': 'postal_code',
    }

    r = requests.get(gmap_api_url)
    if r.status_code == 200:
        rs = json.loads(r.text)
        # print(json.dumps(rs, indent=4))
        if rs['status'] != 'OK':
            print('Status = ', rs['status'])
            return {}
        result = rs['results'][0]
        for component in result['address_components']:
            for field_name in address_fields.keys():
                if field_name in component['types']:
                    address_fields[field_name] = component['short_name']
        formatted_address = ', '.join(result['formatted_address'].split(',')[:-1])
        address_fields['formatted_address'] = formatted_address
        result_set = {}
        for api_key, db_key in map_api_to_apartment_filed_names.iteritems():
            result_set[db_key] = address_fields[api_key]
        return result_set


def crawl_hemnet_page(URL, browser, ua=random.choice(uas)):
    """
    Crawls the specified page url on hemnet.
    Note; Only for villa, has to be improve to support other types.
    Returns
    -------
    result_set : dict
    """
    result_set = {}
    URL = URL
    headers = {
        "Connection": "close",
        "User-Agent": ua
    }
    address_xpaths = [
        ('address',  '//*[@id="item-info"]/div[2]/div[2]/div[1]/div[1]/h1/span/text()'),
        ('address_no', '//*[@id="item-info"]/div[2]/div[2]/div[1]/div[1]/h1/text()'),
        # ('locality', '//*[@id="item-info"]/div[2]/div[2]/div[1]/div[1]/p/text()'),
        ('kommun', '//*[@id="item-info"]/div[2]/div[2]/div[1]/div[1]/p/span/text()'),
    ]
    price_xpath = '//*[@id="item-info"]/div[2]/div[2]/div[1]/div[2]/span/text()'
    price_regex = '\d+'
    fact_table_key_to_regex = {
        'Bostadstyp' : '.*',
        'Boarea': '\d+',
        'rum': '\d+',
        'Avgift': '\d+\s\d+',
        'Driftkostnad': '\d+\s\d+',
        'Byg': '\d+'
    }

    page = requests.get(URL, headers=headers)
    tree = html.fromstring(UnicodeDammit(page.content).unicode_markup)
    address_from_hemnet = ''
    for item in address_xpaths:
        address_from_hemnet += unicode(tree.xpath(item[1])[0].strip()) + ' '
    try:
        gmap_res = address_from_gmaps(address_from_hemnet)
        result_set.update(gmap_res)
    except Exception as e:
        print('Error retrieving gmap info for:' )
        print(URL)
        print(e.message)
        pass

    try:
        owner_name = crawl_name(gmap_res['formal_address'], browser=browser)
        result_set.update({'owner': owner_name})
    except Exception as e:
        # consecutive_failed_tries += 1
        # if consecutive_failed_tries > 3
        #   wait(a_big_time)
        # if index error, try crawling eniro

        print('Error retrieving owner name for:' )
        print('Url', URL)
        print('Address:', gmap_res['formal_address'])
        print(e.message)
    try:
        price_raw = tree.xpath(price_xpath)[0].strip()
        price = ''.join(re.findall(price_regex, price_raw))
        result_set.update({'price': price})
    except:
        pass

    soup = BeautifulSoup(tostring(tree), 'lxml')
    fact_table = soup.findAll('dl')[0]
    dts = fact_table.findChildren('dt')
    processed_facts = {}
    for dt in dts:
        for key, regex in fact_table_key_to_regex.iteritems():
            if key in dt.contents[0].strip():
                # print(key, dt.nextSibling.nextSibling.contents[0].strip())
                processed_facts[key] = re.findall(regex, dt.nextSibling.nextSibling.contents[0].strip(), re.UNICODE)[0]

    fact_dict = {}
    for key, value in processed_facts.iteritems():
        if key == 'Bostadstyp':
            fact_dict['type'] = unicode(value)
        elif key == 'Boarea':
            fact_dict['living_area'] = int(value.split(',')[0])
        elif key == 'rum':
            fact_dict['rooms'] = int(value.split(',')[0])
        elif key == 'Avgift' or key == 'Driftkostnad':
            val = value.replace(u'\xa0', u'')
            fact_dict['avgift'] = int(val)
        elif key == 'Byg':
            fact_dict['year_of_construction'] = int(value)
        elif key == 'Tomtarea':
            pass

    result_set.update(fact_dict)

    return result_set


def update_links(
    rss_feed_url
):
    """Inserts new announcements into the database.
    Notes:
        * Publishing date is the date the property is listed for sale. It might be very old,
        * We insert all entries in the rss feed to the database. The url field is unique so duplicates are not allowed.
        * When querying new entries, keep in mind to query the date based on ( timestamp = 'today' & pubDate =
        'close enough' )so only new listed properties are queried.

    Parameters
    ----------
    rss_feed_url : str

    """
    feed = feedparser.parse(rss_feed_url)

    entries = feed['entries']

    num_new_links = 0

    print('Updating liks database ..')
    result_set = []
    browser = logged_in_browser()
    for cnt, entry in enumerate(entries):
        link = entry['link']
        print(cnt, ':', link)

        time.sleep(random.choice(range(20, 60))/10)
        try:
            data = crawl_hemnet_page(link, browser=browser)
            data.update({'link': link})
            num_new_links += 1
        except Exception as e:
            print('Error crawling hemnet page.', e.message)
            continue
        if num_new_links == 10:
            break
        result_set.append(data)

    print('Done!')
    print('%s new links added.' % num_new_links)
    return result_set


if __name__ == '__main__':
    stockholm_vill_feed_url = 'http://www.hemnet.se/mitt_hemnet/sparade_sokningar/9201235.xml'
    # update_links(stockholm_vill_feed_url)
    hemnet_url = 'http://www.hemnet.se/bostad/villa-5rum-adelso-ekero-kommun-marielundsvagen-87-8975943'
    br = logged_in_browser()
    rs = crawl_hemnet_page(hemnet_url, browser=br)
    from pprint import pprint
    pprint(rs)



