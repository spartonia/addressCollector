# -*- coding: utf-8 -*-
from __future__ import print_function

from crawler import update_links
from writer import write_to_file
from collector import db

if __name__ == '__main__':

    stockholm_villa_feed_url = 'http://www.hemnet.se/mitt_hemnet/sparade_sokningar/9201235.xml'
    uppsala_villa_feed_url = 'http://www.hemnet.se/mitt_hemnet/sparade_sokningar/9251594.xml'
    rs = update_links(stockholm_villa_feed_url)

    collection = db.test_collection
    collection.insert(rs)
