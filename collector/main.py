# -*- coding: utf-8 -*-
from __future__ import print_function

from crawler import update_links
from writer import write_to_file
if __name__ == '__main__':

    stockholm_villa_feed_url = 'http://www.hemnet.se/mitt_hemnet/sparade_sokningar/9201235.xml'
    uppsala_villa_feed_url = 'http://www.hemnet.se/mitt_hemnet/sparade_sokningar/9251594.xml'
    # update_links(uppsala_villa_feed_url)

    write_to_file(date_info_collected='2016-02-04')

    # convert files to pdf
    # $ cd to-path
    # $ lowriter --convert-to pdf *.docx
    # then merge pdfs into one file
    # $ pdftk `ls *.pdf` cat output merged.pdf