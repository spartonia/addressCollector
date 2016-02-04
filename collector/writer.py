# -*- coding: utf-8 -*-
from __future__ import print_function

import os
import random

from datetime import datetime, timedelta
from sqlalchemy.sql.expression import func, and_
from docx import Document
from docx.shared import Inches, Pt
from dateutil import parser as date_parser

from models import DBSession, Apartment, Link


def write_to_one_file(
    coupon,
    coupon_expiry,
    customer_name,
    customer_formal_address,
    coupon_value=500,
    min_order_value=2500,
    doc_template='LetterTemplate.docx',
    file_save_name=None,
    save_path='.'
):

    tell = '070-612 55 55'
    if len(customer_name) > 25:
        name_split = customer_name.split(' ')
        customer_name = name_split[0] + ' ' + name_split[-1]
    customer_name = customer_name
    address_split = customer_formal_address.split(',')
    customer_address = address_split[0]  # u'Karlavägen 68'
    customer_zip_city = address_split[1]  # u'11 459 Stockholm'
    if file_save_name is None:
        file_save_name = coupon + '.docx'

    document = Document(doc_template)
    tbl1 = document.tables[0]
    col1 = tbl1.columns[0]

    cpn_box_val_run = col1.cells[0].paragraphs[0].add_run(str(coupon_value) + ' kr*')
    cpn_box_val_font = cpn_box_val_run.font
    cpn_box_val_font.size = Pt(54)
    coupon_val = col1.cells[2].paragraphs[0].add_run(coupon)
    coupon_val_font = coupon_val.font
    coupon_val_font.size = Pt(18)
    col2 = tbl1.columns[1]
    col2_cell2 = col2.cells[0]
    address_cell = col2_cell2.tables[0].rows[0].cells[1]
    address_cell.text = u'{name}\n{address}\n{zipcity}'.format(
        name=customer_name,
        address=unicode(customer_address),
        zipcity=customer_zip_city
    )
    address_cell_font = address_cell.paragraphs[0].runs[0].font
    address_cell_font.size = Pt(12)

    pars = document.paragraphs
    for par in pars:
        par.paragraph_format.right_indent = Inches(-1.3)

    p7 = pars[7]
    p7_appended = p7.add_run(
        u'Vi har bifogat en kupong med värde {value} kr som en välkomstgåva för din bokning med cleanJoy.se.'.format(
            value=coupon_value
        ),
    )

    p17 = pars[17]
    p17_appended = p17.add_run(
        u'Bok en flyttstädning idag och få {value} kr rabatt på beställningarna ovan {max_value} kr. Erbjudandet gäller till den {expiry}.'
        .format(
            value=coupon_value,
            expiry=coupon_expiry,
            max_value=min_order_value
        )
    )  # bold
    p17_appended_font = p17_appended.font
    p17_appended_font.bold = True

    p18 = pars[18]
    p18_appended = p18.add_run(
        u'Ring vår kundtjänst på {tell} så hjälper de dig med frågor angående bokningen och tjänsten. Eller gå till '.format(
            tell=tell
        )
    )
    p18_appended = p18.add_run(
        u'cleanjoy.se/aktivera'
    )
    p18_appended_font = p18_appended.font
    p18_appended_font.underline = True
    p18_appended = p18.add_run(
        u' för att boka din städning.'
    )


    p25 = pars[25]
    p25_appended = p25.add_run(
        u'Erbjudandet med {value} kr rabatt på din bokning (ovan {min_order_value} kr) går ut den {expiry}.'.format(
           value=coupon_value,
            expiry=coupon_expiry,
            min_order_value=min_order_value
        )
    )
    save_path_name = os.path.join(
        save_path,
        file_save_name
    )
    document.save(save_path_name)


def generate_coupon_code():
        code_chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

        def four_chars():
            four_char = ''
            for i in range(4):
                four_char += random.choice(code_chars)
            return four_char
        coupon_code = 'LTTR-' + '-'.join([four_chars() for i in range(3)])
        return coupon_code


def log_coupons(
    save_path,
    data,
    log_name='coupons.txt'
):
    """
    Writes newly generated coupons to a file along with other info, e.g expiry, etc.

    Parameters
    ----------
    save_path : str
    log_name : str
    data : str
    """

    log_file = os.path.join(
        save_path,
        log_name,
    )
    with open(log_file, mode='a') as log:
        log.write(data + '\n')


def write_to_file(
    save_path='/media/spartonia/DATA/Company_stuff/Marketing/Letters',
    published_days_ago=30,
    date_info_collected=str((datetime.now().date() - timedelta(days=-1))),  # yesterday
    price_range=(0, 20000000),
    living_area_range=(0, 500)
):
    # TODO: add county selection option
    """
    Writes marketing letters.

    Parameters
    ----------
    save_path: str
    published_days_ago:  int
    date_info_collected: str 'YYYY-mm-dd'
    price_range: (int, int)
    living_area_range: (int, int)

    """
    reporting_day_save_path = os.path.join(
        save_path,
        str(datetime.now().date())
    )
    if not os.path.exists(reporting_day_save_path):
        os.makedirs(reporting_day_save_path)

    save_path_reporting_day = os.path.join(
        reporting_day_save_path,
        date_info_collected
    )
    if not os.path.exists(save_path_reporting_day):
        os.mkdir(save_path_reporting_day)

    date_collected = date_parser.parse(date_info_collected).date()
    session = DBSession()
    apts = session.query(Apartment).join(Link).filter(
        Apartment.processed.isnot(True),
        Apartment.owner.isnot(None),
        and_(Apartment.price >= price_range[0], Apartment.price <= price_range[1]),
        and_(Apartment.living_area >= living_area_range[0], Apartment.living_area <= living_area_range[1]),
        Apartment.price.isnot(None),
        Apartment.postal_code.isnot(None),
        Link.timestamp.between(
            str(date_collected), str(date_collected + timedelta(days=1))),
        Link.date.between(str(datetime.now().date() + timedelta(days=-10)), str(datetime.now().date())),
    )
    file_count = 0
    for i, apt in enumerate(apts):
        coupon_code = generate_coupon_code()
        expiry = str(datetime.now().date() + timedelta(days=60))
        print(apt.owner)
        try:
            write_to_one_file(
                coupon=coupon_code,
                coupon_expiry=expiry,
                customer_name=apt.owner,
                customer_formal_address=apt.formal_address,
                save_path=save_path_reporting_day
            )
            log_data = coupon_code + ';' + expiry
            log_coupons(
                reporting_day_save_path,
                data=log_data
            )
            file_count += 1
            apt.processed = True
            try:
                session.commit()
            except Exception as e:
                print(e.message)
                session = DBSession()
        except:
            print('Error creating letter for', apt.formal_address, ',', apt.owner, coupon_code, expiry)

    print('Done!')
    print(file_count, 'letters created.')


if __name__ == '__main__':
    # write_to_file(date_info_collected='2016-02-04')
    pass

    # coupon_code = generate_coupon_code()
    # print(coupon_code)