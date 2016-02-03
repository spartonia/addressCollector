# -*- coding: utf-8 -*-
from __future__ import print_function
from datetime import datetime, timedelta
from sqlalchemy.sql.expression import func, and_
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from models import DBSession, Apartment, Link


def write_to_file(
    save_path='',
    published_days_ago=30,
    date_info_collected=(datetime.now().date() - timedelta(days=-1)),  # yesterday
    price_range=(0, 20000000),
    living_area_range=(0, 500)
):
    session = DBSession()
    apts = session.query(Apartment).join(Link).filter(
        Apartment.processed.isnot(True),
        Apartment.owner.isnot(None),
        and_(Apartment.price >= price_range[0], Apartment.price <= price_range[1]),
        and_(Apartment.living_area >= living_area_range[0], Apartment.living_area <= living_area_range[1]),
        Apartment.price.isnot(None),
        Apartment.postal_code.isnot(None),
        Link.timestamp.between(
            str(date_info_collected), str(date_info_collected + timedelta(days=1))),
        Link.date.between(str(datetime.now().date() + timedelta(days=-10)), str(datetime.now().date())),
    )
    for apt in apts:
        print(apt.formal_address, ',', apt.owner, apt.living_area, )


if __name__ == '__main__':
    # write_to_file(date_info_collected=datetime.now().date())
    coupon_value = 500
    expiry = u'31 mars 2016'
    tell = '070-612 55 55'
    customer_name = u'Saeed Partonia'
    customer_address = u'Karlavägen 68'
    customer_zip_city = u'11 459 Stockholm'


    document = Document('Letter.docx')
    tbl1 = document.tables[0]
    col1 = tbl1.columns[0]

    cpn_box_val_run = col1.cells[0].paragraphs[0].add_run(str(coupon_value) + ' kr*')
    cpn_box_val_font = cpn_box_val_run.font
    cpn_box_val_font.size = Pt(54)
    coupon_val = col1.cells[2].paragraphs[0].add_run('KHTG-KHOP-HLPP-YTFS')
    coupon_val_font = coupon_val.font
    coupon_val_font.size = Pt(18)
    col2 = tbl1.columns[1]
    col2_cell2 = col2.cells[0]
    col2_cell2.tables[0].rows[0].cells[1].text = u'\n\n\n\n{name}\n{address}\n{zipcity}'.format(
        name = customer_name,
        address = customer_address,
        zipcity = customer_zip_city
    )

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
        u'Bok en flyttstädning idag och få {value} kr rabatt på beställningarna ovan {value} kr. Erbjudandet gäller till den {expiry}.'
        .format(
            value=coupon_value,
            expiry=expiry
        )
    )  # bold

    p18 = pars[18]
    p18_appended = p18.add_run(
        u'Ring vår kundtjänst på {tell} så hjälper de dig med frågor angående bokningen och tjänsten. Eller gå till cleanjoy.se/aktivera för att boka din städning.'.format(
            tell=tell
        )
    )

    p25 = pars[25]
    p25_appended = p25.add_run(
        u'Erbjudandet med {value} kr rabatt på din bokning (ovan 2500 kr) går ut den {expiry}.'.format(
           value=coupon_value,
            expiry=expiry
        )
    )

    document.save('modified.docx')
