# -*- coding: utf-8 -*-
from __future__ import print_function

from datetime import datetime, timedelta
from sqlalchemy.sql.expression import func, and_

from docx import Document
from docx.shared import Inches
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
        print(apt.formal_address , ',', apt.owner, apt.living_area,)


if __name__ == '__main__':
    # write_to_file(date_info_collected=datetime.now().date())
    pass

    body_text = u"""
    Se hur cleanjoy.se kan hjälpa dig sälja ditt hus utan krångel.
    Hej!
    Få ditt hem städat!


    Casey Baumer
    CEO, Your Company
        """.format(value=400)
    print(body_text)

    document = Document('Letter.docx')
    tbl1 = document.tables[0]
    col1 = tbl1.columns[0]
    col1_cell1 = col1.cells[0]
    col1_cell1.tables[0].rows[0].cells[0].text = u'Your coupon: NOW-CHANGED'
    col2 = tbl1.columns[1]
    col2_cell2 = col2.cells[0]
    col2_cell2.tables[0].rows[0].cells[1].text = u'NEW NAME OF OWNRE\nAddress of owner\n90730 Umeå'
    # document.add_heading(u'Se hur cleanjoy.se kan hjälpa dig sälja ditt hus utan krångel.', level=1)
    # document.add_heading(u'Se hur cleanjoy.se kan hjälpa dig sälja ditt hus utan krångel.', level=3)
    # document.add_paragraph(body_text)
    document.save('modified.docx')