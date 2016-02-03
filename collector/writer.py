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

    document = Document('Letter.docx')
    # import ipdb; ipdb.set_trace()
    tbl1 = document.tables[0]
    col1 = tbl1.columns[0]
    # col1.cells[0].paragraphs[0]._p.clear()  # clear cell
    # cpn_box_val = col1.cells[0].add_paragraph()
    # cpn_box_val_format = cpn_box_val.paragraph_format
    # cpn_box_val_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    #
    cpn_box_val_run = col1.cells[0].paragraphs[0].add_run('500 kr*')
    # cpn_box_val_run = cpn_box_val.add_run('600 kr')
    cpn_box_val_font = cpn_box_val_run.font
    cpn_box_val_font.size = Pt(54)
    coupon_val = col1.cells[2].paragraphs[0].add_run('KHTG-KHOP-HLPP-YTFS')
    coupon_val_font = coupon_val.font
    coupon_val_font.size = Pt(18)
    # col1_cell1.tables[0].rows[0].cells[0].text = u'Your coupon: NOW-CHANGED'
    col2 = tbl1.columns[1]
    col2_cell2 = col2.cells[0]
    col2_cell2.tables[0].rows[0].cells[1].text = u'\n\nNEW NAME OF OWNRE\nAddress of owner\n90730 Umeå'

    pars = document.paragraphs
    # # for p in pars:
    # #     p.
    coupon_value = 500

    # p7 = pars[17]
    # p7_text = u'Som en person som vill sälja sitt hus vet du hur krångligt det kan vara att hitta en' + \
    #           u'pålitlig städfirma. Där sticker cleanJoy ut med sin enastående service.' + \
    #           u' Vi har bifogat en kupong med värde {value} kr som en välkomstgåva för din bokning med cleanJoy.se.'.format(
    #               value=coupon_value
    #           )
    # p7.add_run(p7_text)

    # p17 = u''
    # #
    # p1 = document.add_paragraph(u'Se hur cleanjoy.se kan hjälpa dig sälja ditt hus utan krångel.')
    # p2 = document.add_paragraph(u'Hej!')

    # p3 = document.add_paragraph(u'Få ditt hem städat!')
    # p4 = document.add_paragraph(
    #         u'Som en person som vill sälja sitt hus vet du hur krångligt det kan vara att hitta en pålitlig städfirma. Där sticker cleanJoy ut med sin enastående service. Vi har bifogat en kupong med värde 500 kr som en välkomstgåva för din bokning med cleanJoy.se.'
    # )
    #
    # p5 = document.add_paragraph(u'Få dit hem städat av cleanJoy idag')
    # p6 = document.add_paragraph(u'CleanJoy är redo att spara dig både tid och pengar. Så här fungerar det: Gå till cleanjoy.se/aktivera och boka din städning på inom 60 sekunder. Ange koden i kupongfältet när du bokar för att ta tillvara på rabatten. ')
    #
    # p7 = document.add_paragraph(u'Detta är några av fördelarna med att boka din städservice med cleanJoy:')
    # p8 = document.add_paragraph(u'Gratis material')  # bullet
    # p9 = document.add_paragraph(u'Gratis fönsterputtsning')  # bullet
    # p10= document.add_paragraph(u'100% garanti på att du blir nöjd')  # bullet
    # p11 = document.add_paragraph(u'Flexibel bokning och avbokning')  # bullet
    #
    # p12 = document.add_paragraph(u'Boka en flyttstädning idag och få 500 kr rabatt på beställningarna ovan 500 kr. Erbjudandet gäller till den 31 mars 2016.')
    # p13 = document.add_paragraph(u'Ring vår kundtjänst på {tell} så hjälper de dig med frågor angående bokningen och tjänsten. Eller gå till cleanjoy.se/aktivera för att boka din städning. ')
    #
    # p14 = document.add_paragraph(u'Vi ser fram emot att hjälpa dig!')
    # p15 = document.add_paragraph(u'Ulrik Radell')
    # p16 = document.add_paragraph(u'Marketing Manager, CleanJoy Sverige')
    #
    # p17 = document.add_paragraph(u'P.S. Erbjudandet med {value} kr rabatt på din bokning (ovan 2500 kr) går ut den {expiry}.')

    for par in pars:
        par.paragraph_format.right_indent = Inches(-1.3)
    p9 = pars[9]
    # p9_font_color = p9.runs[0].font.color
    # p9_font_color = p9.runs[0].font.color
    p9_appended = p9.add_run(
        u'Vi har bifogat en kupong med värde {value} kr som en välkomstgåva för din bokning med cleanJoy.se.'.format(
            value=500
        ),
    )
    p17 = pars[17]
    p17_appended = p17.add_run(
        u'en flyttstädning idag och få {value} kr rabatt på beställningarna ovan {value} kr. Erbjudandet gäller till den {expiry}.'
        .format(
            value=500,
            expiry='31 mars 2016'
        )


    )
    # p9_appended_font = p9_appended.font
    # p9_appended_font.color = p9_font_color


    document.save('modified.docx')
