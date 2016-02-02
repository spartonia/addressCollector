# -*- coding: utf-8 -*-
from __future__ import print_function

import os

from sqlalchemy import Column, String, DateTime, TIMESTAMP, Integer, ForeignKey, UnicodeText, Boolean
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import create_engine
from sqlalchemy.sql import func
from sqlalchemy.orm import sessionmaker, relationship, backref

Base = declarative_base()


class Link(Base):
    __tablename__ = 'link'
    id = Column(Integer, primary_key=True)
    url = Column(String(250), unique=True)
    date = Column(DateTime)
    timestamp = Column(TIMESTAMP, default=func.now())
    # apartment = relationship('Apartment', uselist=False, backref='link')


class Apartment(Base):
    __tablename__ = 'apartment'
    id = Column(Integer, primary_key=True)
    link_id = Column(Integer, ForeignKey('link.id'), nullable=False)
    street = Column(UnicodeText(50))  # Karlavagen
    street_no = Column(String(3))  # 68
    locality = Column(UnicodeText(50), nullable=True)  # Ostermalm sublocality in gmap
    kommun = Column(UnicodeText(50))  # Stockholm , locality in gmap
    postal_code = Column(String(6))
    formal_address = Column(UnicodeText(100))
    price = Column(Integer, nullable=True)
    living_area = Column(Integer, default=0)
    total_area = Column(Integer, default=0)
    rooms = Column(Integer, default=0)
    year_of_construction = Column(String(4), nullable=True)
    avgift = Column(Integer, nullable=True)
    type = Column(UnicodeText(25), nullable=True)

    owner = Column(UnicodeText(100), nullable=True)
    broker_email = Column(String(30), nullable=True)
    broker_url = Column(String(50), nullable=True)

    processed = Column(Boolean, default=False)

    link = relationship(
        Link,
        backref=backref('apartment', uselist=False)
    )

engine = create_engine(
    'sqlite:///' +
    os.path.join(
        '../bin',
        'database.db'
    )
)
engine.echo = False

Base.metadata.create_all(engine)

DBSession = sessionmaker(bind=engine)