#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Apr 13 16:43:23 2021

@author: zhangmuhan
"""

import mysql.connector

mydb = mysql.connector.connect(
  host = "123.56.29.216",
  user= "testadmin",
  password="Aa_11111",
  database = "testDB"
)


mycursor = mydb.cursor()


