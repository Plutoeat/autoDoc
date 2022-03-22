# !/usr/bin/python
# -*- coding:utf-8 -*-
# @author   : GaiusPluto
# @time     : 2022/3/19 16:24
from .docx_schema import *


def savedoc(title: str, doctype: str) -> bool:
    if doctype == 'paper':
        doc = Paper(title)
        doc.save()
    return True
