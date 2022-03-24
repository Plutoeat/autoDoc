# !/usr/bin/python
# -*- coding:utf-8 -*-
# @author   : GaiusPluto
# @time     : 2022/3/24 11:48
def get_font_size(font_size) -> int:
    if type(font_size) is not int:
        try:
            FONT_SIZE = {
                '初号': 42,
                '小初': 36,
                '一号': 26,
                '小一': 24,
                '二号': 22,
                '小二': 18,
                '三号': 16,
                '小三': 15,
                '四号': 14,
                '小四': 12,
                '五号': 10.5,
                '小五': 9,
                '六号': 7.5,
                '小六': 6.5,
                '七号': 5.5,
                '八号': 5
            }
            font_size = FONT_SIZE[font_size]
        except KeyError:
            font_size = 16
    return font_size
