# !/usr/bin/python
# -*- coding:utf-8 -*-
# @author   : GaiusPluto
# @time     : 2022/3/24 12:05
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


def get_paragraph_alignment(align_style) -> WD_PARAGRAPH_ALIGNMENT:
    if type(align_style) is not WD_PARAGRAPH_ALIGNMENT:
        try:
            ALIGN_STYLE = {
                '左对齐': WD_PARAGRAPH_ALIGNMENT.LEFT,
                '居中': WD_PARAGRAPH_ALIGNMENT.CENTER,
                '右对齐': WD_PARAGRAPH_ALIGNMENT.RIGHT,
                '两端对齐': WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
                '分散对齐': WD_PARAGRAPH_ALIGNMENT.DISTRIBUTE
            }
            align_style = ALIGN_STYLE[align_style]
        except KeyError:
            align_style = WD_PARAGRAPH_ALIGNMENT.LEFT
    return align_style
