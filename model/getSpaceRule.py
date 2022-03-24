# !/usr/bin/python
# -*- coding:utf-8 -*-
# @author   : GaiusPluto
# @time     : 2022/3/24 14:36
from docx.enum.text import WD_LINE_SPACING


def get_space_rule(space_rule) -> WD_LINE_SPACING:
    if type(space_rule) is not WD_LINE_SPACING:
        try:
            SPACE_RULE = {
                '单倍行距': WD_LINE_SPACING.SINGLE,
                '1.5 倍行距': WD_LINE_SPACING.ONE_POINT_FIVE,
                '2 倍行距': WD_LINE_SPACING.DOUBLE,
                '最小值': WD_LINE_SPACING.AT_LEAST,
                '固定值': WD_LINE_SPACING.EXACTLY,
                '多倍行距': WD_LINE_SPACING.MULTIPLE
            }
            space_rule = SPACE_RULE[space_rule]
        except KeyError:
            space_rule = WD_LINE_SPACING.SINGLE
    return space_rule
