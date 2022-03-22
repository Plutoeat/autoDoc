# !/usr/bin/python
# -*- coding:utf-8 -*-
# @author   : GaiusPluto
# @time     : 2022/3/18 16:42
from docx import Document

# 访问章节
document = Document()
sections = document.sections
for section in sections:
    print(section.start_type)

# 添加章节
# from docx.enum.section import WD_SECTION_START

current_section = sections[-1]
print(current_section.start_type)
# new_section = document.add_section(WD_SECTION_START.ODD_PAGE)
# print(new_section.start_type)

# 页面属性
from docx.enum.section import WD_ORIENTATION

section = sections[0]
new_width = section.page_height
new_height = section.page_width
section.orientation = WD_ORIENTATION.LANDSCAPE
section.page_width = new_width
section.page_height = new_height
print(section.left_margin, section.right_margin, section.top_margin, section.bottom_margin)
document.save("test_section.docx")
