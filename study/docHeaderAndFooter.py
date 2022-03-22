# !/usr/bin/python
# -*- coding:utf-8 -*-
# @author   : GaiusPluto
# @time     : 2022/3/18 20:12
from docx import Document
# 访问节的标题
document = Document()
section = document.sections[0]
header = section.header

# 添加页眉
paragraph = header.paragraphs[0]
paragraph.text = "Title of my document"
header.is_linked_to_previous = False

# 添加分区标题内容
paragraph.text = "Left Text\tCenter Text\tRight Text"
paragraph.style = document.styles["Header"]

# 移除页眉
header.is_linked_to_previous = True

document.save("test_headerandfooter.docx")
