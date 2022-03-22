# !/usr/bin/python
# -*- coding:utf-8 -*-
# @author   : GaiusPluto
# @time     : 2022/3/19 16:54
from docx import Document
from docx.enum.section import WD_ORIENTATION, WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_LINE_SPACING, WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from model.SETTINGS import *


class DocxSchema:
    def __init__(self, title):
        """
        All rules come from Dalian University
        :param title: doc's title
        """
        self.title = title
        self.document = Document()
        self.sections = self.document.sections
        # 采取A4纸张
        if PAPER_TYPE == 'A4':
            self.sections[0].page_height = Cm(29.7)
            self.sections[0].page_width = Cm(21)
        # 默认横向，有特殊需求则为纵向
        if PAPER_DIRECTION_LANDSCAPE:
            self.sections[0].orientation = WD_ORIENTATION.LANDSCAPE

    def save(self):
        self.document.save("%s.docx" % self.title)


class Paper(DocxSchema):
    def __init__(self, title):
        super().__init__(title)
        # 页面设置
        # 设置页眉
        self.document.sections[0].header.paragraphs[0].text = HEADER_TEXT
        self.document.sections[0].header.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        self.document.sections[0].header.paragraphs[0].style.font.size = Pt(10.5)
        self.document.sections[0].header_distance = Cm(HEADER_DISTANCE)
        # 设置页脚，页码需要自行设置
        self.document.sections[0].footer_distance = Cm(FOOTER_DISTANCE)
        # 页边距一律采取
        self.sections[0].top_margin = Cm(TOP_MARGIN)
        self.sections[0].bottom_margin = Cm(BOTTOM_MARGIN)
        self.sections[0].left_margin = Cm(LEFT_MARGIN)
        self.sections[0].right_margin = Cm(RIGHT_MARGIN)

        # 设置摘要样式
        style = self.document.styles.add_style('Abstract', WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = self.document.styles['Heading 1']
        style.hidden = False
        style.quick_style = True
        style.paragraph_format.first_line_indent = Pt(0)
        style.paragraph_format.line_spacing = 1.5
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.space_after = Pt(11)
        self.document.styles['Abstract'].font.name = '黑体'
        self.document.styles['Abstract'].element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        self.document.styles['Abstract'].font.size = Pt(16)
        self.document.styles['Abstract'].font.bold = False
        self.document.styles['Abstract'].font.color.rgb = RGBColor(0, 0, 0)

        # 标题格式
        style = self.document.styles.add_style('论文标题 1', WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = self.document.styles['Heading 1']
        style.hidden = False
        style.quick_style = True
        style.paragraph_format.first_line_indent = Pt(0)
        style.paragraph_format.line_spacing = 1.5
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.space_after = Pt(11)
        self.document.styles['论文标题 1'].font.name = '黑体'
        self.document.styles['论文标题 1'].element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        self.document.styles['论文标题 1'].font.size = Pt(16)
        self.document.styles['论文标题 1'].font.bold = False
        self.document.styles['论文标题 1'].font.color.rgb = RGBColor(0, 0, 0)

        style = self.document.styles.add_style('论文标题 2', WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = self.document.styles['Heading 2']
        style.hidden = False
        style.quick_style = True
        style.paragraph_format.line_spacing = 1.5
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        style.paragraph_format.space_before = Pt(7)
        style.paragraph_format.space_after = Pt(0)
        style.paragraph_format.first_line_indent = Pt(28)
        self.document.styles['论文标题 2'].font.name = '黑体'
        self.document.styles['论文标题 2'].element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        self.document.styles['论文标题 2'].font.size = Pt(14)
        self.document.styles['论文标题 2'].font.bold = False
        self.document.styles['论文标题 2'].font.color.rgb = RGBColor(0, 0, 0)

        style = self.document.styles.add_style('论文标题 3', WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = self.document.styles['Heading 3']
        style.hidden = False
        style.quick_style = True
        style.paragraph_format.line_spacing = 1.5
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        style.paragraph_format.space_before = Pt(6)
        style.paragraph_format.space_after = Pt(0)
        style.paragraph_format.first_line_indent = Pt(24)
        self.document.styles['论文标题 3'].font.name = '黑体'
        self.document.styles['论文标题 3'].element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        self.document.styles['论文标题 3'].font.size = Pt(12)
        self.document.styles['论文标题 3'].font.bold = False
        self.document.styles['论文标题 3'].font.color.rgb = RGBColor(0, 0, 0)

        # 设置正文样式
        self.document.styles['Normal'].hidden = False
        self.document.styles['Normal'].quick_style = True
        self.document.styles['Normal'].font.name = 'Times New Roman'
        self.document.styles['Normal'].element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        self.document.styles['Normal'].font.size = Pt(12)
        self.document.styles['Normal'].paragraph_format.space_before = Pt(0)
        self.document.styles['Normal'].paragraph_format.space_after = Pt(0)
        self.document.styles['Normal'].paragraph_format.first_line_indent = Pt(24)
        # 行间距固定值(设置值为20),
        self.paragraph_format = self.document.styles['Normal'].paragraph_format
        self.paragraph_format.line_spacing = Pt(20)
        self.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY

        # 关键词样式
        style = self.document.styles.add_style('KeyWord', WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = self.document.styles['Normal']
        style.font.name = '黑体'
        style.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        style.font.size = Pt(12)
        style.font.bold = False
        style.paragraph_format.first_line_indent = Pt(0)

        # 中文摘要
        self.document.add_paragraph('摘    要', style='Abstract').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        self.document.add_paragraph('请在此输入你的正文', style='Normal')
        self.document.add_paragraph('关键词：关键词1；关键词2；关键词3', style='KeyWord').alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        # 英文摘要
        self.document.add_section(WD_SECTION_START.ODD_PAGE)
        self.document.add_paragraph('Abstract', style='Abstract').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        self.document.add_paragraph('Please enter your text here', style='Normal')
        self.document.add_paragraph('Key Words：Key Words1；Key Words2；Key Words3',
                                    style='KeyWord').alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        # 添加目录页
        self.document.add_section(WD_SECTION_START.ODD_PAGE)
        self.document.add_paragraph('目    录', style='Abstract').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # 添加绪论
        self.document.add_section(WD_SECTION_START.ODD_PAGE)
        self.document.add_paragraph('绪    论', style='Abstract').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        self.document.add_paragraph('请在此输入你的绪论', style='Normal')
        self.document.add_paragraph('Please enter your introduction here', style='Normal')

        # 添加标题及正文
        self.document.add_section(WD_SECTION_START.ODD_PAGE)
        self.document.add_paragraph('1 请在此输入你的论文章标题', style='论文标题 1')
        self.document.add_paragraph('1.1 请在此输入你的论文节标题', style='论文标题 2')
        self.document.add_paragraph('1.1.1 请在此输入你的节中一级标题', style='论文标题 3')
        self.document.add_paragraph('请在此输入你的论文正文', style='Normal')

        # 添加结论
        self.document.add_section(WD_SECTION_START.ODD_PAGE)
        self.document.add_paragraph('结    论', style='论文标题 1').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        paragraph_format = self.document.add_paragraph('请在此输入你的结论正文', style='Normal').paragraph_format
        paragraph_format.line_spacing = 1.25
        paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(0)

        # 添加参考文献
        self.document.add_section(WD_SECTION_START.ODD_PAGE)
        paragraph = self.document.add_paragraph('', style='论文标题 1')
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = paragraph.add_run('参 考 文 献')
        run.font.size = Pt(12)
        paragraph_format = self.document.add_paragraph('请在此填写参考文献的著录').paragraph_format
        paragraph_format.line_spacing = 1.25
        paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(0)

        # 添加附录
        self.document.add_section(WD_SECTION_START.ODD_PAGE)
        paragraph = self.document.add_paragraph('', style='论文标题 1')
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = paragraph.add_run('附录1 附录内容')
        run.font.size = Pt(15)
        paragraph_format = self.document.add_paragraph('请在此输入附录正文').paragraph_format
        paragraph_format.line_spacing = 1.3
        paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(0)
