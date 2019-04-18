#!/usr/bin/python
# _*_ coding:UTF-8 _*_
import xlrd
import xlwt
import os
import time

print('wenlei', __name__)


class ExcelExe:
    def __init__(self):
        self.wb = xlwt.Workbook()

    def set_style(self, name, height, bold=False, horz=xlwt.Alignment.HORZ_CENTER, border=False):
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.name = name
        font.height = height
        font.bold = bold
        alignment = xlwt.Alignment()
        alignment.horz = horz
        alignment.vert = xlwt.Alignment.VERT_CENTER
        alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT
        if border:
            border = xlwt.Borders()
            border.left = xlwt.Borders.THIN
            border.right = xlwt.Borders.THIN
            border.bottom = xlwt.Borders.THIN
            border.top = xlwt.Borders.THIN
            border.diag = xlwt.Borders.THIN
            style.borders = border
        style.font = font
        style.alignment = alignment

        return style

    def write(self, obj):
        print(obj)
        fsgb = u'仿宋_GB2312'
        sheet = self.wb.add_sheet(obj['userName'])
        # sheet.col(0).height=
        sheet.write_merge(0, 0, 0, 5, 'DJDX20190000', self.set_style('Cambria', 16 * 20, False, xlwt.Alignment.HORZ_RIGHT))
        sheet.write_merge(1, 1, 0, 5, u'中国大唐集团公司培训项目综合评价表', self.set_style('Times New Roman', 22 * 20, True))
        title = self.set_style(name=fsgb, height=16 * 20, bold=True)
        sheet.write(2, 0, u'年度：', title)
        sheet.write(2, 1, u'2019年', title)
        sheet.write(2, 4, u'岗位：', title)
        sheet.write(2, 5, obj['job'], title)

        style = self.set_style(name=fsgb, height=12 * 20, bold=True, border=True)
        sheet.write(3, 0, u'姓名', style)
        sheet.write(3, 1, obj['userName'], style)
        sheet.write(3, 2, u'性别', style)
        sheet.write(3, 3, obj['sex'], style)
        sheet.write(3, 4, u'身份证号', style)
        sheet.write(3, 5, str(obj['idCard']), style)

        sheet.write(4, 0, u'技术资格', style)
        sheet.write_merge(4, 4, 1, 2, obj['tech'], style)
        sheet.write(4, 3, u'技能等级', style)
        sheet.write_merge(4, 4, 4, 5, obj['skillLevel'], style)

        sheet.write(5, 0, u'二级企业名称', style)
        sheet.write_merge(5, 5, 1, 2, obj['company'], style)
        sheet.write(5, 3, u'所在企业名称', style)
        sheet.write_merge(5, 5, 4, 5, obj['companyName'], style)

        sheet.write(6, 0, u'所在部门', style)
        sheet.write_merge(6, 6, 1, 2, obj['departmentName'], style)
        sheet.write(6, 3, u'现从事岗位', style)
        sheet.write_merge(6, 6, 4, 5, obj['job'], style)

        sheet.write(7, 0, u'培训地点', style)
        sheet.write(7, 1, '', style)
        sheet.write(7, 2, u'培训时间', style)
        sheet.write(7, 3, '', style)
        sheet.write(7, 4, u'总  学  时', style)
        sheet.write(7, 5, '', style)

        sheet.write(8, 0, u'培训考核成绩', style)
        sheet.write_merge(8, 8, 1, 2, '', style)
        sheet.write(8, 3, u'合格分数线', style)
        sheet.write_merge(8, 8, 4, 5, u'60分', style)

        sheet.write_merge(9, 15, 0, 0, u'岗位\n（工种）\n建议', style)
        sheet.write_merge(9, 15, 1, 5, u'\n    根据    规定，该同志在集团公司2019年    培训项目中达到考核标准，企业可依据实际情况，在上岗时予以参考。\n'
                                       u'                    中国大唐集团公司培训专用章'
                                       u'\n                      2019年  月  日', style)

    def execute(self, file_obj, source='source'):
        path = './static/' + str(round(time.time() * 1000)) + '.xls'
        if os.path.exists(path):
            os.unlink(path)
        if file_obj:
            workbook = xlrd.open_workbook(file_contents=file_obj.read())
        else:
            file_url = './' + source + '.xls'
            workbook = xlrd.open_workbook(file_url)

        sh = workbook.sheet_by_index(0)
        total = sh.nrows
        excel_list = []
        for row in range(total - 5):
            if row > 2:
                columns = sh.row_values(row)
                if columns[1]:
                    row_object = {
                        'company': columns[1],
                        'companyName': columns[2],
                        'departmentName': columns[3],
                        'job': columns[4],
                        'userName': columns[6],
                        'sex': columns[7],
                        'idCard': columns[8],
                        'tech': columns[11],
                        'skillLevel': columns[12],
                    }
                    self.write(row_object)
                    excel_list.append(row_object)

        self.wb.save(path)
        return path


if __name__ == '__main__':
    excel = ExcelExe()
    excel.execute()
    print('success!')
