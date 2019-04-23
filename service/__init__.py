#!/usr/bin/python
# _*_ coding:UTF-8 _*_
import xlrd
import xlwt
import os
import time


class ExcelExe:
    def __init__(self):
        self.wb = xlwt.Workbook()

    def set_style(self, name, height, bold=False, horz=xlwt.Alignment.HORZ_CENTER, border=False,
                  part_border={'left': True, 'right': True, 'top': True, 'bottom': True}):
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
            border.left = xlwt.Borders.THIN if (part_border.get('left')) else xlwt.Borders.NO_LINE
            border.right = xlwt.Borders.THIN if (part_border.get('right')) else xlwt.Borders.NO_LINE
            border.bottom = xlwt.Borders.THIN if (part_border.get('bottom')) else xlwt.Borders.NO_LINE
            border.top = xlwt.Borders.THIN if (part_border.get('top')) else xlwt.Borders.NO_LINE
            # border.diag = xlwt.Borders.THIN
            style.borders = border
        style.font = font
        style.alignment = alignment

        return style

    def write(self, obj):
        print(obj)
        fsgb = u'仿宋_GB2312'
        sheet = self.wb.add_sheet(obj['userName'])

        sheet.col(0).width = int(256 * (95 / 6))
        sheet.col(1).width = int(256 * (65 / 6))
        sheet.col(2).width = int(256 * (65 / 6))
        sheet.col(3).width = int(256 * (95 / 6))
        sheet.col(4).width = int(256 * (65 / 6))
        sheet.col(5).width = int(256 * (125 / 6))

        height40 = xlwt.easyxf('font:height 640;')
        sheet.row(0).set_style(height40)
        height60 = xlwt.easyxf('font:height 960;')
        sheet.row(1).set_style(height60)
        height30 = xlwt.easyxf('font:height 480;')
        sheet.row(2).set_style(height30)
        height50 = xlwt.easyxf('font:height 800;')
        sheet.row(3).set_style(height50)
        sheet.row(4).set_style(height40)
        sheet.row(5).set_style(height40)
        sheet.row(6).set_style(height40)
        sheet.row(7).set_style(height40)
        sheet.row(8).set_style(height40)

        height140 = xlwt.easyxf('font:height ' + str(140 * 16) + ';')
        sheet.row(9).set_style(height140)
        height25 = xlwt.easyxf('font:height ' + str(25 * 16) + ';')
        sheet.row(10).set_style(height25)
        sheet.row(11).set_style(height25)

        sheet.write_merge(0, 0, 0, 5, 'DJDX20190000',
                          self.set_style('Cambria', 20 * 16, False, xlwt.Alignment.HORZ_RIGHT))
        sheet.write_merge(1, 1, 0, 5, u'中国大唐集团公司培训项目综合评价表', self.set_style('Times New Roman', 22 * 20, True))
        title = self.set_style(name=fsgb, height=16 * 20)
        sheet.write(2, 0, u'年度：', title)
        sheet.write(2, 1, u'2019年', title)
        sheet.write(2, 4, u'岗位：', title)
        sheet.write(2, 5, obj['job'], title)

        style = self.set_style(name=fsgb, height=12 * 20, bold=True, border=True)
        no_bold_style = self.set_style(name=fsgb, height=12 * 20, border=True)
        comment_style = self.set_style(name=fsgb, height=9 * 20, border=False, horz=xlwt.Alignment.HORZ_LEFT)

        sheet.write(3, 0, u'姓名', style)
        sheet.write(3, 1, obj['userName'], no_bold_style)
        sheet.write(3, 2, u'性别', style)
        sheet.write(3, 3, obj['sex'], no_bold_style)
        sheet.write(3, 4, u'身份证号', style)
        sheet.write(3, 5, str(obj['idCard']), no_bold_style)

        sheet.write(4, 0, u'技术资格', style)
        sheet.write_merge(4, 4, 1, 2, obj['tech'], no_bold_style)
        sheet.write(4, 3, u'技能等级', style)
        sheet.write_merge(4, 4, 4, 5, obj['skillLevel'], no_bold_style)

        sheet.write(5, 0, u'二级企业名称', style)
        sheet.write_merge(5, 5, 1, 2, obj['company'], no_bold_style)
        sheet.write(5, 3, u'所在企业名称', style)
        sheet.write_merge(5, 5, 4, 5, obj['companyName'], no_bold_style)

        sheet.write(6, 0, u'所在部门', style)
        sheet.write_merge(6, 6, 1, 2, obj['departmentName'], no_bold_style)
        sheet.write(6, 3, u'现从事岗位', style)
        sheet.write_merge(6, 6, 4, 5, obj.get('job'), no_bold_style)

        sheet.write(7, 0, u'培训地点', style)
        sheet.write(7, 1, obj.get('position'), no_bold_style)
        sheet.write(7, 2, u'培训时间', style)
        sheet.write(7, 3, obj.get('time'), no_bold_style)
        sheet.write(7, 4, u'总学时', style)
        sheet.write(7, 5, obj.get('study_time'), no_bold_style)

        sheet.write(8, 0, u'培训考核成绩', style)
        sheet.write_merge(8, 8, 1, 2, obj.get('score'), no_bold_style)
        sheet.write(8, 3, u'合格分数线', style)
        sheet.write_merge(8, 8, 4, 5, u'60分', no_bold_style)

        sheet.write_merge(9, 11, 0, 0, u'岗位\n（工种）\n建议', style)

        text = obj.get('advice') if (
            obj.get('advice')) else u'\n    根据    规定，该同志在集团公司2019年    培训项目中达到考核标准，企业可依据实际情况，在上岗时予以参考。'
        sheet.write_merge(9, 9, 1, 5, text,
                          self.set_style(name=fsgb, height=12 * 20, border=True, horz=xlwt.Alignment.HORZ_LEFT,
                                         part_border={'top': True, 'left': True, 'right': True}))

        sheet.write_merge(10, 10, 1, 5, u'                    中国大唐集团公司培训专用章',
                          self.set_style(name=fsgb, height=14 * 20, border=True, part_border={'right': True}))
        sheet.write_merge(11, 11, 1, 5, u'                      ' + obj.get('advice_time'),
                          self.set_style(name=fsgb, height=12 * 20, border=True,
                                         part_border={'bottom': True, 'right': True}))

        sheet.write_merge(12, 12, 0, 5, u'    注:1.综合评价表是参加集团公司培训人员的成绩、水平等的证明材料。', comment_style)
        sheet.write_merge(13, 13, 0, 5, u'        2.岗位（工种）建议是指集团公司依据国家有关部委、行业颁布的职业标准或集团公司岗位标准，',
                          comment_style)
        sheet.write_merge(14, 14, 0, 5, u'        通过考试（考核）、评定等方式对参培人员岗位应具备能力、知识等进行评价，给出的岗位建议。',
                          comment_style)

    def execute(self, file_path=None, source='source'):

        if file_path:
            workbook = xlrd.open_workbook(file_path)
            path = './server/excel/' + str(round(time.time() * 1000)) + '.xls'
        else:
            file_url = '../' + source + '.xls'
            workbook = xlrd.open_workbook(file_url)
            path = '../target.xls'
        if os.path.exists(path):
            os.unlink(path)

        score = workbook.sheet_by_index(1)
        score_total = score.nrows
        score_map = {}
        for row in range(1, score_total):
            val = score.row_values(row)
            date = xlrd.xldate_as_tuple(val[7], 0)
            score_map[val[1]] = {
                "userName": val[0],
                "idCard": val[1],
                "position": val[2],
                "time": val[3],
                "study_time": val[4],
                "score": val[5],
                "advice": val[6],
                "advice_time": str(date[0]) + '年' + str(date[1]) + '月' + str(date[2]) + '日'
            }

        sh = workbook.sheet_by_index(0)
        total = sh.nrows
        excel_list = []
        for row in range(3, total - 2):
            columns = sh.row_values(row)
            if columns[8]:
                single = score_map[columns[8]]
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
                row_object.update(single)
                self.write(row_object)
                excel_list.append(row_object)

        self.wb.save(path)
        return path


if __name__ == '__main__':
    excel = ExcelExe()
    excel.execute()
    print('success!')
