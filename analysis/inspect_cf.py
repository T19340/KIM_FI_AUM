# -*- coding: utf-8 -*-
import re
import zipfile

p = r'C:\Users\USER\Cowork\2026-07-23_수탁고-자동화\template\FI운용본부수탁고_양식.xlsx'
z = zipfile.ZipFile(p)
xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
cf = re.findall(r'<conditionalFormatting[^>]*sqref="([^"]+)"', xml)
print('conditionalFormatting 범위:', cf)
print('dataBar 수:', xml.count('<dataBar'), '| x14:dataBar:', xml.count('x14:dataBar'))
print('negativeFillColor:', set(re.findall(r'negativeFillColor[^/]*rgb="([^"]+)"', xml)))
print('axisPosition:', set(re.findall(r'axisPosition="([^"]+)"', xml)))
m = re.search(r'<extLst>.*?</extLst>', xml, re.S)
if m:
    s = m.group(0)
    print('extLst 길이:', len(s))
    sq = re.findall(r'<xm:sqref>([^<]+)</xm:sqref>', s)
    print('x14 CF 적용 범위:', sq)
