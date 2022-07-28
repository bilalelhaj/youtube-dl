import os
import pathlib
import fnmatch
import time
import glob
from PIL import Image, ImageDraw, ImageFont
import textwrap
import arabic_reshaper
from bidi.algorithm import get_display
import sys
import subprocess
import openpyxl
import pandas as pd

file = 'main.xlsx'
data = pd.ExcelFile(file)
df = data.parse('Tabelle1')
df.info
ps = openpyxl.load_workbook(file)
sheet = ps['Tabelle1']

lang1 = []
url = []
title_to_save = []
upload_title = []
thumbnail = []
full_part = []
start = []
complete_to_end = []
end = []

for row in range(3, sheet.max_row + 1):
    lang1.append(sheet['A' + str(row)].value)
    url.append(sheet['B' + str(row)].value)
    title_to_save.append(sheet['C' + str(row)].value)
    upload_title.append(sheet['D' + str(row)].value)
    thumbnail.append(sheet['E' + str(row)].value)
    full_part.append(sheet['F' + str(row)].value)
    start.append(sheet['G' + str(row)].value)
    complete_to_end.append(sheet['H' + str(row)].value)
    end.append(sheet['I' + str(row)].value)

language = lang1


for x in range(len(language)):
    print(language[x])
    print(url[x])
    print(title_to_save[x])
    print(upload_title[x])
    print(thumbnail[x])
    print(full_part[x])
    print(start[x])
    print(complete_to_end[x])
    print(end[x])
