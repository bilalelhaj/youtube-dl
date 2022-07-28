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

pathfile = ""


def checkLang(lang):
    global pathfile
    if lang == "a":
        pathfile = 'arab'
    elif lang == "e":
        pathfile = 'eng'
    elif lang == "g":
        pathfile = 'ger'
    else:
        print('Please use just "a","e" or "g"!')
        language = input("Arab, English or German? (a,e,g): ")
        checkLang(language)


def createImg(a):
    bidi = a
    reshaped_text = arabic_reshaper.reshape(bidi)
    thumbnail = get_display(reshaped_text)
    para = textwrap.wrap(thumbnail, width=22)
    img = Image.open(f"./{pathfile}/thumbnail.jpg")
    W, H = img.size
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype("./arab/amiri.ttf", 150)
    draw.textsize(thumbnail, font=font)
    if language == 'a':
        current_h, pad = 550, 10
    else:
        current_h, pad = 250, 10
    for line in para:
        w, h = draw.textsize(line, font=font)
        draw.text(((W - w) / 2, current_h), line, font=font)
        if language == 'a':
            current_h -= h + pad
        else:
            current_h += h + pad
    img.save(f'{pathfile}/current.jpg')


def searchNewFileForAudio():
    list_of_files = glob.glob(f'{pathfile}/*.mp3')
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file


def searchNewFileForVideo():
    list_of_files = glob.glob(f'{pathfile}/*.mp4')
    latest_file = max(list_of_files, key=os.path.getctime)
    print(latest_file)


def find(name, path):
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root, name)


def startConvert(n):
    new = n.split(":")
    new[2] = int(new[2])
    if new[2] >= 3:
        new[2] -= 3
    return f"{new[0]}:{new[1]}:{new[2]}"


def difference(n, m):
    new = str(n).split(":")
    mew = str(m).split(":")

    new[0] = int(new[0])
    new[1] = int(new[1])
    new[2] = int(new[2])

    mew[0] = int(mew[0])
    mew[1] = int(mew[1])
    mew[2] = int(mew[2])

    sec = mew[0] - new[0]
    minu = abs(mew[1] - new[1])
    hours = abs(mew[2] - new[2])
    return f"{sec}:{minu}:{hours}"


def get_length(filename):
    result = subprocess.run(["ffprobe", "-v", "error", "-show_entries",
                             "format=duration", "-of",
                             "default=noprint_wrappers=1:nokey=1", filename],
                            stdout=subprocess.PIPE,
                            stderr=subprocess.STDOUT)
    return float(result.stdout)


def upload(pUp, find, cred_num):
    upload_T = pUp
    commandOn = f"youtubeuploader -title '{upload_T}' -description '' -cache './{pathfile}/client/{cred_num}.token' -secrets './{pathfile}/client/{cred_num}.json' -thumbnail './{pathfile}/current.jpg' -filename './{pathfile}/{find}.mp4'"
    command = os.system(commandOn)
    command_line = 'Error making YouTube API call: googleapi: Error 403: The request cannot be completed because you have exceeded your <a href="/youtube/v3/getting-started#quota">quota</a>., quotaExceeded'
    if command == 0:
        print("\f")
        print("Video was succesfully uploaded")
    else:
        try:
            print("\f")
            if cred_num <= 12:
                upload(upload_T, find, cred_num + 1)
            else:
                print("Need more credentials")
        except:
            print("ERROR404")


def main(a,b,c,d,e,f,g,h):
    url = a
    title = b
    title_x = title + "x"
    output = title_x + "x"
    output2 = output + "x"
    upload_T = c
    createImg(d)
    fVideo = e
    if fVideo == "part":
        start = f
        question2 = g
        if question2 == "y": 
            result = os.system(
                f"ffmpeg $(youtube-dl -f 22 -g -o './{pathfile}/{title}' --youtube-skip-dash-manifest '{url}' | sed 's/.*/-ss {start} -i &/') -c copy ./{pathfile}/{title}.mp4")
            if result == 0:
                print("\f")
                print("Video was successfully downloaded!")
            else:
                print("\f")
                print("Video couldn't download!")
        elif question2 == "n":
            end = h
            if end <= start:
                print("\f")
                print("End can't be smaller than or equal to start!")
                sys.exit()

            result = os.system(
                f"ffmpeg $(youtube-dl -f 22 -g -o './{pathfile}/{title}' --youtube-skip-dash-manifest '{url}' | sed 's/.*/-ss {start} -i &/') -t {difference(start, end)} -c copy ./{pathfile}/{title}.mp4")
            if result == 0:
                print("\f")
                print("Video was successfully downloaded!")
            else:
                print("\f")
                print("Video couldn't download!")
        else:
            print("\f")
            print("Just use y/n please!")

    elif fVideo == "full":
        result = os.system(
            f"youtube-dl -f 22 --no-playlist -o './{pathfile}/{title}' '{url}'")
        if result == 0:
            print("\f")
            print("Video was successfully downloaded!")
        else:
            print("\f")
            print("Video couldn't download")
    else:
        print("\f")
        print("Just use full/part please!")

    command2 = f"""ffmpeg -i ./{pathfile}/{title}.mp4 -filter_complex \
    "[0:v]setpts=PTS-STARTPTS[v0];
    movie="sub-animation.mov":s=dv+da[overv][overa];
    [overv]setpts=PTS-STARTPTS+20/TB[v1];
    [v0][v1]overlay=-600:0:eof_action=pass[out1];
    [overa]adelay=20000|20000,volume=0.5[a1];
    [0:a][a1]amix=inputs=2:duration=longest:dropout_transition=0:weights=2 1[outa]" \
    -map "[out1]" -map "[outa]" ./{pathfile}/{title_x}.mp4"""
    os.system(command2)
    os.remove(f"./{pathfile}/{title}.mp4")
    os.system(
        f"ffmpeg -i ./{pathfile}/{title_x}.mp4 -vf 'fade=t=in:st=0:d=5' ./{pathfile}/{output}.mp4")
    os.remove(f"./{pathfile}/{title_x}.mp4")
    length = get_length(f"./{pathfile}/{output}.mp4")
    length -= 5
    os.system(
        f"ffmpeg -i ./{pathfile}/{output}.mp4 -vf 'fade=t=out:st={length}:d=5' ./{pathfile}/{output2}.mp4")
    os.remove(f"./{pathfile}/{output}.mp4")
    upload(upload_T, output2, 1)
    os.remove(f"./{pathfile}/{output2}.mp4")


for x in range(len(language)):
    checkLang(language[x])
    main(url[x],title_to_save[x],upload_title[x],thumbnail[x],full_part[x],start[x],complete_to_end[x],end[x])
