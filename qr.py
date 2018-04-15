#!/usr/bin/python
# -*- coding: utf-8 -*-
import qrcode
import os
import requests
import arabic_reshaper
import persian

from StringIO import StringIO
from openpyxl import load_workbook
from PIL import Image, ImageOps, ImageFont, ImageDraw

from bidi.algorithm import get_display

wb = load_workbook('attendees_list.xlsx')
ws = wb.get_sheet_by_name('report')
attendees = str(ws.max_row)
person_name = ws['F2':'F'+ attendees]

for row in ws.iter_rows('C{}:C{}'.format(ws.min_row+1,ws.max_row)): #ws.max_row
	for cell in row:
		# qr = qrcode.make()
		qr = qrcode.QRCode(
		    version=1,
		    error_correction=qrcode.constants.ERROR_CORRECT_L,
		    box_size=12,
		    border=4,
		)
		qr.add_data(cell.value)
		qr.make(fit=True)
		qrimage = qr.make_image(fill_color="black", back_color="white")
		back = Image.open("./PNG_badge/01 (2).png")
		back.paste(qrimage, (275, 360))
		back.save('./export/'+ws.cell(row=cell.row, column=16).value.capitalize()+'_'+ws.cell(row=cell.row, column=19).value.capitalize()+'(2).png')

		front = Image.open("./PNG_badge/01 (1).png")

		draw = ImageDraw.Draw(front)
		# font = ImageFont.truetype(<font-file>, <font-size>)
		font2 = ImageFont.truetype("./assets/font/vazir-font-v6.3.4/Vazir-Bold.ttf", 46, encoding="unic")
		font1 = ImageFont.truetype("./assets/font/lato/Lato-Regular.ttf", 86, encoding="unic")
		# font = ImageFont.truetype("./Lalezar-Regular.ttf", 26, encoding="unic")
		text1 = ws.cell(row=cell.row, column=16).value.capitalize()+' '+ws.cell(row=cell.row, column=19).value.capitalize()
		text2 = ws.cell(row=cell.row, column=20).value
		# unicode_text = unicode(text, "utf-8")
		# text = get_display(text)
		reshaped_text = arabic_reshaper.reshape(text1)
		final_text = get_display(reshaped_text)
		print final_text
		# draw.text((x, y),"Sample Text",(r,g,b))
		w, h = draw.textsize(final_text, font=font1)
		if w > 900:
			howbig = ((int(w)-900)*100)/900
			font_size = 86 - (86*howbig/100) - 10
		else:
			font_size = 86
		font1 = ImageFont.truetype("./assets/font/lato/Lato-Regular.ttf", font_size, encoding="unic")
		draw.text(((945-w)/2, 600),final_text,(255,255,255),font=font1)
		reshaped_text = arabic_reshaper.reshape(text2)
		final_text = get_display(reshaped_text)
		w, h = draw.textsize(final_text, font=font2)
		if w > 460:
			howbig = ((int(w)-460)*100)/460
			font_size = 46 - (46*howbig/100) - 2
		else:
			font_size = 46
		font2 = ImageFont.truetype("./assets/font/vazir-font-v6.3.4/Vazir-Bold.ttf", font_size, encoding="unic")
		w, h = draw.textsize(final_text, font=font2)
		draw.text(((945-w)/2, 800),final_text,(255,255,255),font=font2)
		r = requests.get(ws.cell(row=cell.row, column=21).value)
		if r.status_code == 200:
			headshot = Image.open(StringIO(r.content))
			headshot = headshot.resize((353, 353));
			bigsize = (headshot.size[0] * 3, headshot.size[1] * 3)
			mask = Image.new('L', bigsize, 0)
			draw = ImageDraw.Draw(mask) 
			draw.ellipse((0, 0) + bigsize, fill=255)
			mask = mask.resize(headshot.size, Image.ANTIALIAS)
			headshot.putalpha(mask)

			output = ImageOps.fit(headshot, mask.size, centering=(0.5, 0.5))
			output.putalpha(mask)

			front.paste(headshot, (297, 204), headshot)
			front.save('./export/'+ws.cell(row=cell.row, column=16).value.capitalize()+'_'+ws.cell(row=cell.row, column=19).value.capitalize()+'(1).png')

