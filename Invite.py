#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import os
import pandas as pd
import openpyxl
from openpyxl import Workbook, load_workbook
from PIL import Image, ImageFont, ImageDraw,  ImageOps
f1 = 'mse3.xlsx'
df = pd.read_excel(f1, sheet_name=0)
msnv = list(df['MSNV'])
name = list(df['Họ & Tên'])
gp = list(df['Group'])


# In[ ]:


src = 'sample.png'


# In[ ]:


#function fill in formation
def draw (src):
    for i in range(len (name)):
        img = Image.open(src)
        title_font = ImageFont.truetype('Playfair_Display\PlayfairDisplay-Italic-VariableFont_wght.ttf',60)
        image_editable = ImageDraw.Draw(img)
        
        # Information ticket
        
        #fill in part1 ticket
        namex = '{}'.format(name[i])
        image_editable.text((800,170), namex, (0, 0, 0), font=title_font)
        msnvx = '{}'.format(msnv[i])
        image_editable.text((1600,600), msnvx, (0, 0, 0), font=title_font)
        #depx = gp[i]
        #image_editable.text((105,800), depx, (255, 255, 255), font=title_font)
        
        #fill in part2 ticket
        
        width, height = 700,100
        font =  ImageFont.truetype('Playfair_Display\PlayfairDisplay-Italic-VariableFont_wght.ttf',60)
        #fill info name rotate 90
        image1 = Image.new('RGBA', (width, height), (255, 255, 255, 1))
        draw1 = ImageDraw.Draw(image1)
        draw1.text((0, 0), text='{}'.format(name[i]), font=font, fill=(0, 0, 0))
        image1 = image1.rotate(270, expand=1)
        if 20 > len(name[i]) > 15:
            px1, py1 = 2250, 180
        elif len(name[i]) >= 20:
            px1, py1 = 2250, 100
        elif 10 < len(name[i]) <= 20:
            px1, py1 = 2250, 210
        elif 7 < len(name[i]) <= 10:
            px1, py1 = 2250, 280
        elif len(name[i]) <= 7:
            px1, py1 = 2250, 300
        sx1, sy1 = image1.size
        img.paste(image1, (px1, py1, px1 + sx1, py1 + sy1), image1) 
        #fill info code ticket 90    
        image2 = Image.new('RGBA', (width, height), (255, 255, 255, 1))
        draw2 = ImageDraw.Draw(image2)
        draw2.text((0, 0), text='{}'.format(msnv[i]), font=font, fill=(0, 0, 0))
        image2 = image2.rotate(270, expand=1)
        px2, py2 = 2140, 503
        sx2, sy2 = image2.size
        img.paste(image2, (px2, py2, px2 + sx2, py2 + sy2), image2)
         
        #note  ticket
        
        #note part1
        title_font_note = ImageFont.truetype('Playfair_Display\PlayfairDisplay-Italic-VariableFont_wght.ttf',30)
        note = 'Vui lòng giữ lại phần phiếu này để xác thực khi nhận thưởng.'
        image_editable.text((200,50), note, (0, 0, 0), font=title_font_note)
        #note part2
        image3 = Image.new('RGBA', (width, height), (255, 255, 255, 1))
        draw3 = ImageDraw.Draw(image3)
        draw3.text((0, 0), text='Phần phiếu này bỏ vào thùng.', font=title_font_note, fill=(0, 0, 0))
        image3 = image3.rotate(270, expand=1)
        px3, py3 = 2000, 250
        sx3, sy3 = image3.size
        img.paste(image3, (px3, py3, px3 + sx3, py3 + sy3), image3)
                        
        #img.show()
        img.save("pic\{}.png".format(i))


# In[ ]:

'''
#tempory function

def draw (src):
    for i in range(len (msnv)):

        img = Image.open(src)
        title_font = ImageFont.truetype('Playfair_Display\PlayfairDisplay-Italic-VariableFont_wght.ttf',60)
        image_editable = ImageDraw.Draw(img)
        
        namex = '{}'.format(name[i])
        image_editable.text((800,170), namex, (0, 0, 0), font=title_font)
        msnvx = '{}'.format(msnv[i])
        image_editable.text((1600,600), msnvx, (0, 0, 0), font=title_font)
        #depx = gp[i]
        #image_editable.text((105,800), depx, (255, 255, 255), font=title_font)
        
        
        width, height = 700,100
        font =  ImageFont.truetype('Playfair_Display\PlayfairDisplay-Italic-VariableFont_wght.ttf',60)
        image1 = Image.new('RGBA', (width, height), (255, 255, 255, 1))
        draw1 = ImageDraw.Draw(image1)
        draw1.text((0, 0), text='{}'.format(name[i]), font=font, fill=(0, 0, 0))
        image1 = image1.rotate(270, expand=1)
        px1, py1 = 2250, 100
        sx1, sy1 = image1.size
        img.paste(image1, (px1, py1, px1 + sx1, py1 + sy1), image1)
        
        
        image2 = Image.new('RGBA', (width, height), (255, 255, 255, 1))
        draw2 = ImageDraw.Draw(image2)
        draw2.text((0, 0), text='{}'.format(msnv[i]), font=font, fill=(0, 0, 0))
        image2 = image2.rotate(270, expand=1)
        px2, py2 = 2140, 503
        sx2, sy2 = image2.size
        img.paste(image2, (px2, py2, px2 + sx2, py2 + sy2), image2)
        
        
        #img.show()
        img.save("pic\{}.jpg".format(i))
'''

# In[ ]:


#run function
draw(src)




'''
# In[ ]:


#Scale image
import PIL
from PIL import Image
import os, sys
path = "pic"
dirs = (os.listdir(path))
print (dirs)


# In[ ]:


#Scale image
def resize():
    for item in dirs:
        img = Image.open('pic\{}'.format(item))
        img = img.resize((200,65), Image.ANTIALIAS)
        img.save('pic1\{}'.format(item)) 
resize()

'''