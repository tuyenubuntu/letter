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

src = 'sample.jpg'
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
        img.save("pic1\\result{}.jpg".format(i))
draw(src)
