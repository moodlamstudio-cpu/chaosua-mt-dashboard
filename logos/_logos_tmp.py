# -*- coding: utf-8 -*-
import io, os
from PIL import Image, ImageDraw, ImageFont

out = r"C:\Users\teera\OneDrive - TIA NGEE HIANG (CHAOSUA) CO.,LTD\Desktop\Chaosua_Ice\Web\LotusMakro-Dashboard\logos"

# ชื่อไฟล์ต้องตรงกับ id ใน data_channels.json + .png
brands = [
    ("7Eleven","7-ELEVEN",(0xE2,0x11,0x37)),
    ("MAKRO","MAKRO",(0x00,0x5E,0xAB)),
    ("LOTUS","LOTUS'S",(0x00,0x46,0x91)),
    ("BigC","Big C",(0x00,0x8F,0xD0)),
    ("Tops","TOPS",(0x00,0x8F,0x39)),
    ("CJExpress","CJ Express",(0xE8,0x15,0x2D)),
    ("Jiffy","JIFFY",(0x00,0x9B,0x4F)),
    ("Foodland","FOODLAND",(0xC8,0x0F,0x2E)),
    ("FamilyMart","FamilyMart",(0x00,0x9E,0x74)),
    ("GoldenPlace","GP",(0xE6,0xA8,0x17)),
    ("Centralfood","Central Food",(0x00,0x47,0x59)),
    ("Maxmart","MAXMART",(0x1A,0x2376,0)),
    ("MaxValue","MaxValue",(0xE1,0x3B,0x3B)),
    ("Villa","VILLA",(0x2E,0x7D,0x32)),
    ("Lawson","LAWSON",(0x0B,0x6F,0xDD)),
]

# ลองฟอนต์ Segoe UI Bold (มีบน Windows)
def font(sz):
    for p in [r"C:\Windows\Fonts\segoeuib.ttf", r"C:\Windows\Fonts\arialbd.ttf"]:
        if os.path.exists(p):
            try: return ImageFont.truetype(p, sz)
            except: pass
    return ImageFont.load_default()

W,H=180,56
for fname,label,color in brands:
    img=Image.new("RGBA",(W,H),(0,0,0,0))
    d=ImageDraw.Draw(img)
    # พื้นหลังกล่องเทาอ่อน
    d.rounded_rectangle([0,0,W-1,H-1],radius=10,fill=(245,247,250,255),outline=(226,232,240,255),width=2)
    # แถบสีแบรนด์ด้านซ้าย
    d.rounded_rectangle([8,9,14,H-10],radius=3,fill=color)
    # ข้อความ
    f=font(15 if len(label)>8 else 17)
    bbox=d.textbbox((0,0),label,font=f)
    tw=bbox[2]-bbox[0]
    d.text(((W+tw)/2, H/2), label, font=f, fill=(51,65,85,255), anchor="mm")
    img.save(os.path.join(out,fname+".png"))
    print(fname, label, "OK")
print("DONE", len(brands), "logos")
