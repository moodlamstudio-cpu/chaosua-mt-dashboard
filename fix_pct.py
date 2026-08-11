import json
d=json.load(open("data_channels.json",encoding="utf-8"))
# fix pct: เก็บ base ใน baht แล้ว pct = mb/total_mb*100
for ch in d:
    if ch=="_channels": continue
    c=d[ch]
    tot=c["total_mb"]
    for it in c["category"]["items"]:
        it["pct"]=round(it["mb"]/tot*100,1) if tot else 0
    sku=c.get("top_sku_all",[])
    for s in sku: s["pct_cat"]=""   # clear
# เปลี่ยน label LOTUS' -> Lotus (ชื่อสะอาด สำหรับ target) แต่เก็บ id ไว้
json.dump(d,open("data_channels.json","w",encoding="utf-8"),ensure_ascii=False,indent=2)
print("MT category pct แก้แล้ว:")
for it in d["MT"]["category"]["items"]: print(" ",it["name"],it["mb"],str(it["pct"])+"%")
