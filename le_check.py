import json
d=json.load(open("data.json",encoding="utf-8"))
print("=== ตรวจ LE แบบที่ควรโชว์ (Actual ม.1-7 + LE ม.8) ===")
for ch in ["total","makro","lotus"]:
    a7=sum(m[ch]["actual"] for m in d["months"] if m["m"]<=7)
    le8=d["months"][7][ch]["le"]  # ม.8
    ly=d["months"][7][ch]["ly"]   # ม.8 LY
    aop=d["months"][7][ch]["aop"]
    print(f"{ch}: Actual ม.1-7={a7:.1f} + LE ม.8={le8:.1f} = {a7+le8:.1f} | LY ม.8={ly:.1f} | AOP ม.8={aop:.1f}")
print()
print("=== YTD ในไฟล์ (ปัจจุบัน ใช้อยู่) ===")
for ch in ["total","makro","lotus"]:
    print(f"{ch}: actual={d['ytd'][ch+'_actual']} le={d['ytd'][ch+'_le']} ly={d['ytd'][ch+'_ly']} aop={d['ytd'][ch+'_aop']}")
