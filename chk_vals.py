import json
d=json.load(open("data.json",encoding="utf-8"))
print("closed_month:",d["closed_month"],"| current:",d["current_month"])
print("YTD total_actual:",d["ytd"]["total_actual"])
print("YTD total_le:",d["ytd"]["total_le"])
act=sum(m["total"]["actual"] for m in d["months"])
print("sum actual all months:",round(act,1))
# actual ถึง ม.8
a8=sum(m["total"]["actual"] for m in d["months"] if m["m"]<=8)
print("actual M1-8:",round(a8,1))
