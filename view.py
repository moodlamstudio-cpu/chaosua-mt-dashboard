import json
d=json.load(open("data.json",encoding="utf-8"))
print("closed_month:",d["closed_month"])
for m in d["months"]:
    print(f"M{m['m']}: act={m['total']['actual']} le={m['total']['le']} ly={m['total']['ly']} aop={m['total']['aop']}")
print("YTD act:",d["ytd"]["total_actual"],"le:",d["ytd"]["total_le"],"ly:",d["ytd"]["total_ly"],"aop:",d["ytd"]["total_aop"])
