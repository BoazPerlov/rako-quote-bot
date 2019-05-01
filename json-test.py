import json
from openpyxl import Workbook

with open("Rak_Prices.json") as f:
    data = json.load(f)

bridge = input("What kind of system is this?")

for product in data:
    if bridge == "wired":
        if product["Code"] == "RAK-WA-BRIDGE":
            product["Qty"] = 1
    elif bridge == "wireless":
        if product["Code"] == "RAK-RA-BRIDGE":
            product["Qty"] = 1

for product in data:
    qty = int(product["Qty"])
    if qty > 0:
        print(product["Code"] + " ordered!")