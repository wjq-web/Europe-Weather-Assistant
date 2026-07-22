import json
from pathlib import Path
from urllib.parse import quote
from urllib.request import Request, urlopen

from openpyxl import Workbook
from openpyxl.styles import Font


CITIES = [
    "Vienna",
    "Prague",
    "Bratislava",
    "Budapest",
    "Munich",
]

excel_file = Path(__file__).parent / "Europe_Weather.xlsx"


def get_advice(max_temp, rain_chance):
    if rain_chance >= 60:
        return "建議帶傘"
    if max_temp >= 30:
        return "注意防曬和補水"
    if max_temp <= 12:
        return "注意保暖"
    if rain_chance <= 20:
        return "適合外出"
    return "可外出，備傘"


rows = []

for city in CITIES:
    url = f"https://wttr.in/{quote(city)}?format=j1"

    request = Request(
        url,
        headers={"User-Agent": "Mozilla/5.0"}
    )

    with urlopen(request, timeout=30) as response:
        data = json.load(response)

    for day in data["weather"]:
        min_temp = int(day["mintempC"])
        max_temp = int(day["maxtempC"])

        rain_chances = [
            int(hour.get("chanceofrain", 0))
            for hour in day["hourly"]
        ]
        max_rain_chance = max(rain_chances)

        midday_hour = day["hourly"][4]
        description = midday_hour["weatherDesc"][0]["value"]

        rows.append({
            "city": city,
            "date": day["date"],
            "min": min_temp,
            "max": max_temp,
            "description": description,
            "rain": max_rain_chance,
            "advice": get_advice(max_temp, max_rain_chance),
        })


workbook = Workbook()
sheet = workbook.active
sheet.title = "Europe Forecast"

headers = [
    "城市",
    "日期",
    "最低溫 °C",
    "最高溫 °C",
    "天氣概況",
    "最高降雨機率",
    "建議",
]

sheet.append(headers)

for cell in sheet[1]:
    cell.font = Font(bold=True)

for row in rows:
    sheet.append([
        row["city"],
        row["date"],
        row["min"],
        row["max"],
        row["description"],
        f'{row["rain"]}%',
        row["advice"],
    ])

sheet.column_dimensions["A"].width = 16
sheet.column_dimensions["B"].width = 14
sheet.column_dimensions["C"].width = 12
sheet.column_dimensions["D"].width = 12
sheet.column_dimensions["E"].width = 28
sheet.column_dimensions["F"].width = 16
sheet.column_dimensions["G"].width = 22
sheet.freeze_panes = "A2"

workbook.save(excel_file)

print("\n成功！")
print(f"共儲存 {len(rows)} 筆歐洲天氣資料。")
print("Excel 檔案位置：")
print(excel_file)

