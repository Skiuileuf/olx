from datetime import datetime, timezone
import json
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.utils import datetime as openpyxl_datetime
from openpyxl.styles import NamedStyle, numbers
import glob
from typing import Callable, Any, Optional

class Column:
    def __init__(self, title: str, value_func: Callable[[dict], Any], format_func: Optional[Callable] = None):
        self.title = title
        self.value_func = value_func
        self.format_func = format_func

def iso_to_excel_date(iso_string: str) -> float:
    dt = datetime.fromisoformat(iso_string)
    dt_utc = dt.astimezone(timezone.utc)
    dt_naive = dt_utc.replace(tzinfo=None)
    return openpyxl_datetime.to_excel(dt_naive)

def get_param_value(params, key: str) -> Optional[dict]:
    return next((param['value'] for param in params if param['key'] == key), None)

target_lat = 44.55048
target_lon = 26.09095

def hyperlink(cell: Cell):
    cell.hyperlink = cell.value
    cell.style = "Hyperlink"

# Define column definitions
columns = [
    Column(
        "id", 
        lambda item: item["id"]
    ),
    Column(
        "url", 
        lambda item: item["url"], 
        lambda cell: hyperlink(cell)
    ),
    Column(
        "title", 
        lambda item: item["title"]
    ),
    Column(
        "last_refresh_time", 
        lambda item: iso_to_excel_date(item["last_refresh_time"]), 
        lambda cell: setattr(cell, "style", "datetime")
    ),
    Column(
        "created_time", 
        lambda item: iso_to_excel_date(item["created_time"]), 
        lambda cell: setattr(cell, "style", "datetime")
    ),
    Column(
        "description", 
        lambda item: item["description"]
    ),
    Column(
        "state", 
        lambda item: get_param_value(item["params"], "state")["label"] if get_param_value(item["params"], "state") else None
    ),
    Column(
        "diagonala", 
        lambda item: get_param_value(item["params"], "diagonala")["label"] if get_param_value(item["params"], "diagonala") else None
    ),
    Column(
        "producator_procesor", 
        lambda item: get_param_value(item["params"], "producator_procesor")["label"] if get_param_value(item["params"], "producator_procesor") else None
    ),
    Column(
        "capacitate_memorie_ram", 
        lambda item: get_param_value(item["params"], "capacitate_memorie_ram")["label"] if get_param_value(item["params"], "capacitate_memorie_ram") else None
    ),
    Column(
        "price", 
        lambda item: get_param_value(item["params"], "price")["value"] if get_param_value(item["params"], "price") else None
    ),
    Column(
        "currency", 
        lambda item: get_param_value(item["params"], "price")["currency"] if get_param_value(item["params"], "price") else None
    ),
    Column(
        "negotiable", 
        lambda item: get_param_value(item["params"], "price")["negotiable"] if get_param_value(item["params"], "price") else None
    ),
    Column(
        "lat", 
        lambda item: item["map"]["lat"]
    ),
    Column(
        "lon", 
        lambda item: item["map"]["lon"]
    ),
    Column(
        "distance", lambda item: ((item["map"]["lat"] - target_lat)**2 + (item["map"]["lon"] - target_lon)**2)**0.5
    ),
    Column(
        "city", 
        lambda item: item["location"]["city"]["name"]
    ),
    Column(
        "region", 
        lambda item: item["location"]["region"]["name"]
    ),
    Column(
        "days_since_last_refresh", 
        lambda item: f"=DATEDIF(D{{row}}, TODAY(), \"D\")"
    ),
    Column(
        "days_since_created_time", 
        lambda item: f"=DATEDIF(E{{row}}, TODAY(), \"D\")"
    ),
    Column(
        "highlighted", 
        lambda item: item["promotion"]["highlighted"]
    ),
]

def main():
    wb = Workbook()
    ws = wb.active

    date_style = NamedStyle(name='datetime', number_format='YYYY-MM-DD HH:MM:SS')
    wb.add_named_style(date_style)

    # Write headers
    ws.append([col.title for col in columns])

    current_row = 2

    files = glob.glob("api/*.json")
    for file_path in files:
        print(file_path)
        with open(file_path, "rb") as f:
            data = json.load(f)

        for item in data["data"]:
            row = []
            for col in columns:
                value = col.value_func(item)
                if isinstance(value, str) and value.startswith('='):
                    value = value.format(row=current_row)
                row.append(value)

            ws.append(row)

            for col_index, col in enumerate(columns, start=1):
                if col.format_func:
                    col.format_func(ws.cell(row=current_row, column=col_index))

            current_row += 1

    wb.save("olx-data.xlsx")

if __name__ == "__main__":
    main()