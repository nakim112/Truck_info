# Truck Info

INC4 HUB vehicle information desktop app.

## Run

```powershell
.\.venv\Scripts\python.exe truck_info_ui_airport.py
```

## Build

```powershell
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed --name truck_info_ui_airport --icon Delivery.ico `
  --collect-all selenium `
  --collect-all webdriver_manager `
  --hidden-import selenium.webdriver.chrome.webdriver `
  --hidden-import selenium.webdriver.chrome.service `
  --hidden-import selenium.webdriver.chrome.options `
  --add-data "Delivery.ico;." `
  --add-data "coupang_logo.png;." `
  --add-data "coupang_rocket.png;." `
  --add-data "camp_settings.json;." `
  --add-data "마스터_양식설정.xlsx;." `
  --add-data "고임목.xlsx;." `
  truck_info_ui_airport.py
```

`출차서.xlsx` and `chromedriver.exe` are local runtime/build files and are not tracked in Git.
