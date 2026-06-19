import base64
from pathlib import Path
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[1]
IMG_DIR = ROOT / "images"
TEMPLATE = IMG_DIR / "dashboard_template.html"
OUT_HTML = IMG_DIR / "_dashboard_render.html"
OUT_PNG = IMG_DIR / "dashboard_preview.png"


def b64img(path):
    data = base64.b64encode(Path(path).read_bytes()).decode()
    return f"data:image/png;base64,{data}"


html = TEMPLATE.read_text()
html = html.replace("__TRAFFIC_CHART__", b64img(IMG_DIR / "eda_traffic_conversion_trend.png"))
html = html.replace("__DEVICE_ALERT_CHART__", b64img(IMG_DIR / "eda_device_alert_mix.png"))
html = html.replace("__REVENUE_DIVISION_CHART__", b64img(IMG_DIR / "eda_revenue_by_division.png"))
html = html.replace("__TICKET_CHART__", b64img(IMG_DIR / "eda_ticket_resolution.png"))
OUT_HTML.write_text(html)

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(viewport={"width": 1600, "height": 1200}, device_scale_factor=2)
    page.goto(f"file://{OUT_HTML}")
    page.wait_for_timeout(300)
    page.screenshot(path=str(OUT_PNG), full_page=True)
    browser.close()

print("Saved", OUT_PNG)
