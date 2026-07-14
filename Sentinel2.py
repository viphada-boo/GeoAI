# -*- coding: utf-8 -*-
"""Sentinel-2 median composite ปลอดเมฆ ปี 2026 (เท่าที่มีข้อมูลถึงปัจจุบัน)
clip เป็นรูปประเทศไทยตามขอบเขต Thailand.shp แสดงบนพื้นหลัง Google Satellite
แล้ว export เป็น Sentinel2Thailand.html
"""
import importlib
import json

import ee
import geemap as _geemap_pkg
import geopandas as gpd

# workaround บั๊ก geemap 0.38.3: __init__.py สร้างตัวแปร basemaps (Box) ทับ submodule
# ทำให้ import geemap.foliumap ล้มเหลว — คืนค่า submodule กลับก่อน import
_geemap_pkg.basemaps = importlib.import_module("geemap.basemaps")
import geemap.foliumap as geemap  # noqa: E402
import folium  # noqa: E402

PROJECT_ID = "ee-viphadaboo"
BOUNDARY_SHP = "Example-20260714T031623Z-1-001/Example/Thailand.shp"
OUT_HTML = "Sentinel2Thailand2.html"
DATE_START, DATE_END = "2026-01-01", "2027-01-01"  # ปี 2026 เท่าที่มีข้อมูล
CS_THRESHOLD = 0.6  # Cloud Score+ : เก็บเฉพาะ pixel ที่มั่นใจว่าปลอดเมฆ

# ---------- เชื่อมต่อ Google Earth Engine ----------
try:
    ee.Initialize(project=PROJECT_ID)
except Exception:
    # ครั้งแรกจะเปิดเบราว์เซอร์ให้ล็อกอิน Google เพื่อขอสิทธิ์ Earth Engine
    ee.Authenticate()
    ee.Initialize(project=PROJECT_ID)


# ---------- ขอบเขตประเทศไทยจาก Thailand.shp ----------
gdf = gpd.read_file(BOUNDARY_SHP).to_crs(epsg=4326)
# simplify ~500 ม. เพื่อให้ geometry เล็กพอส่งเข้า GEE (จาก 12 MB เหลือ ~230 KB)
geom_simplified = gdf.geometry.simplify(0.005, preserve_topology=True).iloc[0]
thailand = ee.Geometry(geom_simplified.__geo_interface__)

# ---------- Sentinel-2 SR + Cloud Score+ mask ----------
s2 = (
    ee.ImageCollection("COPERNICUS/S2_SR_HARMONIZED")
    .filterBounds(thailand)
    .filterDate(DATE_START, DATE_END)
    .filter(ee.Filter.lt("CLOUDY_PIXEL_PERCENTAGE", 60))
)
cs_plus = (
    ee.ImageCollection("GOOGLE/CLOUD_SCORE_PLUS/V1/S2_HARMONIZED")
    .filterBounds(thailand)
    .filterDate(DATE_START, DATE_END)
)


def mask_clouds(image):
    """เก็บเฉพาะ pixel ที่ Cloud Score+ (cs_cdf) มั่นใจว่าปลอดเมฆ แล้ว scale เป็น reflectance"""
    cs = image.select("cs_cdf")
    return image.updateMask(cs.gte(CS_THRESHOLD)).divide(10000)


composite = (
    s2.linkCollection(cs_plus, ["cs_cdf"])
    .map(mask_clouds)
    .median()
    .clip(thailand)  # ตัดภาพเป็นรูปประเทศไทย นอกขอบเขตโปร่งใสเห็นพื้นหลัง
)

# ---------- NDVI / NDWI จากภาพเดือนที่ผ่านมา (มิ.ย. 2026) ----------
MONTH_START, MONTH_END = "2026-06-01", "2026-07-01"
monthly = (
    s2.filterDate(MONTH_START, MONTH_END)
    .linkCollection(cs_plus.filterDate(MONTH_START, MONTH_END), ["cs_cdf"])
    .map(mask_clouds)
    .median()
    .clip(thailand)
)
ndvi = monthly.normalizedDifference(["B8", "B4"]).rename("NDVI")   # พืชพรรณ (NIR-Red)
ndwi = monthly.normalizedDifference(["B3", "B8"]).rename("NDWI")   # แหล่งน้ำ McFeeters (Green-NIR)

rgb_vis = {"min": 0.0, "max": 0.3, "bands": ["B4", "B3", "B2"]}          # True color
false_vis = {"min": 0.0, "max": 0.4, "bands": ["B8", "B4", "B3"]}        # False color (พืชพรรณ)
ndvi_vis = {
    "min": -0.2, "max": 0.8,
    "palette": ["#d73027", "#f46d43", "#fee08b", "#d9ef8b", "#66bd63", "#1a9850"],
}
ndwi_vis = {
    "min": -0.5, "max": 0.5,
    "palette": ["#8c510a", "#d8b365", "#f5f5f5", "#67a9cf", "#0571b0"],
}

# ---------- สร้างแผนที่บนพื้นหลัง Google Satellite ----------
b = gdf.total_bounds  # [minx, miny, maxx, maxy]
m = geemap.Map(center=[(b[1] + b[3]) / 2, (b[0] + b[2]) / 2], zoom=6, tiles=None)

folium.TileLayer(
    tiles="https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}",
    attr="© Google",
    name="Google Satellite",
    max_zoom=20,
).add_to(m)

m.add_layer(composite, rgb_vis, "Sentinel-2 Median 2026 (ปลอดเมฆ)")
m.add_layer(composite, false_vis, "Sentinel-2 False Color (NIR)", shown=False)
m.add_layer(ndvi, ndvi_vis, "NDVI มิ.ย. 2026 (พืชพรรณ)", shown=False)
m.add_layer(ndwi, ndwi_vis, "NDWI มิ.ย. 2026 (แหล่งน้ำ)", shown=False)

folium.GeoJson(
    json.loads(gpd.GeoDataFrame(geometry=[geom_simplified], crs=4326).to_json()),
    name="ขอบเขตประเทศไทย",
    style_function=lambda f: {"fillOpacity": 0, "color": "#FFD60A", "weight": 1.5},
).add_to(m)

m.add_layer_control()
m.fit_bounds([[b[1], b[0]], [b[3], b[2]]])

m.to_html(filename=OUT_HTML, title="Sentinel-2 Thailand 2026", width="100%", height="880px")
print(f"saved: {OUT_HTML}")
