# -*- coding: utf-8 -*-
"""สร้างแผนที่ interactive Thailand.html จาก Thailand.geojson บนพื้นหลัง Google Maps"""
import folium
import geopandas as gpd

SRC = "Example-20260714T031623Z-1-001/Example/Thailand.geojson"
OUT = "Example-20260714T031623Z-1-001/Example/Thailand.html"

gdf = gpd.read_file(SRC)
# ลดความละเอียดเส้นขอบ ~50 ม. เพื่อให้ไฟล์ HTML เล็กและลื่น (มองระดับประเทศไม่เห็นความต่าง)
gdf["geometry"] = gdf.geometry.simplify(0.0005, preserve_topology=True)

b = gdf.total_bounds  # [minx, miny, maxx, maxy]
center = [(b[1] + b[3]) / 2, (b[0] + b[2]) / 2]

m = folium.Map(location=center, zoom_start=6, tiles=None, control_scale=True)

google_layers = [
    ("m", "Google Maps"),
    ("s", "Google Satellite"),
    ("y", "Google Hybrid"),
    ("p", "Google Terrain"),
]
for code, name in google_layers:
    folium.TileLayer(
        tiles=f"https://mt1.google.com/vt/lyrs={code}&x={{x}}&y={{y}}&z={{z}}",
        attr="© Google",
        name=name,
        max_zoom=20,
        show=(code == "m"),
    ).add_to(m)

folium.GeoJson(
    gdf,
    name="ขอบเขตประเทศไทย",
    style_function=lambda f: {
        "fillColor": "#1F4E79",
        "color": "#E63946",
        "weight": 2.5,
        "fillOpacity": 0.15,
    },
    highlight_function=lambda f: {"fillOpacity": 0.35, "weight": 4},
    tooltip="ประเทศไทย (Thailand)",
).add_to(m)

folium.LayerControl(collapsed=False).add_to(m)
m.fit_bounds([[b[1], b[0]], [b[3], b[2]]])

title_html = """
<div style="position: fixed; top: 12px; left: 60px; z-index: 9999;
            background: rgba(255,255,255,.92); padding: 8px 16px; border-radius: 8px;
            box-shadow: 0 2px 6px rgba(0,0,0,.3);
            font-family: 'Segoe UI','Leelawadee UI',sans-serif;">
  <b style="color:#1F4E79; font-size:15px;">ขอบเขตประเทศไทย (Thailand Boundary)</b><br>
  <span style="font-size:11px; color:#666;">ที่มา: Thailand.geojson (WGS84) | พื้นหลัง: Google Maps</span>
</div>"""
m.get_root().html.add_child(folium.Element(title_html))

m.save(OUT)
print("saved:", OUT)
