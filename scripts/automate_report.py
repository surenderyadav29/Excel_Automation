import os
import sys
import time
import logging
import warnings
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, AreaChart, Reference

warnings.filterwarnings("ignore")

# --- 1. ENTERPRISE LOGGING SETUP ---
LOG_FILE = "output/automation.log"
os.makedirs("output", exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.FileHandler(LOG_FILE), logging.StreamHandler(sys.stdout)]
)

start_time = time.time()
logging.info("Starting Multi-Sheet Elite Autonomous Analytics Engine for BlinkIT Dataset...")

# --- 2. DATA INGESTION ---
RAW_DATA_PATH = "BlinkIT Grocery Data.xlsx"
OUTPUT_PATH = "output/BlinkIT_Executive_Dashboard.xlsx"

if not os.path.exists(RAW_DATA_PATH):
    logging.error(f"Dataset missing at {RAW_DATA_PATH}. Execution halted.")
    exit()

# Reading the BlinkIT excel file directly
df = pd.read_excel(RAW_DATA_PATH)
total_rows, total_cols = df.shape

# Core metrics definition for BlinkIT
primary_cat = 'Item Type'
primary_num = 'Sales'
secondary_num = 'Rating'
secondary_cat = 'Outlet Type'
location_cat = 'Outlet Location Type'
fat_cat = 'Item Fat Content'

# --- 3. ADVANCED STATISTICAL MATH (Anomaly Detection via Z-Score) ---
anomalies_msg = "💡 NO SIGNIFICANT ANOMALIES: All operational metrics fall within normal standard distribution limits."
if primary_num in df.columns:
    mean_val = df[primary_num].mean()
    std_val = df[primary_num].std()
    if std_val > 0:
        df['z_score'] = (df[primary_num] - mean_val) / std_val
        outliers = df[df['z_score'].abs() > 2.0]
        if not outliers.empty:
            anomalies_msg = f"⚠️ ANOMALY DETECTED: Found {len(outliers)} statistical outlier(s) in '{primary_num}'. Peak anomaly row index {outliers[primary_num].idxmax()} with value: {outliers[primary_num].max():,.2f}."

# --- 4. PRE-COMPUTED AGGREGATIONS FOR EXECUTIVE MULTI-SHEET ARCHITECTURE ---
g1 = df.groupby(primary_cat)[primary_num].sum().reset_index().sort_values(by=primary_num, ascending=False).head(8)
insights_1 = f"📊 DISTRIBUTION STRATA: Product category leader is '{g1.iloc[0][primary_cat]}' managing maximum accumulation in '{primary_num}' of {g1.iloc[0][primary_num]:,.2f}."

g2 = df.groupby(secondary_cat)[primary_num].sum().reset_index().sort_values(by=primary_num, ascending=False)
insights_2 = f"📈 CORE CHANNEL REPORT: Analyzing secondary organizational layer '{secondary_cat}' points out '{g2.iloc[0][secondary_cat]}' as primary core revenue engine."

g3 = df.groupby(location_cat)[primary_num].sum().reset_index().sort_values(by=primary_num, ascending=False)
share_pct = (g3.iloc[0][primary_num] / df[primary_num].sum()) * 100
insights_3 = f"🍕 GEOGRAPHIC SHARE PROFILE: '{g3.iloc[0][location_cat]}' secures an imposing organizational grasp of {share_pct:.2f}% inside global revenue matrix."

g4 = df.groupby(fat_cat).size().reset_index(name='Volume_Count').sort_values(by='Volume_Count', ascending=False)
insights_4 = f"🎯 OPERATIONAL DENSITY: Volumetric product load hits concentration peak at element '{g4.iloc[0][fat_cat]}' with frequency count of {g4.iloc[0]['Volume_Count']} items."

# --- 5. WORKBOOK STYLING TEMPLATE ---
wb = Workbook()

# Theme Elements (Premium Charcoal & Slate)
font_banner = Font(name="Segoe UI", size=16, bold=True, color="FFFFFF")
font_sec_hdr = Font(name="Segoe UI", size=12, bold=True, color="1A252C")
font_card_lbl = Font(name="Segoe UI", size=9, bold=True, color="7F8C8D")
font_card_val = Font(name="Segoe UI", size=15, bold=True, color="2C3E50")
font_tbl_hdr = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
font_insight = Font(name="Segoe UI", size=11, color="2C3E50")
font_cell = Font(name="Segoe UI", size=10, color="000000")

fill_banner = PatternFill(start_color="1A252C", end_color="1A252C", fill_type="solid")
fill_tbl_hdr = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
fill_insight = PatternFill(start_color="F8F9F9", end_color="F8F9F9", fill_type="solid")
fill_card = PatternFill(start_color="EAEDED", end_color="EAEDED", fill_type="solid")

thin_gray = Side(style='thin', color='BDC3C7')
border_all = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
border_card = Border(left=Side(style='medium', color='1A252C'), right=thin_gray, top=thin_gray, bottom=thin_gray)

def setup_sheet_banner(ws, title_text):
    ws.views.sheetView[0].showGridLines = True
    ws.merge_cells("A1:G1")
    ws["A1"] = title_text.upper()
    ws["A1"].font = font_banner
    ws["A1"].fill = fill_banner
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

def build_grid_table(ws, dataframe, st_r, st_c, h1, h2):
    ws.cell(row=st_r, column=st_c, value=h1).font = font_tbl_hdr
    ws.cell(row=st_r, column=st_c).fill = fill_tbl_hdr
    ws.cell(row=st_r, column=st_c+1, value=h2).font = font_tbl_hdr
    ws.cell(row=st_r, column=st_c+1).fill = fill_tbl_hdr
    
    curr = st_r + 1
    for _, row in dataframe.iterrows():
        cell_a = ws.cell(row=curr, column=st_c, value=row.iloc[0])
        cell_b = ws.cell(row=curr, column=st_c+1, value=row.iloc[1])
        cell_a.font = font_cell; cell_a.border = border_all
        cell_b.font = font_cell; cell_b.border = border_all; cell_b.number_format = '#,##0.00' if isinstance(row.iloc[1], (int, float)) else '@'
        curr += 1
    return curr - 1

# --- SHEET 1: EXECUTIVE OVERVIEW ---
ws_summary = wb.active
ws_summary.title = "Executive Overview"
setup_sheet_banner(ws_summary, "BlinkIT Executive Performance KPI Overview")

# KPI Cards Generation
kpi_metrics = [
    {"label": f"TOTAL SALES", "val": df[primary_num].sum(), "format": "$#,##0.00"},
    {"label": f"AVERAGE RATING", "val": df[secondary_num].mean(), "format": "0.00"},
    {"label": "TOTAL UNIQUE ITEMS", "val": df['Item Identifier'].nunique(), "format": "#,##0"},
    {"label": f"ACTIVE OUTLETS", "val": df['Outlet Identifier'].nunique(), "format": "#,##0"}
]
card_cols = [("A", "B"), ("C", "D"), ("E", "F"), ("G", "G")]
for idx, metric in enumerate(kpi_metrics):
    c_start, c_end = card_cols[idx]
    ws_summary.merge_cells(f"{c_start}3:{c_end}3")
    ws_summary.merge_cells(f"{c_start}4:{c_end}4")
    
    lbl = ws_summary[f"{c_start}3"]
    lbl.value = metric["label"]; lbl.font = font_card_lbl; lbl.alignment = Alignment(horizontal="center", vertical="center")
    
    val = ws_summary[f"{c_start}4"]
    val.value = metric["val"]; val.font = font_card_val; val.alignment = Alignment(horizontal="center", vertical="center")
    val.number_format = metric["format"]
    
    for r in [3, 4]:
        ws_summary.row_dimensions[r].height = 20
        for col_char in [chr(alpha) for alpha in range(ord(c_start), ord(c_end)+1)]:
            cell = ws_summary[f"{col_char}{r}"]
            cell.fill = fill_card; cell.border = border_card

# Automated Insights Panel on Main Page
ws_summary.cell(row=6, column=1, value="⚡ ENTERPRISE SYSTEM ANALYTICAL INTERPRETATIONS").font = font_sec_hdr
all_insights = [insights_1, insights_2, insights_3, insights_4, anomalies_msg]
for i, ins in enumerate(all_insights, start=7):
    ws_summary.row_dimensions[i].height = 25
    ws_summary.merge_cells(start_row=i, start_column=1, end_row=i, end_column=7)
    cell = ws_summary.cell(row=i, column=1, value=ins)
    cell.font = font_insight; cell.fill = fill_insight; cell.alignment = Alignment(vertical="center", indent=1)

# --- SHEET 2: BAR CHART (Product Distribution) ---
ws_s2 = wb.create_sheet(title="Product Performance")
setup_sheet_banner(ws_s2, f"Distribution: {primary_cat} Sales")
ws_s2.cell(row=3, column=1, value=insights_1).font = font_insight
end_r1 = build_grid_table(ws_s2, g1, 5, 1, primary_cat, "Total_Sales")

chart1 = BarChart()
chart1.type = "col"; chart1.style = 10; chart1.title = f"Top Categories Matrix by Sales"
chart1.add_data(Reference(ws_s2, min_col=2, min_row=5, max_row=end_r1), titles_from_data=True)
chart1.set_categories(Reference(ws_s2, min_col=1, min_row=6, max_row=end_r1))
chart1.width, chart1.height = 16, 11
ws_s2.add_chart(chart1, "D5")

# --- SHEET 3: LINE CHART (Channel Analysis) ---
ws_s3 = wb.create_sheet(title="Channel Matrix")
setup_sheet_banner(ws_s3, "Outlet Type Operational Performance")
ws_s3.cell(row=3, column=1, value=insights_2).font = font_insight
end_r2 = build_grid_table(ws_s3, g2, 5, 1, secondary_cat, "Total_Sales")

chart2 = LineChart()
chart2.style = 13; chart2.title = f"Revenue Velocity via Outlet Blueprint"
chart2.add_data(Reference(ws_s3, min_col=2, min_row=5, max_row=end_r2), titles_from_data=True)
chart2.set_categories(Reference(ws_s3, min_col=1, min_row=6, max_row=end_r2))
chart2.width, chart2.height = 16, 11
ws_s3.add_chart(chart2, "D5")

# --- SHEET 4: PIE CHART (Geographic Composition) ---
ws_s4 = wb.create_sheet(title="Regional Share")
setup_sheet_banner(ws_s4, "Geographic Regional Volume Share")
ws_s4.cell(row=3, column=1, value=insights_3).font = font_insight
end_r3 = build_grid_table(ws_s4, g3, 5, 1, location_cat, "Total_Sales")

chart3 = PieChart()
chart3.title = f"Composition Profile: Regional Spread"
chart3.add_data(Reference(ws_s4, min_col=2, min_row=5, max_row=end_r3), titles_from_data=True)
chart3.set_categories(Reference(ws_s4, min_col=1, min_row=6, max_row=end_r3))
chart3.width, chart3.height = 16, 11
ws_s4.add_chart(chart3, "D5")

# --- SHEET 5: AREA CHART (Fat Content Density) ---
ws_s5 = wb.create_sheet(title="Product Density")
setup_sheet_banner(ws_s5, "Item Composition Intensity Load")
ws_s5.cell(row=3, column=1, value=insights_4).font = font_insight
end_r4 = build_grid_table(ws_s5, g4, 5, 1, fat_cat, "Volume_Count")

chart4 = AreaChart()
chart4.style = 11; chart4.title = f"Operational Intensity Profile: Fat Content"
chart4.add_data(Reference(ws_s5, min_col=2, min_row=5, max_row=end_r4), titles_from_data=True)
chart4.set_categories(Reference(ws_s5, min_col=1, min_row=6, max_row=end_r4))
chart4.width, chart4.height = 16, 11
ws_s5.add_chart(chart4, "D5")

# --- SHEET 6: CLEAN SOURCE DATA BACKUP ---
ws_data = wb.create_sheet(title="Clean Dataset Source")
ws_data.views.sheetView[0].showGridLines = True
# Drop temporary column before writing backup
for r in dataframe_to_rows(df.drop(columns=['z_score'], errors='ignore'), index=False, header=True):
    ws_data.append(r)
for col_idx in range(1, ws_data.max_column + 1):
    cell = ws_data.cell(row=1, column=col_idx)
    cell.font = font_tbl_hdr; cell.fill = fill_tbl_hdr

# --- 6. AUTO-FIT COLUMN WIDTHS ACROSS ALL SHEETS ---
for ws in wb.worksheets:
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        valid_lengths = [len(str(cell.value)) for cell in col if cell.value is not None]
        max_len = max(valid_lengths) if valid_lengths else 0
        if max_len > 0:
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 35)
        else:
            ws.column_dimensions[col_letter].width = 11

# Save Final Refactored Output
wb.save(OUTPUT_PATH)

execution_time = (time.time() - start_time) * 1000
logging.info(f"[SUCCESS] BlinkIT Executive Architecture Executed perfectly in {execution_time:.2f}ms.")
logging.info(f"Workbook compiled and saved successfully to -> {OUTPUT_PATH}")