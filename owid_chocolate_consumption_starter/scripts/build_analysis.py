#!/usr/bin/env python3
"""Build an Applied Statistics project from the OWID cocoa supply CSV."""
from pathlib import Path
import re, json, math, statistics, warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, ScatterChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data/raw/chocolate-consumption-per-person.csv"
PROCESSED = ROOT / "data/processed"
CHARTS = ROOT / "charts"
EXCEL = ROOT / "excel"
REPORT = ROOT / "report"
for d in (PROCESSED, CHARTS, EXCEL, REPORT):
    d.mkdir(parents=True, exist_ok=True)

SOURCE_URL = "https://ourworldindata.org/grapher/chocolate-consumption-per-person.csv?v=1&csvType=full&useColumnShortNames=false"
META_URL = "https://ourworldindata.org/grapher/chocolate-consumption-per-person.metadata.json?v=1&csvType=full&useColumnShortNames=false"
VALUE = "Cocoa beans - Food available for consumption (kilograms per year per capita)"
SOURCE_CITATION = ("Food and Agriculture Organization of the United Nations (2025) – processed by Our World in Data. "
                   "“Yearly per capita supply of cocoa beans” [dataset]. Food and Agriculture Organization of the "
                   "United Nations, Food Balances. Retrieved 12 November 2026 from " + 
                   "https://ourworldindata.org/grapher/chocolate-consumption-per-person")
STUDENT = "Aryan S Jain"
USN = "1RUA25CSE0087"
COURSE = "SE2041 Applied Statistics"
FACULTY = "Prof. Varadurga Bhat"
UNIVERSITY = "RV University, Bengaluru"
ACADEMIC_YEAR = "2026–2027"

def read_source():
    if not RAW.exists():
        raise FileNotFoundError(f"Place the downloaded CSV at {RAW}")
    df = pd.read_csv(RAW)
    expected = {"Entity", "Code", "Year", VALUE}
    if not expected.issubset(df.columns):
        raise ValueError(f"Unexpected columns. Expected {expected}; got {list(df.columns)}")
    return df

def prepare_data(raw):
    df = raw.copy()
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")
    df[VALUE] = pd.to_numeric(df[VALUE], errors="coerce")
    # Preserve all rows; missing codes are retained because aggregate/region rows may lack codes.
    df = df.dropna(subset=["Entity", "Year", VALUE]).copy()
    df["Year"] = df["Year"].astype(int)
    df["Entity_Type"] = np.where(df["Code"].fillna("").str.fullmatch(r"[A-Z]{3}"),
                                 "Country/territory", "Region or aggregate")
    df = df.sort_values(["Entity", "Year"]).reset_index(drop=True)
    if df.duplicated(["Entity", "Year"]).any():
        raise ValueError("Duplicate Entity-Year pairs found; inspect source before proceeding.")
    return df

def stats(series):
    s = pd.Series(series).dropna()
    mode_vals = s.mode()
    return {
        "Count": int(s.count()), "Mean": float(s.mean()), "Median": float(s.median()),
        "Mode": float(mode_vals.iloc[0]) if len(mode_vals) else np.nan,
        "Minimum": float(s.min()), "Maximum": float(s.max()), "Range": float(s.max()-s.min()),
        "Sample variance": float(s.var(ddof=1)), "Sample standard deviation": float(s.std(ddof=1)),
        "Q1": float(s.quantile(.25)), "Q2": float(s.quantile(.50)), "Q3": float(s.quantile(.75)),
        "IQR": float(s.quantile(.75)-s.quantile(.25)),
        "P10": float(s.quantile(.10)), "P90": float(s.quantile(.90)),
        "Coefficient of variation (%)": float(s.std(ddof=1)/s.mean()*100) if s.mean()!=0 else np.nan
    }

def make_charts(df, country):
    plt.rcParams.update({"font.family":"DejaVu Sans", "font.size":9})
    latest_year = int(df["Year"].max())
    latest = df[(df["Year"] == latest_year) & (df["Entity_Type"] == "Country/territory")].copy()
    top = latest.nlargest(12, VALUE).sort_values(VALUE)
    paths = {}
    def save(name):
        p = CHARTS / name
        plt.tight_layout()
        plt.savefig(p, dpi=220, bbox_inches="tight")
        plt.close()
        paths[name] = p
    plt.figure(figsize=(8,5))
    plt.barh(top["Entity"], top[VALUE])
    plt.xlabel("kg per person per year")
    plt.title(f"Top 12 countries/territories, {latest_year}")
    save("01_top_countries_latest_year.png")
    plt.figure(figsize=(7,4.5))
    latest["Entity_Type"].value_counts().plot(kind="pie", autopct="%1.1f%%", ylabel="")
    plt.title(f"Composition of entities in {latest_year}")
    save("02_entity_type_pie.png")
    plt.figure(figsize=(7,4.5))
    plt.hist(df[VALUE], bins=25, edgecolor="black")
    plt.xlabel("kg per person per year")
    plt.ylabel("Frequency")
    plt.title("Distribution of per-capita cocoa supply (all entity-years)")
    save("03_supply_histogram.png")
    plt.figure(figsize=(7,4.5))
    plt.boxplot(df[VALUE], vert=True, patch_artist=False)
    plt.ylabel("kg per person per year")
    plt.title("Boxplot of per-capita cocoa supply")
    save("04_supply_boxplot.png")
    # Panel scatter: country/territory rows only; transparency and downsampling for readability.
    panel = df[df["Entity_Type"] == "Country/territory"]
    plotdata = panel.sample(min(3500, len(panel)), random_state=42)
    plt.figure(figsize=(7,4.5))
    plt.scatter(plotdata["Year"], plotdata[VALUE], alpha=.25, s=9)
    z = np.polyfit(panel["Year"], panel[VALUE], 1)
    xx = np.linspace(panel["Year"].min(), panel["Year"].max(), 100)
    plt.plot(xx, np.poly1d(z)(xx), linewidth=2)
    plt.xlabel("Year")
    plt.ylabel("kg per person per year")
    plt.title("Year and per-capita cocoa supply (country-year panel)")
    save("05_year_supply_scatter.png")
    # World aggregate time series if available.
    world = df[df["Code"].eq("OWID_WRL")].sort_values("Year")
    if len(world):
        plt.figure(figsize=(8,4.5))
        plt.plot(world["Year"], world[VALUE], marker="o", markersize=2)
        plt.xlabel("Year")
        plt.ylabel("kg per person per year")
        plt.title("World per-capita cocoa supply over time")
        save("06_world_trend.png")
    return paths

def write_csvs(df):
    clean_path = PROCESSED / "chocolate_consumption_clean.csv"
    df.to_csv(clean_path, index=False)
    return clean_path

def build_workbook(raw, df, paths, panel):
    wb = Workbook()
    ws = wb.active; ws.title = "Read_Me"
    ws.append(["Applied Statistics – Cocoa Supply Analysis"])
    ws.append(["Source", "FAO (2025), processed by Our World in Data"])
    ws.append(["Indicator", VALUE])
    ws.append(["Unit", "kilograms per person per year"])
    ws.append(["Source URL", SOURCE_URL])
    ws.append(["Important interpretation", "Food supply available at consumer level; not measured actual intake; consumer waste is not deducted."])
    ws.append(["Panel note", "Country-year observations repeat entities over time; pooled correlation is descriptive, not an independent-observation causal model."])
    for row in raw.columns:
        pass
    rawws = wb.create_sheet("Raw_Data")
    rawws.append(list(raw.columns))
    for r in raw.itertuples(index=False, name=None): rawws.append(list(r))
    cleanws = wb.create_sheet("Clean_Data")
    cleanws.append(list(df.columns))
    for r in df.itertuples(index=False, name=None): cleanws.append(list(r))
    # Formula-driven calculations for core numeric variables.
    calc = wb.create_sheet("Descriptive_Stats")
    calc.append(["Measure", "Year", "Per-capita cocoa supply (kg/person/year)"])
    n = len(df)
    year_col = f"Clean_Data!C2:C{n+1}"
    val_col = f"Clean_Data!D2:D{n+1}"
    formulas = [
        ("Count", f"=COUNT({year_col})", f"=COUNT({val_col})"),
        ("Mean", f"=AVERAGE({year_col})", f"=AVERAGE({val_col})"),
        ("Median", f"=MEDIAN({year_col})", f"=MEDIAN({val_col})"),
        ("Mode", f"=IFERROR(MODE({year_col}),NA())", f"=IFERROR(MODE({val_col}),NA())"),
        ("Minimum", f"=MIN({year_col})", f"=MIN({val_col})"),
        ("Maximum", f"=MAX({year_col})", f"=MAX({val_col})"),
        ("Range", f"=MAX({year_col})-MIN({year_col})", f"=MAX({val_col})-MIN({val_col})"),
        ("Sample variance", f"=VAR({year_col})", f"=VAR({val_col})"),
        ("Sample standard deviation", f"=STDEV({year_col})", f"=STDEV({val_col})"),
        ("Q1", f"=QUARTILE({year_col},1)", f"=QUARTILE({val_col},1)"),
        ("Median / Q2", f"=QUARTILE({year_col},2)", f"=QUARTILE({val_col},2)"),
        ("Q3", f"=QUARTILE({year_col},3)", f"=QUARTILE({val_col},3)"),
        ("IQR", f"=QUARTILE({year_col},3)-QUARTILE({year_col},1)", f"=QUARTILE({val_col},3)-QUARTILE({val_col},1)"),
        ("P10", f"=PERCENTILE({year_col},0.1)", f"=PERCENTILE({val_col},0.1)"),
        ("P90", f"=PERCENTILE({year_col},0.9)", f"=PERCENTILE({val_col},0.9)"),
        ("CV (%)", "=NA()", f"=STDEV({val_col})/AVERAGE({val_col})*100"),
    ]
    for row in formulas: calc.append(list(row))
    corr = wb.create_sheet("Correlation")
    corr.append(["Analysis", "Value"])
    corr.append(["Pearson correlation: Year vs supply (country/territory rows)", float(panel["Year"].corr(panel[VALUE]))])
    corr.append(["N country-year observations", len(panel)])
    corr.append(["Interpretation", "Descriptive pooled panel association; repeated observations per entity mean standard independent-observation inference is not appropriate."])
    corr.append(["Regression slope", float(np.polyfit(panel["Year"], panel[VALUE], 1)[0])])
    corr.append(["Regression intercept", float(np.polyfit(panel["Year"], panel[VALUE], 1)[1])])
    corr.append(["R squared", float(panel["Year"].corr(panel[VALUE])**2)])
    # Frequency tables
    freq = wb.create_sheet("Frequency_Tables")
    freq.append(["Entity type", "Frequency"])
    for k,v in df["Entity_Type"].value_counts().items(): freq.append([k,int(v)])
    freq.append([])
    freq.append(["Per-capita supply class (kg)", "Frequency"])
    bins = [0, .1, .5, 1, 2, 4, 8, 16]
    labels = ["0–<0.1","0.1–<0.5","0.5–<1","1–<2","2–<4","4–<8","8–16"]
    vc = pd.cut(df[VALUE], bins=bins, labels=labels, include_lowest=True).value_counts().sort_index()
    for k,v in vc.items(): freq.append([str(k),int(v)])
    # Native editable Excel charts. Supporting data are kept on separate sheets.
    chartws = wb.create_sheet("Charts")
    chartws["A1"] = "Editable Excel charts; matching high-resolution PNGs are in charts/."
    latest_year = int(df["Year"].max())
    latest = df[(df["Year"] == latest_year) & (df["Entity_Type"] == "Country/territory")].nlargest(12, VALUE).sort_values(VALUE)
    topws = wb.create_sheet("Top_Countries_Data")
    topws.append(["Entity", "Supply (kg/person/year)"])
    for _,r in latest.iterrows(): topws.append([r.Entity, float(r[VALUE])])
    bar = BarChart(); bar.type = "bar"; bar.style = 10; bar.title = f"Top countries/territories ({latest_year})"
    bar.y_axis.title = "Entity"; bar.x_axis.title = "kg/person/year"
    bar.add_data(Reference(topws,min_col=2,min_row=1,max_row=len(latest)+1),titles_from_data=True)
    bar.set_categories(Reference(topws,min_col=1,min_row=2,max_row=len(latest)+1)); bar.height=8; bar.width=15
    chartws.add_chart(bar,"A3")
    entws = wb.create_sheet("Entity_Type_Data")
    entws.append(["Entity type", "Frequency"])
    for k,v in df["Entity_Type"].value_counts().items(): entws.append([k,int(v)])
    pie = __import__("openpyxl.chart",fromlist=["PieChart"]).PieChart()
    pie.title = f"Entity type composition ({latest_year})"
    pie.add_data(Reference(entws,min_col=2,min_row=1,max_row=3),titles_from_data=True)
    pie.set_categories(Reference(entws,min_col=1,min_row=2,max_row=3)); pie.height=7; pie.width=10
    chartws.add_chart(pie,"A20")
    histws = wb.create_sheet("Histogram_Data")
    histws.append(["Supply class (kg/person/year)", "Frequency"])
    for k,v in vc.items(): histws.append([str(k),int(v)])
    hist = BarChart(); hist.type="col"; hist.title="Per-capita cocoa supply distribution"
    hist.y_axis.title="Frequency"; hist.x_axis.title="Supply class (kg/person/year)"
    hist.add_data(Reference(histws,min_col=2,min_row=1,max_row=len(vc)+1),titles_from_data=True)
    hist.set_categories(Reference(histws,min_col=1,min_row=2,max_row=len(vc)+1)); hist.height=7; hist.width=15
    chartws.add_chart(hist,"J3")
    # Scatter chart for country/territory panel; data are a reproducible display sample.
    scatterws = wb.create_sheet("Scatter_Data")
    scatterws.append(["Year", "Supply (kg/person/year)"])
    scatter_sample = panel.sample(min(3500,len(panel)), random_state=42)
    for _,r in scatter_sample.iterrows(): scatterws.append([int(r.Year),float(r[VALUE])])
    sc = ScatterChart(); sc.title="Year vs per-capita cocoa supply (display sample)"
    sc.x_axis.title="Year"; sc.y_axis.title="kg/person/year"; sc.style=13
    xvalues=Reference(scatterws,min_col=1,min_row=2,max_row=len(scatter_sample)+1)
    yvalues=Reference(scatterws,min_col=2,min_row=2,max_row=len(scatter_sample)+1)
    sc.series.append(Series(yvalues,xvalues,title="Country-year observations")); sc.height=8; sc.width=15
    chartws.add_chart(sc,"J20")
    # World aggregate line chart.
    world = df[df["Code"].eq("OWID_WRL")].sort_values("Year")
    if len(world):
        dataws = wb.create_sheet("World_Trend_Data")
        dataws.append(["Year", "Per-capita cocoa supply"])
        for _,r in world.iterrows(): dataws.append([int(r.Year), float(r[VALUE])])
        ch = LineChart(); ch.title = "World per-capita cocoa supply"; ch.y_axis.title = "kg/person/year"; ch.x_axis.title = "Year"
        ch.add_data(Reference(dataws,min_col=2,min_row=1,max_row=len(world)+1),titles_from_data=True)
        ch.set_categories(Reference(dataws,min_col=1,min_row=2,max_row=len(world)+1)); ch.height=8; ch.width=15
        chartws.add_chart(ch,"A37")
    # Basic styles
    for sh in wb.worksheets:
        sh.freeze_panes = "A2"
        sh.sheet_view.showGridLines = False
        for c in sh[1]:
            c.font = Font(name="Times New Roman", size=12, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for col in sh.columns:
            letter = col[0].column_letter
            maxlen = min(55, max((len(str(x.value or "")) for x in col[:300]), default=12))
            sh.column_dimensions[letter].width = max(14, maxlen+2)
    out = EXCEL / "chocolate_consumption_analysis.xlsx"
    wb.save(out)
    return out

def add_table(doc, headers, rows, font_size=12):
    # Professor requires 12 pt throughout, including table text.
    font_size = 12
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"; table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(headers):
        cell=table.rows[0].cells[i]; cell.text=str(h)
        for r in cell.paragraphs[0].runs: r.bold=True; r.font.name="Times New Roman"; r.font.size=Pt(font_size)
    for row in rows:
        cells=table.add_row().cells
        for i,x in enumerate(row):
            cells[i].text=str(x)
            for p in cells[i].paragraphs:
                for r in p.runs: r.font.name="Times New Roman"; r.font.size=Pt(font_size)
    return table

def build_report(raw, df, panel, paths, clean_path, workbook_path):
    doc=Document()
    sec=doc.sections[0]; sec.page_width=Cm(21); sec.page_height=Cm(29.7)
    sec.top_margin=Cm(2.2); sec.bottom_margin=Cm(2.2); sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)
    normal=doc.styles["Normal"]; normal.font.name="Times New Roman"; normal.font.size=Pt(12)
    normal.paragraph_format.line_spacing=1.15
    for sname in ["Heading 1","Heading 2","Heading 3"]:
        st=doc.styles[sname]; st.font.name="Times New Roman"
    def para(text="", bold=False, align=None):
        p=doc.add_paragraph()
        if align is not None: p.alignment=align
        r=p.add_run(text); r.bold=bold; r.font.name="Times New Roman"; r.font.size=Pt(12)
        p.paragraph_format.line_spacing=1.15
        return p
    def heading(text, level=1):
        p=doc.add_heading(text,level=level); p.paragraph_format.line_spacing=1.15
        for r in p.runs: r.font.name="Times New Roman"
    def picture(filename, caption):
        p=doc.add_picture(str(paths[filename]),width=Inches(5.8))
        doc.paragraphs[-1].alignment=WD_ALIGN_PARAGRAPH.CENTER
        cp=para(caption,align=WD_ALIGN_PARAGRAPH.CENTER)
        for r in cp.runs: r.italic=True
    # Cover
    for _ in range(2): para("")
    para("RV UNIVERSITY, BENGALURU",True,WD_ALIGN_PARAGRAPH.CENTER)
    para("SCHOOL OF COMPUTER SCIENCE AND ENGINEERING",True,WD_ALIGN_PARAGRAPH.CENTER)
    para(COURSE,True,WD_ALIGN_PARAGRAPH.CENTER)
    para("CIE 3 – STATISTICAL ANALYSIS PROJECT AND PRESENTATION",True,WD_ALIGN_PARAGRAPH.CENTER)
    for _ in range(2): para("")
    para("Statistical Analysis of Per-Capita Cocoa Supply",True,WD_ALIGN_PARAGRAPH.CENTER)
    for _ in range(2): para("")
    para("Submitted by",True,WD_ALIGN_PARAGRAPH.CENTER); para(f"{STUDENT} ({USN})",False,WD_ALIGN_PARAGRAPH.CENTER)
    para("Under the guidance of",True,WD_ALIGN_PARAGRAPH.CENTER); para(FACULTY,False,WD_ALIGN_PARAGRAPH.CENTER)
    for _ in range(3): para("")
    para(f"Academic Year {ACADEMIC_YEAR} | 12 November 2026",False,WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()
    heading("CERTIFICATE",1)
    para(f"Certified that the {COURSE} CIE 3 project titled “Statistical Analysis of Per-Capita Cocoa Supply” is carried out by {STUDENT} ({USN}), a student of the School of Computer Science and Engineering, {UNIVERSITY}, during the academic year {ACADEMIC_YEAR}.")
    para("")
    para("Faculty Guide: __________________________                 Program Director: __________________________")
    doc.add_page_break()
    heading("1. Problem Statement and Research Objectives")
    para("This project examines how per-capita cocoa supply varies over time and across countries and regions. The research question is: “Is there an association between year and per-capita cocoa supply across countries?” The analysis uses the annual country/territory observations in the Our World in Data indicator dataset. It describes distributions, compares entities, and examines the pooled association between year and the indicator.")
    heading("2. Dataset Description and Source")
    para(f"The downloaded dataset contains {len(raw):,} rows and four original variables. The observation is an entity-year record. The years span {int(df.Year.min())}–{int(df.Year.max())}; the file contains {df.Entity.nunique()} distinct entities. The indicator is measured in kilograms per person per year.")
    para("The underlying provider is the Food and Agriculture Organization of the United Nations (FAO), with data processed by Our World in Data (OWID). The indicator represents food supply available at the consumer level, not directly measured intake; consumer waste is not deducted. It combines cocoa beans and specified cocoa products. The report therefore uses the term “per-capita cocoa supply” rather than claiming to measure actual chocolate eaten.")
    para("Citation: " + SOURCE_CITATION)
    heading("3. Variables and Measurement Scales")
    rows=[
        ["Entity","Qualitative","Nominal","Country, territory or region name"],
        ["Code","Qualitative","Nominal","Geographic identifier; some aggregate rows have no code"],
        ["Year","Quantitative – discrete","Interval","Calendar year"],
        [VALUE,"Quantitative – continuous","Ratio","Annual food supply in kg/person/year"],
        ["Entity_Type","Qualitative – derived","Nominal","Country/territory or region/aggregate classification based on code format"]
    ]
    add_table(doc,["Variable","Classification","Scale","Description"],rows,8)
    para("The cleaned analytical file retains rows with missing Code values because these can represent OWID aggregate or region entities. Records are excluded only if Entity, Year or the numerical indicator is missing. No values are imputed.")
    heading("4. Data Cleaning and Sampling")
    para(f"The source file has {len(raw):,} observations. After checking required fields and removing rows with missing Entity, Year or indicator values, {len(df):,} observations remain. Missing Code values are retained and classified as region/aggregate rather than treated as missing observations. The full available dataset is used; no random sample is drawn.")
    heading("5. Frequency Tables and Visualizations")
    para("A frequency table is prepared for entity type. The indicator is continuous, so its values are grouped into numerical intervals for a frequency distribution. The latest-year bar chart compares countries/territories, while the pie chart describes the composition of entity types in the latest year.")
    et=df["Entity_Type"].value_counts()
    add_table(doc,["Entity type","Frequency"],[[k,int(v)] for k,v in et.items()],9)
    picture("01_top_countries_latest_year.png","Figure 1. Countries/territories with the highest per-capita cocoa supply in the latest year")
    picture("02_entity_type_pie.png","Figure 2. Entity-type composition in the latest year")
    picture("03_supply_histogram.png","Figure 3. Distribution of per-capita cocoa supply across entity-year observations")
    picture("04_supply_boxplot.png","Figure 4. Box-and-whisker plot of per-capita cocoa supply")
    if "06_world_trend.png" in paths: picture("06_world_trend.png","Figure 5. World aggregate per-capita cocoa supply over time")
    heading("6. Descriptive Statistical Analysis")
    st=stats(df[VALUE])
    rows=[[k, f"{v:.4f}" if isinstance(v,(float,np.floating)) else v] for k,v in st.items()]
    add_table(doc,["Measure","Per-capita cocoa supply"],rows,9)
    para(f"The mean is {st['Mean']:.4f} kg/person/year and the median is {st['Median']:.4f}. The standard deviation is {st['Sample standard deviation']:.4f}, while the coefficient of variation is {st['Coefficient of variation (%)']:.2f}%. The mean exceeding the median may indicate right-skewness; the histogram and boxplot should be considered alongside these summaries. These pooled statistics combine different entities and years and should not be interpreted as the experience of a typical individual country.")
    heading("7. Relationship Between Year and Per-Capita Cocoa Supply")
    r=float(panel["Year"].corr(panel[VALUE]))
    slope,intercept=np.polyfit(panel["Year"],panel[VALUE],1)
    r2=r*r
    picture("05_year_supply_scatter.png","Figure 6. Scatterplot of year and per-capita cocoa supply for country/territory observations (display sample)")
    para(f"For country/territory observations with three-letter codes, the pooled Pearson correlation between Year and the indicator is r = {r:.4f} (n = {len(panel):,}). The fitted descriptive line is supply = {slope:.6f} × year + {intercept:.4f}, with R² = {r2:.4f}. This pooled panel correlation describes the overall linear association in the combined country-year observations.")
    para("Important limitation: observations repeat for the same country/territory across years, and countries differ substantially in their long-run supply levels. The pooled Pearson coefficient therefore does not satisfy the usual independent-observation assumption for conventional significance testing. It can also combine within-country time trends with between-country differences. It is not evidence that calendar year causes a change in supply. A stronger follow-up would estimate country-specific trends or use a panel-data model.")
    heading("8. Conclusion")
    para(f"The dataset provides {len(df):,} entity-year observations from {df.Year.min()} to {df.Year.max()} and includes qualitative entity identifiers alongside the quantitative year and per-capita supply variables. Descriptive statistics and visualizations show the distribution and variation of the indicator. The pooled country-year correlation is {r:.4f}; it should be interpreted cautiously because the observations form a repeated-measures panel. The results describe food supply available per person, not actual consumption after waste.")
    heading("9. Limitations and Future Scope")
    para("The indicator is a food-supply measure and not direct survey measurement of individual consumption. The dataset combines country/territory observations and regional aggregates; aggregate rows must not be mixed with countries when making country-level comparisons. Repeated observations over time are not independent. Future work can focus on a single country’s time series, compare a preselected group of countries, or use panel regression with entity effects.")
    heading("10. References")
    para(SOURCE_CITATION)
    para("Our World in Data. “Cocoa bean consumption per person.” https://ourworldindata.org/grapher/chocolate-consumption-per-person")
    para("Data accessed 12 November 2026. The dataset and OWID-produced materials are made available under CC BY; original FAO data remain subject to the provider’s terms. Credit both FAO and Our World in Data.")
    out=REPORT/"CIE3_Statistical_Analysis_Report.docx"
    doc.save(out)
    return out

def main():
    raw=read_source()
    df=prepare_data(raw)
    clean_path=write_csvs(df)
    paths=make_charts(df, None)
    panel=df[df["Entity_Type"]=="Country/territory"].copy()
    workbook=build_workbook(raw,df,paths,panel)
    report=build_report(raw,df,panel,paths,clean_path,workbook)
    summary={
        "raw_rows":len(raw),"clean_rows":len(df),"entities":int(df.Entity.nunique()),
        "year_min":int(df.Year.min()),"year_max":int(df.Year.max()),
        "country_year_rows":len(panel),"pearson_r_year_supply":float(panel["Year"].corr(panel[VALUE])),
        "clean_csv":str(clean_path.relative_to(ROOT)),"workbook":str(workbook.relative_to(ROOT)),
        "report":str(report.relative_to(ROOT)),"charts":[str(p.relative_to(ROOT)) for p in paths.values()]
    }
    (ROOT/"build_summary.json").write_text(json.dumps(summary,indent=2))
    print(json.dumps(summary,indent=2))
    print("\\nIMPORTANT: Open the XLSX in Excel/LibreOffice to recalculate formulas and inspect chart editability.")
    print("Review the DOCX, especially the panel-correlation explanation and the latest-year chart, before submission.")

if __name__=="__main__":
    main()
