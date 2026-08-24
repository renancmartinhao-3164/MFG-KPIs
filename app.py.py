"""
Manufacturing South America Performance Dashboard
Single-file Streamlit application ready for GitHub.

Expected repository structure:
    app.py
    requirements.txt
    data/inputdataSA.xlsx
    assets/agco_logo.jpg  # optional

Run with: streamlit run app.py
"""
from pathlib import Path
import re
import math
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go


AGCO_GREEN = "#4C8C2B"
AGCO_DARK = "#24451D"
AMBER = "#F5A623"
RED = "#D64545"
GRAY = "#667085"

SITE_COUNTRY = {
    "Canoas": "Brazil", "Mogi das Cruzes": "Brazil", "Santa Rosa": "Brazil",
    "Ibirubá": "Brazil", "Gen.Rodríguez": "Argentina",
    "Consolidado SA": "South America",
}

# First matching column is used. direction: high = higher is better, low = lower is better.
KPI_CATALOG = {
    "Safety": [
        {"name":"TRIR", "actual":["TRIR","TCIR_Month","TCIR_R12"], "target":["TCIR_M"], "direction":"low", "fmt":".2f"},
        {"name":"LTIR", "actual":["LTIR_Month","LTIR_R12"], "target":["LTIR_M"], "direction":"low", "fmt":".2f"},
        {"name":"Safety Observations", "actual":["Observations per employee_M","Observations per employee_YTD"], "target":[], "direction":"high", "fmt":".2f"},
    ],
    "Quality": [
        {"name":"RFT / FPY", "actual":["RFT_Month","RFT_Month_M"], "target":[], "direction":"high", "fmt":".1%"},
        {"name":"Claim Rate", "actual":["Claim Rate_R12"], "target":[], "direction":"low", "fmt":".2f"},
        {"name":"DPU", "actual":["DPU TOTAL_Month","DPU TOTAL_YTD"], "target":[], "direction":"low", "fmt":".2f"},
    ],
    "Delivery": [
        {"name":"PAT", "actual":["PAT_Month_%","PAT_Plan_%"], "target":[], "direction":"high", "fmt":".1%"},
        {"name":"Offline", "actual":["Offline_M","Offline_Actual"], "target":[], "direction":"high", "fmt":".1%"},
        {"name":"Blued", "actual":["Blued_M","Blued_Actual"], "target":[], "direction":"high", "fmt":".1%"},
    ],
    "Cost": [
        {"name":"Conversion Cost", "actual":["Conversion Cost","Conversion_Cost"], "target":["Conversion Cost Plan"], "direction":"low", "fmt":",.0f"},
        {"name":"Cost per Unit", "actual":["Cost per Unit","Cost_per_Unit"], "target":["Cost per Unit Plan"], "direction":"low", "fmt":",.0f"},
        {"name":"Savings", "actual":["Savings","Savings_Actual"], "target":["Savings Plan"], "direction":"high", "fmt":",.0f"},
    ],
    "Inventory": [
        {"name":"Inventory Turns", "actual":["Turns_Actual"], "target":["Turns_Plan"], "direction":"high", "fmt":".2f"},
        {"name":"DOH", "actual":["DOH_Actual"], "target":["DOH_Plan"], "direction":"low", "fmt":".1f"},
        {"name":"E&O", "actual":["E&O_Actual"], "target":["E&O/_M"], "direction":"low", "fmt":",.0f"},
        {"name":"Net Inventory", "actual":["Total Net Inventory"], "target":[], "direction":"low", "fmt":",.0f"},
    ],
    "People": [
        {"name":"Training Hours", "actual":["Training hours per employee_YTD"], "target":["Training hours per employee_M"], "direction":"high", "fmt":".2f"},
        {"name":"Headcount", "actual":["Headcount","HC"], "target":["Headcount Plan"], "direction":"low", "fmt":",.0f"},
        {"name":"Absenteeism", "actual":["Absenteeism","Absenteísmo"], "target":["Absenteeism Plan"], "direction":"low", "fmt":".1%"},
    ],
}


def clean_name(value):
    return re.sub(r"\s+", " ", str(value).replace("\n", " ").strip())

def match_column(columns, aliases):
    lookup = {clean_name(c).casefold(): c for c in columns}
    return next((lookup[clean_name(a).casefold()] for a in aliases if clean_name(a).casefold() in lookup), None)

@st.cache_data(show_spinner="Reading manufacturing workbook...")
def load_workbook(path_string, modified_ns, size):
    """Cache key includes file timestamp and size, so replacement triggers reload."""
    path = Path(path_string)
    excel = pd.ExcelFile(path, engine="openpyxl")
    sheet_lookup = {clean_name(s).casefold(): s for s in excel.sheet_names}
    fact_sheet = sheet_lookup.get("general", excel.sheet_names[0])
    df = pd.read_excel(path, sheet_name=fact_sheet, engine="openpyxl")
    df.columns = [clean_name(c) for c in df.columns]
    site = match_column(df.columns, ["Site", "Plant", "Location"])
    date = match_column(df.columns, ["Month", "Date", "Mês", "Mes"])
    if not site or not date:
        raise ValueError("The fact sheet must contain Site and Month/Date columns.")
    df = df.rename(columns={site:"Site", date:"Date"})
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"]).copy()
    df["Site"] = df["Site"].astype(str).str.strip()
    df["Year"] = df["Date"].dt.year
    df["MonthNo"] = df["Date"].dt.month
    df["MonthLabel"] = df["Date"].dt.strftime("%b")
    df["Country"] = df["Site"].map(SITE_COUNTRY).fillna("Not mapped")
    df["Region"] = "South America"
    protected = {"Site","Date","Country","Region","MonthLabel"}
    for col in df.columns:
        if col not in protected:
            df[col] = pd.to_numeric(df[col].replace({"-":None,"–":None,"—":None}), errors="coerce")
    targets = pd.DataFrame()
    if "metas" in sheet_lookup:
        targets = pd.read_excel(path, sheet_name=sheet_lookup["metas"], engine="openpyxl")
        targets.columns = [clean_name(c) for c in targets.columns]
    return df, targets, excel.sheet_names

def get_data(path):
    stat = path.stat()
    return load_workbook(str(path), stat.st_mtime_ns, stat.st_size)


def finite(value):
    try:
        value=float(value); return value if math.isfinite(value) else None
    except (TypeError,ValueError): return None

def resolve(columns, aliases):
    lookup={str(c).strip().casefold():c for c in columns}
    return next((lookup[a.strip().casefold()] for a in aliases if a.strip().casefold() in lookup),None)

def summarize(df, definition):
    actual_col=resolve(df.columns,definition["actual"])
    target_col=resolve(df.columns,definition["target"])
    if not actual_col:return None
    cols=["Date",actual_col]+([target_col] if target_col and target_col!=actual_col else [])
    monthly=df[cols].dropna(subset=[actual_col]).groupby("Date",as_index=False).mean(numeric_only=True).sort_values("Date").tail(12)
    if monthly.empty:return None
    actual=finite(monthly.iloc[-1][actual_col]); target=finite(monthly.iloc[-1][target_col]) if target_col else None
    values=monthly[actual_col].dropna()
    return {"actual_col":actual_col,"target_col":target_col,"actual":actual,"target":target,
      "gap":None if actual is None or target is None else actual-target,"avg":values.mean(),
      "best":values.max() if definition["direction"]=="high" else values.min(),
      "worst":values.min() if definition["direction"]=="high" else values.max(),"monthly":monthly}

def status(actual,target,direction):
    if actual is None or target is None:return "⚪","No target",None
    ratio=(actual/target if target else None) if direction=="high" else (target/actual if actual else 1.0)
    if ratio is None:return "⚪","No target",None
    if ratio>=1:return "🟢","Target achieved",ratio
    if ratio>=.95:return "🟡","Within 5% of target",ratio
    return "🔴","More than 5% from target",ratio

def fmt(value,spec): return "N/A" if finite(value) is None else format(float(value),spec)

def insight(label,summary,direction,spec):
    if not summary:return f"{label}: no valid data for the selected filters."
    values=summary["monthly"][summary["actual_col"]].dropna()
    notes=[]
    if len(values)>=2:
        delta=values.iloc[-1]-values.iloc[-2]
        good=delta>=0 if direction=="high" else delta<=0
        notes.append(f"{label} {'improved' if good else 'deteriorated'} by {fmt(abs(delta),spec)} versus the prior month.")
    if len(values)>=3 and summary["target"] is not None:
        below=(values.tail(3)<summary["target"]).all() if direction=="high" else (values.tail(3)>summary["target"]).all()
        if below:notes.append(f"{label} remained below target during the latest three valid observations.")
    if summary["actual"]==summary["best"]:notes.append(f"{label} is the best result in the latest 12 valid observations.")
    return " ".join(notes) or f"{label}: insufficient history for a comparative insight."


def style(fig,title):
    fig.update_layout(title=title,template="plotly_white",height=390,margin=dict(l=20,r=15,t=55,b=20),
      font=dict(family="Arial",color="#252A2D"),paper_bgcolor="white",plot_bgcolor="white",
      legend=dict(orientation="h",y=1.08,x=0),hoverlabel=dict(bgcolor="white"))
    fig.update_xaxes(showgrid=False);fig.update_yaxes(gridcolor="#E8ECEF")
    return fig

def empty(title):
    fig=go.Figure();fig.add_annotation(text="No data available",x=.5,y=.5,showarrow=False,font=dict(size=16,color=GRAY))
    fig.update_xaxes(visible=False);fig.update_yaxes(visible=False);return style(fig,title)

def trend(summary,label):
    if not summary:return empty(f"{label} | 12-month trend")
    d=summary["monthly"].copy();col=summary["actual_col"];d["MA3"]=d[col].rolling(3,min_periods=1).mean()
    fig=go.Figure([go.Scatter(x=d.Date,y=d[col],name=label,mode="lines+markers",line=dict(color=AGCO_GREEN,width=3)),
                   go.Scatter(x=d.Date,y=d.MA3,name="3M moving average",mode="lines",line=dict(color=AGCO_DARK,dash="dot"))])
    if summary["target"] is not None:fig.add_hline(y=summary["target"],line_dash="dash",line_color=AMBER,annotation_text="Target")
    return style(fig,f"{label} | 12-month trend")

def site_chart(df,col,label,direction):
    if not col or df.empty:return empty(f"{label} | Site comparison")
    d=df.dropna(subset=[col]).sort_values("Date").groupby("Site",as_index=False).tail(1)
    d=d[d.Site!="Consolidado SA"].sort_values(col,ascending=(direction=="low"))
    if d.empty:return empty(f"{label} | Site comparison")
    fig=px.bar(d,x=col,y="Site",orientation="h",color=col,color_continuous_scale=["#D64545","#F5A623","#4C8C2B"])
    fig.update_coloraxes(showscale=False);return style(fig,f"{label} | Site comparison")


st.set_page_config(page_title="Manufacturing SA Performance",page_icon="📊",layout="wide",initial_sidebar_state="expanded")
BASE=Path(__file__).resolve().parent
DATA_FILE=BASE/"data"/"inputdataSA.xlsx"
LOGO=BASE/"assets"/"agco_logo.jpg"

st.markdown("""<style>
.block-container{padding-top:1.1rem}.main-title{text-align:center;color:#252A2D;margin:.2rem 0}.subtle{text-align:center;color:#667085}
[data-testid="stMetric"]{background:white;border-radius:14px;padding:16px;border-left:5px solid #4C8C2B;box-shadow:0 4px 16px rgba(25,40,32,.07)}
.insight{background:white;border-left:5px solid #4C8C2B;padding:16px 20px;border-radius:12px;box-shadow:0 4px 16px rgba(25,40,32,.07)}
.small{color:#667085;font-size:.78rem}
</style>""",unsafe_allow_html=True)

if not DATA_FILE.exists():
    st.error("Excel file not found. Place inputdataSA.xlsx in the data folder.")
    st.code(str(DATA_FILE));st.stop()

try:df,targets,sheets=get_data(DATA_FILE)
except Exception as exc:st.exception(exc);st.stop()

head1,head2,head3=st.columns([1,4,1])
with head1:
    if LOGO.exists():st.image(str(LOGO),width=145)
with head2:
    st.markdown('<h1 class="main-title">Manufacturing South America Performance Dashboard</h1>',unsafe_allow_html=True)
    st.markdown('<div class="subtle">Safety • Quality • Delivery • Cost • Inventory • People</div>',unsafe_allow_html=True)
with head3:
    stamp=pd.Timestamp(DATA_FILE.stat().st_mtime,unit="s").strftime("%d %b %Y %H:%M")
    st.caption(f"Excel updated\n\n{stamp}")

st.sidebar.header("Global Filters")
def choose(label,values,default="All"):
    options=["All"]+sorted(pd.Series(values).dropna().unique().tolist())
    return st.sidebar.selectbox(label,options,index=0 if default=="All" else options.index(default))
region=choose("Region",df.Region);country=choose("Country",df.Country);site=choose("Site",df.Site)
years=sorted(df.Year.dropna().unique().tolist());year=st.sidebar.selectbox("Year",["All"]+years,index=len(years) if years else 0)
month=st.sidebar.selectbox("Month",["All"]+list(range(1,13)),format_func=lambda x:"All" if x=="All" else pd.Timestamp(2025,int(x),1).strftime("%B"))
filtered=df.copy()
for col,val in [("Region",region),("Country",country),("Site",site),("Year",year),("MonthNo",month)]:
    if val!="All":filtered=filtered[filtered[col]==val]
st.sidebar.caption(f"{len(filtered):,} rows after filters")

page=st.sidebar.radio("Navigation",["Executive Overview","Site Benchmark"])
available=[]
for pillar,items in KPI_CATALOG.items():
    for item in items:
        s=summarize(filtered,item)
        if s:available.append((pillar,item,s))

if page=="Executive Overview":
    st.subheader("Executive Overview")
    cols=st.columns(6)
    for idx,(pillar,items) in enumerate(KPI_CATALOG.items()):
        candidate=next(((item,summarize(filtered,item)) for item in items if summarize(filtered,item)),None)
        with cols[idx]:
            if candidate:
                item,s=candidate;icon,label,_=status(s["actual"],s["target"],item["direction"])
                st.metric(f"{icon} {pillar} | {item['name']}",fmt(s["actual"],item["fmt"]),None if s["gap"] is None else f"Gap {fmt(s['gap'],item['fmt'])}",help=f"Target: {fmt(s['target'],item['fmt'])} | 12M average: {fmt(s['avg'],item['fmt'])} | Best: {fmt(s['best'],item['fmt'])} | Worst: {fmt(s['worst'],item['fmt'])}")
                st.caption(label)
            else:st.metric(f"⚪ {pillar}","N/A",help="No configured KPI column was found.")
    if not available:st.warning("No configured KPI columns were found in this selection.");st.stop()
    labels=[f"{p} | {i['name']}" for p,i,s in available]
    selected=st.selectbox("Detailed KPI",labels)
    pillar,item,summary=available[labels.index(selected)]
    st.markdown(f'<div class="insight"><b>Executive insight:</b> {insight(item["name"],summary,item["direction"],item["fmt"])}</div>',unsafe_allow_html=True)
    c1,c2=st.columns([1.6,1])
    c1.plotly_chart(trend(summary,item["name"]),use_container_width=True)
    comparison=df.copy()
    for col,val in [("Region",region),("Country",country),("Year",year),("MonthNo",month)]:
        if val!="All":comparison=comparison[comparison[col]==val]
    c2.plotly_chart(site_chart(comparison,summary["actual_col"],item["name"],item["direction"]),use_container_width=True)
    with st.expander("Data quality and workbook information"):
        st.write({"Workbook sheets":sheets,"Fact rows":len(df),"Filtered rows":len(filtered),"Target rows":len(targets)})
        st.dataframe(filtered.tail(20),use_container_width=True)
else:
    st.subheader("Site Benchmark")
    rows=[]
    source=filtered if site=="All" else df[(df.Region==region) if region!="All" else pd.Series(True,index=df.index)]
    for site_name in [x for x in source.Site.dropna().unique() if x!="Consolidado SA"]:
        site_df=source[source.Site==site_name]
        for pillar,items in KPI_CATALOG.items():
            for item in items:
                s=summarize(site_df,item)
                if s and s["target"] is not None:
                    _,_,ratio=status(s["actual"],s["target"],item["direction"])
                    if ratio is not None:rows.append({"Site":site_name,"Pillar":pillar,"KPI":item["name"],"Score":min(max(ratio,0),1.2)*100})
    scores=pd.DataFrame(rows)
    if scores.empty:st.info("Benchmark requires KPIs with both actual and target values.")
    else:
        ranking=scores.groupby("Site",as_index=False).Score.mean().sort_values("Score")
        c1,c2=st.columns([1,1.5])
        fig=px.bar(ranking,x="Score",y="Site",orientation="h",color="Score",color_continuous_scale=["#D64545","#F5A623","#4C8C2B"]);fig.update_coloraxes(showscale=False)
        c1.plotly_chart(style(fig,"Overall Site Ranking"),use_container_width=True)
        pivot=scores.pivot_table(index="Site",columns="KPI",values="Score",aggfunc="mean")
        heat=px.imshow(pivot,aspect="auto",text_auto=".0f",color_continuous_scale=["#D64545","#F5A623","#4C8C2B"],zmin=80,zmax=105)
        c2.plotly_chart(style(heat,"KPI Performance Heatmap"),use_container_width=True)
        st.markdown(f'<div class="insight"><b>Best overall:</b> {ranking.iloc[-1].Site} ({ranking.iloc[-1].Score:.1f}) &nbsp; | &nbsp; <b>Lowest overall:</b> {ranking.iloc[0].Site} ({ranking.iloc[0].Score:.1f})</div>',unsafe_allow_html=True)
