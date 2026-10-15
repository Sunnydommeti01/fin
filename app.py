from pathlib import Path
from io import BytesIO
import re
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title='FIN | Financial Intelligence',page_icon='📊',layout='wide')
st.markdown('''<style>.block-container{max-width:1540px;padding:0.65rem 1.3rem 1.5rem}.stApp{background:#f6f8fc}h1{font-size:1.65rem!important;color:#142744;margin-bottom:0}h2,h3{color:#17365d}div[data-testid="stMetric"]{background:white;border:1px solid #e3eaf3;border-radius:12px;padding:9px 13px}div[data-testid="stMetricValue"]{font-size:1.35rem}div[data-testid="stVerticalBlock"]{gap:.5rem}section[data-testid="stSidebar"]{background:#eff4fa}</style>''',unsafe_allow_html=True)
ROOT=Path(__file__).resolve().parent
COSTS=['Material Cost','Transportation Cost','Installation Cost','Net meter cost','Vendor Margin']
NUMS=['Project Value','Taxable Value','Dynamic Price','Gross Profit']+COSTS
DIM=['Month','Region','KW/PHASE','Module','DCR/NON-DCR']
COLORS={'Revenue':'#2563eb','Spending':'#f59e0b','Profit':'#059669','Material Cost':'#7c3aed','Projects':'#2563eb','Margin %':'#059669'}

def money(v):
    if pd.isna(v):return '—'
    v=float(v);sgn='−' if v<0 else '';v=abs(v)
    if v>=1e7:return f'{sgn}₹{v/1e7:,.2f}Cr'
    if v>=1e5:return f'{sgn}₹{v/1e5:,.2f}L'
    s=f'{v:.0f}';return sgn+'₹'+(s[:-3].replace(',','') if False else indian(s))
def indian(s):
    if len(s)<=3:return s
    tail=s[-3:];front=s[:-3];parts=[]
    while front:parts.insert(0,front[-2:]);front=front[:-2]
    return ','.join(parts+[tail])
def lbl(v,metric):
    if pd.isna(v):return '—'
    if metric in ('Projects','Loss Projects'):return str(int(v))
    if metric=='Margin %':return f'{v:.1f}%'
    return money(v)

def parse_dates(v):
    n=pd.to_numeric(v,errors='coerce')
    dt=pd.to_datetime(v,errors='coerce',dayfirst=True)
    excel=pd.to_datetime(n,unit='D',origin='1899-12-30',errors='coerce')
    return dt.mask(n.between(20000,80000),excel)

@st.cache_data(ttl=90,show_spinner=False)
def read_excel(raw):
    df=pd.read_excel(BytesIO(raw))
    df.columns=df.columns.astype(str).str.strip()
    needed=['Project Value','KW/PHASE','Module','DCR/NON-DCR','Region','INV DATE- BK']
    missing=[x for x in needed if x not in df]
    if missing:raise ValueError('Missing Excel columns: '+', '.join(missing))
    for x in NUMS:
        if x not in df:df[x]=np.nan
        df[x]=pd.to_numeric(df[x].astype(str).str.replace(r'[^0-9.\-]','',regex=True),errors='coerce')
    for x in ['KW/PHASE','Module','DCR/NON-DCR','Region']:
        df[x]=df[x].fillna('Unknown').astype(str).str.strip().replace('', 'Unknown')
    df['DCR/NON-DCR']=df['DCR/NON-DCR'].str.upper().str.replace('NON DCR','NON-DCR',regex=False)
    df['INV DATE- BK']=parse_dates(df['INV DATE- BK'])
    df['Month']=df['INV DATE- BK'].dt.strftime('%Y-%m').fillna('Unknown')
    df['Spending']=df[COSTS].fillna(0).sum(axis=1)
    df['Profit']=df['Gross Profit'].where(df['Gross Profit'].notna(),df['Project Value']-df['Spending'])
    df['Material Cost']=df['Material Cost'].fillna(0)
    df['Cost Complete']=df[COSTS].notna().all(axis=1)
    return df

def summary(df,keys):
    t=df.groupby(keys,dropna=False,observed=True).agg(Projects=('Project Value','size'),Revenue=('Project Value','sum'),Spending=('Spending','sum'),Profit=('Profit','sum'),**{'Material Cost':('Material Cost','sum')}).reset_index()
    t['Margin %']=np.where(t.Revenue.ne(0),100*t.Profit/t.Revenue,np.nan)
    t['Avg Spending']=t.Spending/t.Projects
    t['Avg Profit']=t.Profit/t.Projects
    return t

def fig_style(fig,height=280,legend=False):
    fig.update_layout(height=height,margin=dict(l=5,r=12,t=28 if not legend else 42,b=40),paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',font=dict(color='#263e5b',size=11),xaxis_title=None,yaxis_title=None,legend=dict(orientation='h',x=0,y=1.13,title=None,font=dict(size=10)),bargap=.28)
    fig.update_yaxes(gridcolor='#e8edf5',zerolinecolor='#cbd5e1')
    fig.update_xaxes(tickangle=-20,automargin=True)
    return fig

def bars(t,x,metrics,title,kind='bar',height=280):
    st.markdown(f'**{title}**')
    if t.empty:st.info('No matching data');return
    if kind=='line':
        fig=go.Figure()
        for m in metrics:
            fig.add_trace(go.Scatter(x=t[x],y=t[m],mode='lines+markers+text',name=m,text=[lbl(v,m) for v in t[m]],textposition='top center',line=dict(color=COLORS.get(m,'#7c3aed'),width=2.5),textfont=dict(size=10)))
    else:
        fig=go.Figure()
        for m in metrics:
            fig.add_trace(go.Bar(x=t[x],y=t[m],name=m,marker_color=COLORS.get(m,'#7c3aed'),text=[lbl(v,m) for v in t[m]],textposition='outside',textfont=dict(size=9),cliponaxis=False))
        fig.update_layout(barmode='group')
    fig_style(fig,height,legend=len(metrics)>1)
    fig.update_yaxes(rangemode='tozero')
    st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})

def heat(t,rows,cols,metric,title):
    st.markdown(f'**{title}**')
    if t.empty:st.info('No matching data');return
    p=t.pivot_table(index=rows,columns=cols,values=metric,aggfunc='sum' if metric!='Margin %' else 'mean',observed=True)
    if p.empty:return
    zz=p.to_numpy(dtype=float)
    txt=[[lbl(v,metric) for v in line] for line in zz]
    fig=go.Figure(go.Heatmap(z=zz,x=[str(v) for v in p.columns],y=[str(v) for v in p.index],text=txt,texttemplate='%{text}',textfont=dict(size=10),colorscale='RdYlGn' if metric in ['Profit','Margin %'] else 'Blues',colorbar=dict(title='%',len=.7) if metric=='Margin %' else dict(len=.7),hovertemplate='%{y} × %{x}<br>%{text}<extra></extra>'))
    fig_style(fig,min(550,max(245,75+len(p)*37)))
    st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})

def pivot_table(t,key):
    st.markdown('**Interactive pivot · Excel-style filters**')
    st.caption('Use the filter row beneath each column header, or sort by clicking a header. Select dimensions and metrics below to rebuild the pivot.')
    try:
        from st_aggrid import AgGrid,GridOptionsBuilder
        gb=GridOptionsBuilder.from_dataframe(t)
        gb.configure_default_column(sortable=True,filter=True,floatingFilter=True,resizable=True,minWidth=115)
        gb.configure_grid_options(domLayout='normal',enableRangeSelection=True)
        AgGrid(t,gridOptions=gb.build(),height=min(420,max(180,42*len(t)+85)),theme='alpine',fit_columns_on_grid_load=False,allow_unsafe_jscode=False,key=key)
    except ImportError:
        st.dataframe(t,use_container_width=True,hide_index=True,height=min(420,max(180,42*len(t)+85)))

def insight(t,dim):
    if t.empty:return
    worst=t.loc[t['Margin %'].idxmin()]
    losses=t[t.Profit<0]
    st.caption(f'⚠️ Lowest margin: {worst[dim]} ({lbl(worst["Margin %"],"Margin %")}); profit {money(worst.Profit)}.'+(f' {len(losses)} loss-making group(s).' if len(losses) else ''))

st.title('FIN · Solar Financial Intelligence')
st.caption('Projects → Revenue → Spending → Profitability | Values in Indian ₹ | Visible chart labels')
with st.sidebar:
    up=st.file_uploader('Replace Excel (optional)',type=['xlsx'])
    files=sorted(ROOT.glob('*.xlsx'))
    path=next((f for f in files if f.name=='FIN_Financial_Sample_100_Projects.xlsx'),files[0] if files else None)
    st.caption('Source: '+('Uploaded Excel' if up else path.name if path else 'No Excel found'))
if not up and not path:st.error('Add your Excel file beside app.py or upload it in the sidebar.');st.stop()
try:df=read_excel(up.getvalue() if up else path.read_bytes())
except Exception as e:st.error(str(e));st.stop()
with st.sidebar:
    st.markdown('### Global filters')
    picks={}
    for dim in DIM:
        values=sorted(df[dim].dropna().unique().tolist(),key=str)
        picks[dim]=st.multiselect(dim,values,default=values)
    st.caption('Revenue = Project Value. Spending = five listed costs. Profit = Gross Profit where available, otherwise Revenue − Spending.')
d=df.copy()
for dim in DIM:d=d[d[dim].isin(picks[dim])]
if d.empty:st.warning('No records match these filters.');st.stop()
total=summary(d,[] ) if False else None
rev=d['Project Value'].sum();sp=d.Spending.sum();pr=d.Profit.sum()
a,b,c,e,f=st.columns(5)
a.metric('Projects',f'{len(d):,}');b.metric('Revenue',money(rev));c.metric('Spending',money(sp));e.metric('Profit',money(pr));f.metric('Profit margin',f'{pr/rev*100:.1f}%' if rev else '—')
if abs(rev-sp-pr)>max(100,abs(rev)*.005):st.caption(f'⚠️ Profit reconciliation: Revenue − Spending differs from recorded Gross Profit by {money(rev-sp-pr)}. Verify the accounting definition.')

st.markdown('### Financial story')
section=st.selectbox('Navigate dashboard', ['All analyses','01 · Month: projects & revenue','02 · Region: projects & revenue','03 · Financial trend graph','04 · Monthly profit margin','05 · Month × Region pivot','06 · KW: spending & profit','07 · KW × Month','08 · KW × Region','09 · Profit heatmap (4 dimensions)','10 · DCR / Non-DCR × Region × Month','Interactive Pivot Builder'],label_visibility='collapsed')
def show(i):return section=='All analyses' or section.startswith(f'{i:02d} ·')
month=summary(d[d.Month!='Unknown'],['Month']).sort_values('Month')
region=summary(d,['Region']).sort_values('Revenue',ascending=False)
kw=summary(d,['KW/PHASE']).sort_values('KW/PHASE')
if show(1):
    st.markdown('#### 01 · Month-wise projects & revenue')
    x,y=st.columns(2)
    with x:bars(month,'Month',['Projects'],'Monthly project count')
    with y:bars(month,'Month',['Revenue'],'Monthly revenue')
if show(2):
    st.markdown('#### 02 · Region-wise projects & revenue')
    x,y=st.columns(2)
    with x:bars(region,'Region',['Projects'],'Projects by region')
    with y:bars(region,'Region',['Revenue'],'Revenue by region')
if show(3):
    st.markdown('#### 03 · Monthly financial trends')
    bars(month,'Month',['Revenue','Spending','Profit'],'Revenue vs spending vs profit',height=310)
    x,y=st.columns(2)
    with x:bars(month,'Month',['Material Cost'],'Monthly material cost')
    with y:bars(month,'Month',['Avg Spending'],'Average spend per project')
if show(4):
    st.markdown('#### 04 · Monthly profit margin')
    bars(month,'Month',['Margin %'],'Profit margin %',kind='line')
    insight(month,'Month')
if show(5):
    st.markdown('#### 05 · Month × Region financial pivot')
    mr=summary(d[d.Month!='Unknown'],['Month','Region'])
    met=st.selectbox('Value',['Projects','Revenue','Spending','Profit','Margin %','Material Cost','Avg Spending'],key='mr_metric')
    heat(mr,'Region','Month',met,'Month × Region · '+met)
    pivot_table(mr.sort_values(['Month','Region']),'mr_grid')
if show(6):
    st.markdown('#### 06 · KW-wise spending & profit')
    bars(kw,'KW/PHASE',['Spending','Profit'],'Spending and profit by KW',height=310)
    insight(kw,'KW/PHASE')
if show(7):
    st.markdown('#### 07 · KW × Month spending & profit')
    km=summary(d[d.Month!='Unknown'],['KW/PHASE','Month'])
    met=st.radio('Display',['Spending','Profit'],horizontal=True,key='km_met')
    heat(km,'KW/PHASE','Month',met,'KW × Month · '+met)
    pivot_table(km,'km_grid')
if show(8):
    st.markdown('#### 08 · KW × Region spending & profit')
    kr=summary(d,['KW/PHASE','Region'])
    met=st.radio('Display',['Spending','Profit'],horizontal=True,key='kr_met')
    heat(kr,'KW/PHASE','Region',met,'KW × Region · '+met)
    pivot_table(kr,'kr_grid')
if show(9):
    st.markdown('#### 09 · Profit heatmap · KW × Module × Region × Month')
    st.caption('Choose any two axes; filter the remaining dimensions above. This avoids unreadable four-dimensional charts.')
    a,b,c=st.columns([1,1,1])
    with a:row=st.selectbox('Heatmap rows',DIM,index=2,key='hrow')
    with b:col=st.selectbox('Heatmap columns',[x for x in DIM if x!=row],index=2 if len([x for x in DIM if x!=row])>2 else 0,key='hcol')
    with c:met=st.selectbox('Heatmap value',['Profit','Margin %','Spending','Revenue','Projects'],key='hmet')
    h=summary(d[d.Month!='Unknown'] if 'Month' in [row,col] else d,[row,col])
    heat(h,row,col,met,f'{row} × {col} · {met}')
    pivot_table(h,'heat_grid')
if show(10):
    st.markdown('#### 10 · DCR / Non-DCR × Region × Month')
    x,y=st.columns(2)
    with x:met=st.selectbox('Measure',['Profit','Spending','Revenue','Projects','Margin %'],key='dcr_metric')
    with y:axis=st.selectbox('Compare against',['Region','Month'],key='dcr_axis')
    dd=d[d.Month!='Unknown'] if axis=='Month' else d
    dc=summary(dd,['DCR/NON-DCR',axis])
    heat(dc,'DCR/NON-DCR',axis,met,f'DCR status × {axis} · {met}')
    pivot_table(summary(d[d.Month!='Unknown'],['DCR/NON-DCR','Region','Month']),'dcr_grid')

if section in ('All analyses','Interactive Pivot Builder'):
    st.markdown('#### Interactive Pivot Builder · Custom Excel-style analysis')
    c1,c2=st.columns([2,2])
    with c1:dims=st.multiselect('Group by',DIM,default=['Month','Region'],key='custom_dims')
    with c2:metrics=st.multiselect('Metrics',['Projects','Revenue','Spending','Profit','Margin %','Material Cost','Avg Spending','Avg Profit'],default=['Projects','Revenue','Spending','Profit','Margin %'],key='custom_metrics')
    if dims and metrics:
        pt=summary(d,dims)[dims+metrics].sort_values(dims)
        pivot_table(pt,'custom_grid')
        st.download_button('Download pivot as CSV',pt.to_csv(index=False).encode('utf-8-sig'),'FIN_custom_pivot.csv','text/csv')
    else:st.info('Select at least one dimension and metric.')
with st.expander('Project-level data · filter, sort and export'):
    cols=[x for x in ["RNG's",'LEAD','CUS Name','INV DATE- BK','Month','KW/PHASE','DCR/NON-DCR','Module','Region','Project Value','Material Cost','Transportation Cost','Installation Cost','Net meter cost','Vendor Margin','Spending','Profit','Remarks'] if x in d]
    pivot_table(d[cols].reset_index(drop=True),'projects_grid')
    st.download_button('Export filtered projects CSV',d[cols].to_csv(index=False).encode('utf-8-sig'),'FIN_projects.csv','text/csv')
st.caption('Profit and costs are based on the supplied Excel. Missing cost fields are treated as zero for spending; review incomplete rows before financial decisions.')
