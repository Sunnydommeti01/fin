import re
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title='FIN | Solar Financial Intelligence', page_icon='☀️', layout='wide', initial_sidebar_state='expanded')
st.markdown('''<style>
.block-container{padding-top:1.1rem;padding-bottom:1rem;max-width:1550px}h1{font-size:1.65rem!important;color:#152744;letter-spacing:-.04em}h2,h3{color:#193150}div[data-testid="stMetric"]{background:#fff;border:1px solid #e3eaf2;border-radius:13px;padding:12px 15px;box-shadow:0 2px 10px #102d4d08}div[data-testid="stMetricLabel"]{font-size:.8rem}div[data-testid="stMetricValue"]{font-size:1.35rem}div[data-testid="stTabs"] button{font-weight:650}section[data-testid="stSidebar"]{background:#f4f8fc} .stApp{background:#f9fbfe}div[data-testid="stDataFrame"]{border:1px solid #e3eaf2;border-radius:10px}</style>''',unsafe_allow_html=True)

DATA_FILE = Path(__file__).parent / 'FIN_Financial_Sample_100_Projects.xlsx'
COST_COLS = ['Material Cost','Transportation Cost','Installation Cost','Net meter cost','Vendor Margin']
MONEY_COLS = ['Project Value','Taxable Value','Dynamic Price','Gross Profit']+COST_COLS

@st.cache_data(ttl=60,show_spinner=False)
def load_data(file_bytes=None):
    if file_bytes is None:
        if not DATA_FILE.exists():
            return None
        df=pd.read_excel(DATA_FILE)
    else:
        from io import BytesIO
        df=pd.read_excel(BytesIO(file_bytes))
    df.columns=df.columns.astype(str).str.strip()
    required=['KW/PHASE','DCR/NON-DCR','Module','Region','Project Value']+COST_COLS
    missing=[c for c in required if c not in df.columns]
    if missing: raise ValueError('Missing Excel columns: '+', '.join(missing))
    for c in MONEY_COLS:
        if c not in df: df[c]=np.nan
        df[c]=pd.to_numeric(df[c].astype(str).str.replace(r'[^0-9.\-]','',regex=True),errors='coerce')
    for c in ['KW/PHASE','DCR/NON-DCR','Module','Region']:
        df[c]=df[c].fillna('Unknown').astype(str).str.strip()
    df['DCR/NON-DCR']=df['DCR/NON-DCR'].str.upper().str.replace('NON DCR','NON-DCR',regex=False)
    df['KW']=pd.to_numeric(df['KW/PHASE'].str.extract(r'(\d+(?:\.\d+)?)')[0],errors='coerce')
    for c in ['INV DATE- BK','Dispatch Date']:
        if c in df:
            raw=df[c]
            numeric=pd.to_numeric(raw,errors='coerce')
            dates=pd.to_datetime(raw,errors='coerce',dayfirst=True)
            excel_dates=pd.to_datetime(numeric,unit='D',origin='1899-12-30',errors='coerce')
            df[c]=dates.mask(numeric.between(20000,80000),excel_dates)
    df['Spending']=df[COST_COLS].fillna(0).sum(axis=1)
    # Savings is reported gross profit when supplied; otherwise calculated as project value less spending.
    df['Savings']=df['Gross Profit'].where(df['Gross Profit'].notna(),df['Project Value']-df['Spending'])
    df['Savings %']=np.where(df['Project Value'].ne(0),100*df['Savings']/df['Project Value'],np.nan)
    df['Month']=df['INV DATE- BK'].dt.strftime('%Y-%m').fillna('Unknown') if 'INV DATE- BK' in df else 'Unknown'
    return df

def money(v):
    if pd.isna(v):return '—'
    s='-' if v<0 else ''
    v=abs(float(v))
    if v>=1e7:return f'{s}₹{v/1e7:,.2f} Cr'
    if v>=1e5:return f'{s}₹{v/1e5:,.2f} L'
    return f'{s}₹{v:,.0f}'

def summary(d,by):
    t=d.groupby(by,dropna=False,observed=True).agg(Projects=('Project Value','size'),Revenue=('Project Value','sum'),Spending=('Spending','sum'),Savings=('Savings','sum'),Avg_Savings=('Savings','mean')).reset_index()
    t['Savings %']=np.where(t['Revenue'].ne(0),t['Savings']/t['Revenue']*100,0)
    return t.sort_values('Savings',ascending=False)

def metrics(d):
    a,b,c,d1,e=st.columns(5)
    rev=d['Project Value'].sum(); sav=d['Savings'].sum()
    a.metric('Projects',f'{len(d):,}')
    b.metric('Project Value',money(rev))
    c.metric('Total Spending',money(d['Spending'].sum()))
    d1.metric('Total Savings',money(sav))
    e.metric('Savings Margin',f'{sav/rev*100:.1f}%' if rev else '—')

def chart_bar(t,x='KW/PHASE',title='Spending vs Savings'):
    long=t.melt(id_vars=[x],value_vars=['Spending','Savings'],var_name='Metric',value_name='Amount')
    fig=px.bar(long,x=x,y='Amount',color='Metric',barmode='group',color_discrete_map={'Spending':'#f59e0b','Savings':'#10b981'},title=title)
    fig.update_layout(height=340,margin=dict(l=5,r=5,t=45,b=10),paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',legend=dict(orientation='h',y=1.13,x=0),xaxis_title='',yaxis_title='₹',font=dict(color='#263b57'))
    st.plotly_chart(fig,use_container_width=True)

def table(t):
    out=t.copy()
    for c in ['Revenue','Spending','Savings','Avg_Savings']:
        if c in out:out[c]=out[c].map(money)
    if 'Savings %' in out:out['Savings %']=out['Savings %'].map(lambda x:f'{x:.1f}%')
    st.dataframe(out,hide_index=True,use_container_width=True,height=min(380,45+len(out)*36))

st.title('☀️ FIN | Solar Financial Intelligence')
st.caption('Project economics • Spending • Savings • Profitability')
with st.sidebar:
    st.markdown('### Data source')
    upload=st.file_uploader('Optional: upload updated Excel',type=['xlsx'])
try:
    df=load_data(upload.getvalue() if upload else None)
except Exception as ex:
    st.error(f'Could not read Excel: {ex}');st.stop()
if df is None:
    st.warning('Upload the sample Excel to the same GitHub folder as app.py, or use the sidebar Excel uploader.');st.stop()
with st.sidebar:
    st.markdown('### Filters')
    def choose(col,label):
        options=sorted(df[col].dropna().unique().tolist(),key=str)
        return st.multiselect(label,options,default=options)
    regions=choose('Region','Region')
    dcrs=choose('DCR/NON-DCR','DCR category')
    modules=choose('Module','Module')
    kws=choose('KW/PHASE','KW / Phase')
    months=choose('Month','Invoice month')
    st.caption('Savings = reported Gross Profit where available; otherwise Project Value − listed costs. Spending includes Vendor Margin. Verify the definition before using real financial data.')
mask=df['Region'].isin(regions)&df['DCR/NON-DCR'].isin(dcrs)&df['Module'].isin(modules)&df['KW/PHASE'].isin(kws)&df['Month'].isin(months)
d=df.loc[mask].copy()
if d.empty:st.info('No projects match the selected filters.');st.stop()
metrics(d)

t1,t2,t3,t4,t5=st.tabs(['Overview','KW-wise','Module-wise','KW × Module','DCR / Non-DCR'])
with t1:
    a,b=st.columns([1.1,1])
    with a:
        cost=d[COST_COLS].sum().reset_index();cost.columns=['Cost Type','Amount']
        fig=px.pie(cost,names='Cost Type',values='Amount',hole=.65,title='Where the money is spent',color_discrete_sequence=['#2563eb','#06b6d4','#f59e0b','#10b981','#8b5cf6'])
        fig.update_layout(height=320,margin=dict(l=0,r=0,t=45,b=5),paper_bgcolor='rgba(0,0,0,0)',legend=dict(orientation='h',y=-.15),font=dict(color='#263b57'))
        st.plotly_chart(fig,use_container_width=True)
    with b:
        chart_bar(summary(d,'Region'),x='Region',title='Region-wise spending and savings')
    a,b=st.columns(2)
    with a:
        if 'Month' in d:
            trend=summary(d,'Month').sort_values('Month')
            fig=px.line(trend,x='Month',y=['Spending','Savings'],markers=True,title='Monthly financial trend',color_discrete_sequence=['#f59e0b','#10b981'])
            fig.update_layout(height=320,margin=dict(l=0,r=0,t=45,b=0),paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',yaxis_title='₹',xaxis_title='',legend_title='')
            st.plotly_chart(fig,use_container_width=True)
    with b:
        st.markdown('#### Most profitable regions')
        table(summary(d,'Region'))
with t2:
    t=summary(d,'KW/PHASE')
    chart_bar(t,'KW/PHASE','KW-wise spending vs savings')
    table(t)
with t3:
    t=summary(d,'Module')
    chart_bar(t,'Module','Module-wise spending vs savings')
    table(t)
with t4:
    t=summary(d,['KW/PHASE','Module'])
    st.markdown('#### Savings heatmap — KW × Module')
    pv=t.pivot(index='KW/PHASE',columns='Module',values='Savings').fillna(0)
    if not pv.empty:
        fig=px.imshow(pv,labels=dict(x='Module',y='KW / Phase',color='Savings ₹'),color_continuous_scale='Teal',aspect='auto')
        fig.update_layout(height=max(320,55*len(pv)),margin=dict(l=0,r=0,t=15,b=10),paper_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig,use_container_width=True)
    table(t)
with t5:
    t=summary(d,'DCR/NON-DCR')
    chart_bar(t,'DCR/NON-DCR','DCR vs Non-DCR spending and savings')
    table(t)
    st.markdown('#### DCR × KW × Module combinations')
    combo=summary(d,['DCR/NON-DCR','KW/PHASE','Module'])
    table(combo)

st.markdown('### Project-level financial details')
show=['RNG\'s','LEAD','CUS Name','INV DATE- BK','KW/PHASE','DCR/NON-DCR','Module','Region','Project Value','Spending','Savings','Savings %']+COST_COLS+['Remarks']
show=[c for c in show if c in d]
st.dataframe(d[show],hide_index=True,use_container_width=True,height=330)
st.download_button('⬇ Export filtered financial data',data=d.to_csv(index=False).encode('utf-8-sig'),file_name='FIN_filtered_financials.csv',mime='text/csv')
