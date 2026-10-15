import re
from io import BytesIO
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title='FIN | Financial Intelligence', page_icon='☀️', layout='wide')
st.markdown('''<style>.block-container{max-width:1550px;padding-top:1rem;padding-bottom:1rem}.stApp{background:#f7f9fd}h1{color:#152b4d;font-size:1.75rem!important}h2,h3{color:#203b5e}div[data-testid="stMetric"]{background:white;border:1px solid #e1e8f1;border-radius:12px;padding:10px 14px}div[data-testid="stMetricValue"]{font-size:1.4rem}section[data-testid="stSidebar"]{background:#eff5fb}div[data-testid="stDataFrame"]{border:1px solid #dfe7f0;border-radius:10px}</style>''',unsafe_allow_html=True)
FILE=Path(__file__).parent/'FIN_Financial_Sample_100_Projects.xlsx'
COSTS=['Material Cost','Transportation Cost','Installation Cost','Net meter cost','Vendor Margin']
BLUE='#2563eb'; AMBER='#e99a22'; GREEN='#079b78'; RED='#dc4c59'; NAVY='#203b5e'

def inr(x,compact=True):
    if pd.isna(x):return '—'
    x=float(x); sign='-' if x<0 else ''; a=abs(x)
    if compact and a>=1e7:return f'{sign}₹{a/1e7:,.2f} Cr'
    if compact and a>=1e5:return f'{sign}₹{a/1e5:,.2f} L'
    s=f'{a:.0f}'; z=s[-3:]; s=s[:-3]
    while s: z=s[-2:]+','+z;s=s[:-2]
    return f'{sign}₹{z}'

def pct(x):return f'{x:+.1f}%' if pd.notna(x) else '—'

@st.cache_data(ttl=60,show_spinner=False)
def load_data(raw=None):
    if raw is None:
        if not FILE.exists():return None
        df=pd.read_excel(FILE)
    else:df=pd.read_excel(BytesIO(raw))
    df.columns=df.columns.astype(str).str.strip()
    req=['KW/PHASE','DCR/NON-DCR','Module','Region','Project Value']+COSTS
    missing=[c for c in req if c not in df.columns]
    if missing:raise ValueError('Missing columns: '+', '.join(missing))
    for c in ['Project Value','Gross Profit','Taxable Value','Dynamic Price']+COSTS:
        if c not in df:df[c]=np.nan
        df[c]=pd.to_numeric(df[c].astype(str).str.replace(r'[^0-9.\-]','',regex=True),errors='coerce')
    for c in ['KW/PHASE','DCR/NON-DCR','Module','Region']:
        df[c]=df[c].fillna('Unknown').astype(str).str.strip()
    df['DCR/NON-DCR']=df['DCR/NON-DCR'].str.upper().str.replace('NON DCR','NON-DCR',regex=False)
    for c in ['INV DATE- BK','Dispatch Date']:
        if c in df:
            v=df[c]; n=pd.to_numeric(v,errors='coerce')
            dates=pd.to_datetime(v,errors='coerce',dayfirst=True)
            excel=pd.to_datetime(n,unit='D',origin='1899-12-30',errors='coerce')
            df[c]=dates.mask(n.between(20000,80000),excel)
    df['Spending']=df[COSTS].fillna(0).sum(axis=1)
    df['Profit']=df['Gross Profit'].where(df['Gross Profit'].notna(),df['Project Value']-df['Spending'])
    df['Margin %']=np.where(df['Project Value'].ne(0),df['Profit']/df['Project Value']*100,np.nan)
    df['Month']=df['INV DATE- BK'].dt.strftime('%Y-%m').fillna('Unknown') if 'INV DATE- BK' in df else 'Unknown'
    return df

def roll(d,keys):
    t=d.groupby(keys,dropna=False,observed=True).agg(Projects=('Project Value','size'),Revenue=('Project Value','sum'),Spending=('Spending','sum'),Profit=('Profit','sum')).reset_index()
    t['Margin %']=np.where(t['Revenue'].ne(0),100*t['Profit']/t['Revenue'],np.nan)
    t['Profit / Project']=t['Profit']/t['Projects']
    return t

def layout(fig,height=350):
    fig.update_layout(height=height,margin=dict(l=8,r=14,t=50,b=25),paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',font=dict(color=NAVY,size=12),legend=dict(orientation='h',y=1.15,x=0,title=None),xaxis_title=None,yaxis_title=None)
    fig.update_yaxes(gridcolor='#e7edf5',zerolinecolor='#c9d5e4')
    return fig

def chart(t,key,title,columns=('Revenue','Spending','Profit'),height=365):
    tmp=t.melt(id_vars=[key],value_vars=list(columns),var_name='Metric',value_name='Amount')
    tmp['Label']=tmp['Amount'].map(inr)
    fig=px.bar(tmp,x=key,y='Amount',color='Metric',barmode='group',text='Label',title=title,color_discrete_map={'Revenue':BLUE,'Spending':AMBER,'Profit':GREEN})
    fig.update_traces(textposition='outside',textfont_size=10,cliponaxis=False,hovertemplate='%{x}<br>%{fullData.name}: %{text}<extra></extra>')
    fig.update_layout(uniformtext_minsize=9,uniformtext_mode='show',bargap=.25)
    layout(fig,height)
    fig.update_yaxes(tickformat='~s',rangemode='tozero')
    st.plotly_chart(fig,use_container_width=True)

def counts(t,key,title):
    fig=px.bar(t,x=key,y='Projects',text='Projects',title=title,color_discrete_sequence=[BLUE]);fig.update_traces(textposition='outside',cliponaxis=False)
    layout(fig,310);fig.update_yaxes(rangemode='tozero',dtick=1 if t['Projects'].max()<12 else None)
    st.plotly_chart(fig,use_container_width=True)

def table(t):
    view=t.copy()
    for c in ['Revenue','Spending','Profit','Profit / Project']:
        if c in view:view[c]=view[c].map(inr)
    if 'Margin %' in view:view['Margin %']=view['Margin %'].map(lambda v:f'{v:.1f}%' if pd.notna(v) else '—')
    st.dataframe(view,hide_index=True,use_container_width=True,height=min(430,42+len(view)*35))

def insights(t,key):
    if t.empty:return
    low=t.loc[t['Margin %'].idxmin()]; worst=t.loc[t['Profit'].idxmin()]
    neg=t[t['Profit']<0]; total=t['Revenue'].sum(); margin=100*t['Profit'].sum()/total if total else 0
    notes=[f'**Lowest margin:** {low[key]} at **{low["Margin %"]:.1f}%** (profit {inr(low["Profit"])}).',f'**Largest profit drag:** {worst[key]} at **{inr(worst["Profit"])}** profit.']
    if not neg.empty:notes.append(f'**Loss-making groups:** {len(neg)}; combined loss **{inr(neg["Profit"].sum())}**.')
    below=t[t['Margin %']<margin]
    if len(below):notes.append(f'**Below overall margin ({margin:.1f}%):** {len(below)} group(s).')
    st.markdown('**Where we are taking a hit**  ·  '+'  '.join(notes))

st.title('☀️ FIN | Financial Intelligence')
st.caption('Financial story • Revenue • Spending • Profit • Cost pressure')
with st.sidebar:
    upload=st.file_uploader('Upload updated Excel (optional)',type='xlsx')
try:df=load_data(upload.getvalue() if upload else None)
except Exception as ex:st.error(f'Excel error: {ex}');st.stop()
if df is None:st.warning(f'Upload {FILE.name} alongside app.py or use the sidebar uploader.');st.stop()
with st.sidebar:
    st.markdown('### Filters')
    def pick(c,label):
        vals=sorted(df[c].dropna().unique().tolist(),key=str)
        return st.multiselect(label,vals,default=vals)
    months=pick('Month','Invoice month');regions=pick('Region','Region');kws=pick('KW/PHASE','KW / Phase');modules=pick('Module','Module');dcr=pick('DCR/NON-DCR','DCR status')
    st.caption('Revenue = Project Value. Spending = sum of five listed costs. Profit = Gross Profit when present, otherwise Revenue − Spending. Confirm accounting definitions before financial decisions.')
d=df[df['Month'].isin(months)&df['Region'].isin(regions)&df['KW/PHASE'].isin(kws)&df['Module'].isin(modules)&df['DCR/NON-DCR'].isin(dcr)].copy()
if d.empty:st.info('No projects match your filters.');st.stop()
rev=d['Project Value'].sum();spend=d['Spending'].sum();profit=d['Profit'].sum()
a,b,c,e,f=st.columns(5)
a.metric('Projects',f'{len(d):,}');b.metric('Revenue',inr(rev));c.metric('Spending',inr(spend));e.metric('Profit / Savings',inr(profit));f.metric('Profit Margin',f'{profit/rev*100:.1f}%' if rev else '—')
if abs(rev-spend-profit)>max(100,abs(rev)*.005):st.warning(f'Accounting check: Revenue − Spending differs from reported Gross Profit by {inr(rev-spend-profit)}. Profit uses the Excel Gross Profit column when available.')

tabs=st.tabs(['Monthly Story','Region Story','KW-wise','Module-wise','KW × Module','DCR Analysis','Costs & Projects'])
with tabs[0]:
    t=roll(d[d['Month']!='Unknown'],'Month').sort_values('Month')
    if t.empty:st.info('No valid invoice dates available for monthly trends.')
    else:
        counts(t,'Month','1 · Monthly project counts')
        chart(t,'Month','2 · Monthly revenue, spending & profit')
        x,y=st.columns([1.1,1])
        with x:
            fig=px.line(t,x='Month',y='Margin %',markers=True,text=t['Margin %'].map(lambda v:f'{v:.1f}%'),title='3 · Monthly profit margin',color_discrete_sequence=[GREEN]);fig.update_traces(textposition='top center',line_width=3)
            layout(fig,315);st.plotly_chart(fig,use_container_width=True)
        with y:
            st.markdown('#### Month-by-month figures');table(t)
        insights(t,'Month')
        if len(t)>1:
            z=t.copy();z['Revenue Growth %']=z['Revenue'].pct_change()*100;z['Spending Growth %']=z['Spending'].pct_change()*100
            hits=z[(z['Spending Growth %']>z['Revenue Growth %']) & z['Revenue Growth %'].notna()]
            if not hits.empty:st.info('**Spending growing faster than revenue:** '+', '.join(f'{r.Month} ({pct(r["Spending Growth %"])} spend vs {pct(r["Revenue Growth %"])} revenue)' for _,r in hits.iterrows()))
with tabs[1]:
    t=roll(d,'Region').sort_values('Revenue',ascending=False)
    counts(t,'Region','1 · Projects by region');chart(t,'Region','2 · Regional revenue, spending & profit')
    st.markdown('#### 3 · Region profitability');table(t);insights(t,'Region')
with tabs[2]:
    t=roll(d,'KW/PHASE').sort_values('KW/PHASE');chart(t,'KW/PHASE','KW-wise spending & profit',('Spending','Profit'));table(t);insights(t,'KW/PHASE')
with tabs[3]:
    t=roll(d,'Module').sort_values('Revenue',ascending=False);chart(t,'Module','Module-wise spending & profit',('Spending','Profit'));table(t);insights(t,'Module')
with tabs[4]:
    t=roll(d,['KW/PHASE','Module']);pv=t.pivot(index='KW/PHASE',columns='Module',values='Profit').fillna(0)
    if not pv.empty:
        fig=px.imshow(pv,aspect='auto',text_auto=False,color_continuous_scale='RdYlGn',color_continuous_midpoint=0,title='Profit heatmap · KW × Module')
        for i,kw in enumerate(pv.index):
            for j,mod in enumerate(pv.columns):fig.add_annotation(x=mod,y=kw,text=inr(pv.iloc[i,j]),showarrow=False,font=dict(size=11,color=NAVY))
        layout(fig,max(330,65*len(pv)));st.plotly_chart(fig,use_container_width=True)
    table(t);insights(t.assign(Combination=t['KW/PHASE']+' / '+t['Module']),'Combination')
with tabs[5]:
    t=roll(d,'DCR/NON-DCR');chart(t,'DCR/NON-DCR','DCR vs Non-DCR spending & profit',('Spending','Profit'));table(t);insights(t,'DCR/NON-DCR')
    st.markdown('#### DCR × KW × Module');table(roll(d,['DCR/NON-DCR','KW/PHASE','Module']).sort_values('Profit'))
with tabs[6]:
    costs=d[COSTS].sum().sort_values(ascending=False).reset_index();costs.columns=['Cost','Amount'];costs['Label']=costs['Amount'].map(inr)
    fig=px.bar(costs,x='Cost',y='Amount',text='Label',title='Cost drivers · Total spending by type',color_discrete_sequence=[AMBER]);fig.update_traces(textposition='outside',cliponaxis=False);layout(fig);st.plotly_chart(fig,use_container_width=True)
    biggest=costs.iloc[0];st.markdown(f'**Largest cost component:** {biggest["Cost"]} — {inr(biggest["Amount"])} ({biggest["Amount"]/spend*100:.1f}% of total spending).' if spend else 'No spending recorded.')
    loss=d[d['Profit']<0];st.markdown(f'**Loss-making projects:** {len(loss)} / {len(d)} · Total losses: {inr(loss["Profit"].sum())}')
    st.markdown('#### Project-level financial records')
    cols=["RNG's",'LEAD','CUS Name','INV DATE- BK','KW/PHASE','DCR/NON-DCR','Module','Region','Project Value','Spending','Profit','Margin %']+COSTS+['Remarks']
    st.dataframe(d[[c for c in cols if c in d]],hide_index=True,use_container_width=True,height=390)
st.download_button('⬇ Export filtered data',d.to_csv(index=False).encode('utf-8-sig'),'FIN_filtered_data.csv','text/csv')
