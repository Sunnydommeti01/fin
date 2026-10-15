from pathlib import Path
from io import BytesIO
import hashlib
import re
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title='FIN | Financial Intelligence',page_icon='📈',layout='wide')
st.markdown("""<style>
.block-container{max-width:1500px;padding:.55rem 1.2rem 2rem}
.stApp{background:#f6f8fc;color:#182b46}
h1{font-size:1.72rem!important;letter-spacing:-.04em;color:#142b4c;margin:0}
h2,h3{color:#19385d}h4{font-size:1.08rem!important;margin:.3rem 0!important}
[data-testid="stMetric"]{background:white;border:1px solid #e0e9f3;border-radius:13px;padding:10px 14px;box-shadow:0 2px 9px #19385d08}
[data-testid="stMetricValue"]{font-size:1.38rem!important;color:#173d6a}
[data-testid="stVerticalBlock"]{gap:.45rem}
section[data-testid="stSidebar"]{background:#eef4fb}
[data-testid="stExpander"]{background:white;border:1px solid #dce7f2;border-radius:11px;overflow:hidden}
[data-testid="stExpander"] summary{font-weight:600;color:#1b426f}
[data-testid="stTabs"] button{font-size:.88rem}
[data-testid="stPlotlyChart"]{background:white;border:1px solid #e5ecf5;border-radius:12px;padding:3px 4px}
.fin-note{background:#fff;border:1px solid #e0e8f1;border-left:3px solid #b45309;border-radius:9px;padding:10px 14px;color:#3a4b60;font-size:.88rem}
[data-testid="stDataFrame"]{border:1px solid #e4ebf4;border-radius:10px}
hr{margin:.45rem 0}
</style>""",unsafe_allow_html=True)
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

@st.cache_data(ttl=600,show_spinner=False,max_entries=4)
def read_excel(raw):
    df=pd.read_excel(BytesIO(raw),engine='openpyxl')
    df.columns=df.columns.astype(str).str.strip()
    df=df.dropna(how='all').reset_index(drop=True)
    if df.columns.duplicated().any():
        raise ValueError('Duplicate Excel column names: '+', '.join(df.columns[df.columns.duplicated()]))
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
    df['Month']=df['Month'].astype(str)
    # One Excel row = one project. Do not silently drop repeated lead IDs.
    df['Spending']=df[COSTS].fillna(0).sum(axis=1)
    df['Missing Cost Fields']=df[COSTS].isna().sum(axis=1)
    df['Profit']=df['Gross Profit'].where(df['Gross Profit'].notna(),df['Project Value']-df['Spending'])
    df['Material Cost']=df['Material Cost'].fillna(0)
    df['Cost Complete']=df[COSTS].notna().all(axis=1)
    df['Profit Source']=np.where(df['Gross Profit'].notna(),'Excel Gross Profit','Revenue less listed costs')
    return df

def summary(df,keys):
    t=df.groupby(keys,dropna=False,observed=True).agg(Projects=('Project Value','size'),Revenue=('Project Value','sum'),Spending=('Spending','sum'),Profit=('Profit','sum'),**{'Material Cost':('Material Cost','sum')}).reset_index()
    t['Margin %']=np.where(t.Revenue.ne(0),100*t.Profit/t.Revenue,np.nan)
    t['Avg Spending']=t.Spending/t.Projects
    t['Avg Profit']=t.Profit/t.Projects
    return t

def month_label(value):
    """Display an actual calendar month without allowing Plotly date ticks."""
    try:
        return pd.Period(str(value), freq='M').strftime('%b %Y')
    except (ValueError, TypeError):
        return str(value)


def category_axis(fig, labels, axis='x', monthly=False):
    """Force exact discrete categories, never a continuous date scale."""
    values = [str(v) for v in labels]
    display = [month_label(v) for v in values] if monthly else values
    opts = dict(type='category', categoryorder='array', categoryarray=values,
                tickmode='array', tickvals=values, ticktext=display)
    if axis == 'x':
        fig.update_xaxes(**opts)
    else:
        fig.update_yaxes(**opts)


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
    category_axis(fig, t[x].tolist(), monthly=(x == 'Month'))
    if len(t)>16:
        fig.update_traces(textposition='none')
        st.caption('Labels hidden for crowded charts; use the detailed table below for exact values.')
    fig.update_yaxes(rangemode='tozero',tickformat='.2s')
    st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})

def heat(t,rows,cols,metric,title):
    st.markdown(f'**{title}**')
    if t.empty:st.info('No matching data');return
    if metric=='Margin %':
        revenue=t.pivot_table(index=rows,columns=cols,values='Revenue',aggfunc='sum',observed=True)
        profit=t.pivot_table(index=rows,columns=cols,values='Profit',aggfunc='sum',observed=True)
        p=(profit/revenue.replace(0,np.nan))*100
    elif metric=='Avg Spending':
        spending=t.pivot_table(index=rows,columns=cols,values='Spending',aggfunc='sum',observed=True)
        projects=t.pivot_table(index=rows,columns=cols,values='Projects',aggfunc='sum',observed=True)
        p=spending/projects.replace(0,np.nan)
    else:
        p=t.pivot_table(index=rows,columns=cols,values=metric,aggfunc='sum',observed=True)
    if p.empty:return
    if cols == 'Month': p=p.reindex(sorted(p.columns), axis=1)
    if rows == 'Month': p=p.reindex(sorted(p.index), axis=0)
    zz=p.to_numpy(dtype=float)
    txt=[[lbl(v,metric) for v in line] for line in zz]
    show_text=(len(p.index)*len(p.columns)<=160)
    fig=go.Figure(go.Heatmap(z=zz,x=[str(v) for v in p.columns],y=[str(v) for v in p.index],text=txt,texttemplate='%{text}' if show_text else None,textfont=dict(size=10),colorscale='RdYlGn' if metric in ['Profit','Margin %'] else 'Blues',colorbar=dict(title='%',len=.7) if metric=='Margin %' else dict(len=.7),hovertemplate='%{y} × %{x}<br>%{text}<extra></extra>'))
    fig_style(fig,min(650,max(280,90+len(p)*30)))
    category_axis(fig, p.columns.tolist(), axis='x', monthly=(cols == 'Month'))
    category_axis(fig, p.index.tolist(), axis='y', monthly=(rows == 'Month'))
    if not show_text:st.caption('Dense heatmap: hover for exact amounts, or open Interactive data details for filterable values.')
    st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})

def pivot_table(t, key, height=390, export_name=None):
    """Excel-style searchable checklist filters in each column menu."""
    st.caption('Filter any column using its ☰ menu: search, Select All, Apply or Clear. Sort by clicking a header.')
    try:
        from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
    except ImportError:
        st.error('Install streamlit-aggrid in requirements.txt to enable Excel-style filters.')
        st.dataframe(t, use_container_width=True, hide_index=True)
        return t

    view = t.copy().reset_index(drop=True)
    # AG Grid can have difficulty with pandas datetime objects; display dates as text.
    for column in view.select_dtypes(include=['datetime64[ns]', 'datetimetz']).columns:
        view[column] = view[column].dt.strftime('%d %b %Y').fillna('')
    view = view.replace([np.inf, -np.inf], np.nan)
    for column in view.columns:
        if pd.api.types.is_object_dtype(view[column]):
            view[column] = view[column].fillna('').astype(str)

    gb = GridOptionsBuilder.from_dataframe(view)
    gb.configure_default_column(
        sortable=True, filter=True, resizable=True,
        floatingFilter=False, minWidth=135,
        suppressMenu=False,
    )
    for column in view.columns:
        if pd.api.types.is_numeric_dtype(view[column]):
            gb.configure_column(column, filter='agNumberColumnFilter',
                filterParams={'buttons':['apply','clear'], 'closeOnApply':True})
        else:
            gb.configure_column(column, filter='agSetColumnFilter',
                filterParams={
                    'buttons':['apply','clear'],
                    'closeOnApply':True,
                    'excelMode':'windows',
                    'suppressSelectAll':False,
                    'suppressMiniFilter':False,
                })
    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=25)
    gb.configure_grid_options(
        animateRows=False, enableCellTextSelection=True, ensureDomOrder=True,
        suppressRowClickSelection=True,
        sideBar={'toolPanels':[
            {'id':'filters','labelDefault':'Filters','labelKey':'filters',
             'iconKey':'filter','toolPanel':'agFiltersToolPanel'},
            {'id':'columns','labelDefault':'Columns','labelKey':'columns',
             'iconKey':'columns','toolPanel':'agColumnsToolPanel'},
        ], 'defaultToolPanel':''},
    )
    response = AgGrid(
        view, gridOptions=gb.build(), height=height,
        theme='streamlit', fit_columns_on_grid_load=False,
        update_mode=GridUpdateMode.FILTERING_CHANGED,
        data_return_mode=DataReturnMode.FILTERED,
        enable_enterprise_modules=True,
        allow_unsafe_jscode=False, key=key,
    )
    filtered = response.get('data', view)
    if filtered is None:
        filtered = view
    if not isinstance(filtered, pd.DataFrame):
        filtered = pd.DataFrame(filtered)
    if export_name:
        st.download_button('⬇ Download filtered rows (CSV)',
            filtered.to_csv(index=False).encode('utf-8-sig'),
            export_name, 'text/csv', key=f'export_{key}')
    return filtered

def insight(t,dim):
    if t.empty:return
    worst=t.loc[t['Margin %'].idxmin()]
    losses=t[t.Profit<0]
    st.caption(f'⚠️ Lowest margin: {worst[dim]} ({lbl(worst["Margin %"],"Margin %")}); profit {money(worst.Profit)}.'+(f' {len(losses)} loss-making group(s).' if len(losses) else ''))


# ===============================================================
# FIN V7 — ONE CLICK PROJECT DETAIL FOR EVERY REPORT
# ===============================================================
PROJECT_COLUMNS=["RNG's",'LEAD','CUS Name','INV DATE- BK','Dispatch Date','Month','KW/PHASE','DCR/NON-DCR','Module','Region','Project Value','Taxable Value','Dynamic Price','Material Cost','Transportation Cost','Installation Cost','Net meter cost','Vendor Margin','Spending','Profit','Margin %','Gross Profit','Remarks']


def project_view(frame):
    x=frame.copy()
    x['Margin %']=np.where(x['Project Value'].ne(0),x['Profit']/x['Project Value']*100,np.nan)
    return x[[c for c in PROJECT_COLUMNS if c in x.columns]].reset_index(drop=True)


def detail_box(report_key, frame, group_dims, description=''):
    """Independent expandable details and Excel-like column filters per report."""
    with st.expander('▸ Interactive data details', expanded=False):
        if description:st.caption(description)
        search=st.text_input('Search projects',placeholder='Search lead, customer, module, region, amount...',key=f'search_{report_key}')
        subset=frame
        if search.strip():
            q=search.strip().casefold()
            fields=[c for c in ['LEAD','CUS Name','Region','Module','KW/PHASE','DCR/NON-DCR','Month',"RNG's",'Remarks'] if c in frame]
            if fields:
                mask=pd.Series(False,index=frame.index)
                for c in fields:mask=mask|frame[c].fillna('').astype(str).str.casefold().str.contains(q,regex=False)
                subset=frame.loc[mask]
        tab_summary,tab_projects=st.tabs(['Summary / Pivot','Projects'])
        with tab_summary:
            if group_dims:
                result=summary(subset,group_dims).sort_values(group_dims)
                pivot_table(result,key=f'v7_{report_key}_summary',height=300,export_name=f'FIN_{report_key}_summary.csv')
            else:
                result=project_view(subset)
                pivot_table(result,key=f'v7_{report_key}_summary',height=300,export_name=f'FIN_{report_key}_summary.csv')
        with tab_projects:
            st.caption(f'{len(subset):,} matching projects. Click the filter menu in any column header to filter like Excel.')
            pivot_table(project_view(subset),key=f'v7_{report_key}_projects',height=370,export_name=f'FIN_{report_key}_projects.csv')


def section_title(n,title,subtitle):
    st.markdown(f'#### {n:02d} · {title}')
    st.caption(subtitle)

st.title('FIN  /  Financial Intelligence')
st.caption('Solar project profitability  •  Month → Region → KW → Module → DCR → Top 5  •  ₹ Indian format')
with st.sidebar:
    st.markdown('### FIN · Data source')
    up=st.file_uploader('Upload replacement Excel (optional)',type=['xlsx'])
    files=sorted((f for f in ROOT.glob('*.xlsx') if not f.name.startswith('~$')),key=lambda f:f.name)
    preferred=next((f for f in files if f.name.lower() in ('fin_data.xlsx','financial_data.xlsx')),None)
    default_index=files.index(preferred) if preferred else 0
    path=st.selectbox('Excel file in repository',files,index=default_index,format_func=lambda f:f.name) if files and up is None else None
    st.caption('Source: '+('Uploaded Excel' if up else path.name if path else 'No Excel file'))
if not up and not path:
    st.error('Place your Excel file beside app.py, or upload it using the sidebar.');st.stop()
try:df=read_excel(up.getvalue() if up else path.read_bytes())
except Exception as exc:st.error(f'Excel loading error: {exc}');st.stop()
with st.sidebar:
    st.markdown('### Global filters')
    picks={}
    for dim in DIM:
        values=sorted(df[dim].dropna().unique().tolist(),key=str)
        picks[dim]=st.multiselect(dim,values,default=values,key='global_'+dim)
    st.caption('Calendar months are grouped by invoice date (INV DATE- BK). Revenue = Project Value; Profit = Gross Profit where present.')
    st.caption(f'Source rows: {len(df):,} · Months: {df.Month.nunique()}')
d=df.copy()
for dim in DIM:d=d[d[dim].isin(picks[dim])]
if d.empty:st.warning('No records match the filters.');st.stop()
if len(d)>0 and d['INV DATE- BK'].isna().any():
    st.caption(f'ℹ️ {int(d["INV DATE- BK"].isna().sum()):,} rows have missing/unreadable invoice dates; they appear as Unknown Month and are excluded from monthly charts.')
rev=d['Project Value'].sum();sp=d['Spending'].sum();pr=d['Profit'].sum()
cols=st.columns(5)
for c,name,val in zip(cols,['Projects','Revenue','Spending','Gross profit','Profit margin'],[f'{len(d):,}',money(rev),money(sp),money(pr),f'{pr/rev*100:.1f}%' if rev else '—']):c.metric(name,val)
if d['Missing Cost Fields'].gt(0).any():
    st.caption(f'⚠️ {int(d["Missing Cost Fields"].gt(0).sum()):,} projects have missing cost fields. Missing values are provisionally treated as ₹0 in Spending.')
if abs(rev-sp-pr)>max(100,abs(rev)*.005):
    st.caption(f'⚠️ Reconciliation: Project Value − listed spending differs from Gross Profit by {money(rev-sp-pr)}. Check the Excel profit definition.')

st.markdown('### Performance story')
OPTIONS=['All 11 reports','01 · Month-wise projects & revenue','02 · Region-wise projects & revenue','03 · Financial trend','04 · Monthly profit margin','05 · Month × Region','06 · KW-wise revenue, spending & profit','07 · KW × Month','08 · KW × Region','09 · Profit heatmap','10 · DCR × Region × Month','11 · Top 5 Intelligence','Custom Pivot Builder']
nav=st.selectbox('Jump to report',OPTIONS,index=1,label_visibility='collapsed')
st.caption('Select All 11 reports to view the full story; one report at a time is faster for larger datasets.')
def show(n):return nav=='All 11 reports' or nav.startswith(f'{n:02d} ·')
month=summary(d[d.Month!='Unknown'],['Month']).sort_values('Month')
region=summary(d,['Region']).sort_values('Revenue',ascending=False)
kw=summary(d,['KW/PHASE']).sort_values('KW/PHASE')

if show(1):
    section_title(1,'Month-wise projects & revenue','Monthly project volume and total project value')
    a,b=st.columns(2)
    with a:bars(month,'Month',['Projects'],'Projects by month',height=275)
    with b:bars(month,'Month',['Revenue'],'Revenue by month',height=275)
    detail_box('monthly',d,['Month'],'Monthly summary and source projects')
if show(2):
    section_title(2,'Region-wise projects & revenue','Which regions contribute most to the business?')
    a,b=st.columns(2)
    with a:bars(region,'Region',['Projects'],'Projects by region',height=285)
    with b:bars(region,'Region',['Revenue'],'Revenue by region',height=285)
    detail_box('regional',d,['Region'])
if show(3):
    section_title(3,'Monthly financial trend','Revenue, spending and gross profit · plus a compact monthly loss tracker')
    bars(month,'Month',['Revenue','Spending','Profit'],'Financial trend',height=315)

    # FIN V8: compact, immediately visible pivot beneath Graph 3.
    # Loss = a project whose recorded/calculated Profit is strictly below zero.
    monthly_source=d[d['Month']!='Unknown'].copy()
    monthly_source['Loss Flag']=(monthly_source['Profit']<0).astype(int)
    monthly_source['Loss Amount']=monthly_source['Profit'].clip(upper=0).abs()
    monthly_pivot=monthly_source.groupby('Month',observed=True).agg(
        Projects=('Profit','size'),Revenue=('Project Value','sum'),
        Spending=('Spending','sum'),Profit=('Profit','sum'),
        **{'Loss Projects':('Loss Flag','sum'),'Loss Amount':('Loss Amount','sum')}
    ).reset_index().sort_values('Month')
    monthly_pivot['Margin %']=np.where(monthly_pivot['Revenue'].ne(0),
        monthly_pivot['Profit']/monthly_pivot['Revenue']*100,np.nan)
    st.markdown('**Monthly financial snapshot & losses**')
    st.caption('Negative-profit projects are counted separately, even when the overall month is profitable. Use the column menus for Excel-style filters.')
    if not monthly_pivot.empty:
        loss_count=int(monthly_pivot['Loss Projects'].sum())
        loss_total=float(monthly_pivot['Loss Amount'].sum())
        a,b,c=st.columns(3)
        a.metric('Loss-making projects',f'{loss_count:,}')
        b.metric('Total losses',money(loss_total))
        c.metric('Months with losses',f'{int((monthly_pivot["Loss Projects"]>0).sum())}')
        selected=st.selectbox('Quick pivot view',
            ['Full financial snapshot','Revenue · Spending · Profit','Loss analysis'],
            key='v8_monthly_pivot_view',label_visibility='collapsed')
        cols_by_view={
            'Full financial snapshot':['Month','Projects','Revenue','Spending','Profit','Margin %','Loss Projects','Loss Amount'],
            'Revenue · Spending · Profit':['Month','Projects','Revenue','Spending','Profit','Margin %'],
            'Loss analysis':['Month','Projects','Profit','Margin %','Loss Projects','Loss Amount']
        }
        pivot_table(monthly_pivot[cols_by_view[selected]],
            key='v8_graph3_mini_pivot',height=255,
            export_name='FIN_monthly_financial_losses.csv')
        if loss_count:
            worst=monthly_pivot.loc[monthly_pivot['Loss Amount'].idxmax()]
            st.markdown('<div class="fin-note">⚠️ Largest monthly loss concentration: '
                +str(worst['Month'])+' · '+money(worst['Loss Amount'])+
                ' across '+str(int(worst['Loss Projects']))+' loss-making project(s).</div>',
                unsafe_allow_html=True)
        else:
            st.caption('No negative-profit projects in the selected period.')
    detail_box('trend',d,['Month'])
if show(4):
    section_title(4,'Monthly profit margin','Weighted margin = sum of gross profit / sum of project value')
    bars(month,'Month',['Margin %'],'Gross profit margin (%)',kind='line',height=290)
    insight(month,'Month')
    detail_box('margin',d,['Month'])
if show(5):
    section_title(5,'Month × Region','Find regional profit and margin pressure month by month')
    mr=summary(d[d.Month!='Unknown'],['Month','Region'])
    met=st.selectbox('Heatmap measure',['Profit','Margin %','Revenue','Spending','Projects','Material Cost','Avg Spending'],key='mr_metric')
    heat(mr,'Region','Month',met,'Month × Region · '+met)
    detail_box('month_region',d[d.Month!='Unknown'],['Month','Region'])
if show(6):
    section_title(6,'KW-wise revenue, spending & profit','Compare system capacity by revenue and cost efficiency')
    bars(kw,'KW/PHASE',['Revenue','Spending','Profit'],'KW profitability',height=330)
    insight(kw,'KW/PHASE')
    detail_box('kw',d,['KW/PHASE'])
if show(7):
    section_title(7,'KW × Month','Track capacity-wise spending and profit over time')
    km=summary(d[d.Month!='Unknown'],['KW/PHASE','Month'])
    met=st.radio('Measure',['Spending','Profit','Revenue','Margin %'],horizontal=True,key='km_met')
    heat(km,'KW/PHASE','Month',met,'KW × Month · '+met)
    detail_box('kw_month',d[d.Month!='Unknown'],['KW/PHASE','Month'])
if show(8):
    section_title(8,'KW × Region','Spot expensive capacities in individual regions')
    kr=summary(d,['KW/PHASE','Region'])
    met=st.radio('Measure',['Spending','Profit','Revenue','Margin %'],horizontal=True,key='kr_met')
    heat(kr,'KW/PHASE','Region',met,'KW × Region · '+met)
    detail_box('kw_region',d,['KW/PHASE','Region'])
if show(9):
    section_title(9,'Profit heatmap · KW × Module × Region × Month','Choose two chart axes; use sidebar filters to isolate the remaining dimensions')
    allowed=['KW/PHASE','Module','Region','Month']
    a,b,c=st.columns(3)
    with a:row=st.selectbox('Heatmap rows',allowed,index=0,key='hrow')
    remaining=[x for x in allowed if x!=row]
    with b:col=st.selectbox('Heatmap columns',remaining,index=0,key='hcol')
    with c:met=st.selectbox('Measure',['Profit','Margin %','Spending','Revenue','Projects'],key='hmet')
    base=d[d.Month!='Unknown'] if 'Month' in [row,col] else d
    heat(summary(base,[row,col]),row,col,met,f'{row} × {col} · {met}')
    detail_box('profit_heatmap',base,['KW/PHASE','Module','Region','Month'])
if show(10):
    section_title(10,'DCR vs Non-DCR · Region × Month','Compare DCR categories across geography and time')
    a,b=st.columns(2)
    with a:met=st.selectbox('Measure',['Profit','Spending','Revenue','Projects','Margin %'],key='dcr_metric')
    with b:axis=st.selectbox('Compare against',['Region','Month'],key='dcr_axis')
    base=d[d.Month!='Unknown'] if axis=='Month' else d
    heat(summary(base,['DCR/NON-DCR',axis]),'DCR/NON-DCR',axis,met,f'DCR status × {axis} · {met}')
    detail_box('dcr',d[d.Month!='Unknown'],['DCR/NON-DCR','Region','Month'])


# ===============================================================
# 11 · TOP 5 INTELLIGENCE — RANKING + FULL PROJECT AUDIT
# ===============================================================
if show(11):
    section_title(11, 'Top 5 · Profit & Loss Intelligence',
        'Rank the best and worst projects, costs and business segments. Every ranking opens its full underlying project data.')
    rank_options = {
        'Most profitable projects': ('project', 'Profit', False),
        'Biggest loss-making projects': ('loss', 'Profit', True),
        'Highest revenue projects': ('project', 'Project Value', False),
        'Highest spending projects': ('project', 'Spending', False),
        'Lowest-margin projects': ('project', 'Margin %', True),
        'Highest-margin projects': ('project', 'Margin %', False),
        'Highest material cost projects': ('project', 'Material Cost', False),
        'Highest transportation cost projects': ('project', 'Transportation Cost', False),
        'Highest installation cost projects': ('project', 'Installation Cost', False),
        'Highest net-meter cost projects': ('project', 'Net meter cost', False),
        'Highest vendor margin projects': ('project', 'Vendor Margin', False),
        'Most profitable regions': ('Region', 'Profit', False),
        'Least profitable regions': ('Region', 'Profit', True),
        'Most profitable months': ('Month', 'Profit', False),
        'Least profitable months': ('Month', 'Profit', True),
        'Most profitable KW capacities': ('KW/PHASE', 'Profit', False),
        'Least profitable KW capacities': ('KW/PHASE', 'Profit', True),
        'Most profitable modules': ('Module', 'Profit', False),
        'Least profitable modules': ('Module', 'Profit', True),
        'Most profitable DCR categories': ('DCR/NON-DCR', 'Profit', False),
        'Least profitable DCR categories': ('DCR/NON-DCR', 'Profit', True),
        'Highest-spending regions': ('Region', 'Spending', False),
        'Highest-spending months': ('Month', 'Spending', False),
        'Highest-spending KW capacities': ('KW/PHASE', 'Spending', False),
        'Highest-spending modules': ('Module', 'Spending', False),
        'Best KW × Module combinations': ('KW/PHASE|Module', 'Profit', False),
        'Worst KW × Module combinations': ('KW/PHASE|Module', 'Profit', True),
        'Best Region × Month combinations': ('Region|Month', 'Profit', False),
        'Worst Region × Month combinations': ('Region|Month', 'Profit', True),
    }
    a,b,c=st.columns([2.5,1,1])
    with a: rank_name=st.selectbox('Choose your Top 5 ranking',list(rank_options),key='v9_rank_choice')
    with b: rank_size=st.selectbox('Show', [5,10,20],index=0,key='v9_rank_size')
    with c: chart_measure=st.selectbox('Chart', ['Ranked bars','Table only'],key='v9_rank_chart')
    category, metric, ascending=rank_options[rank_name]
    rank_source=d.copy()
    rank_source['Margin %']=np.where(rank_source['Project Value'].ne(0),
        100*rank_source['Profit']/rank_source['Project Value'],np.nan)
    if category in ('project','loss'):
        if category=='loss':
            rank_source=rank_source[rank_source['Profit']<0].copy()
        rank_source=rank_source.dropna(subset=[metric]).sort_values(metric,ascending=ascending,kind='stable').head(rank_size)
        rank_source['Rank']=range(1,len(rank_source)+1)
        display_cols=['Rank','LEAD','CUS Name','Month','Region','KW/PHASE','Module','DCR/NON-DCR',
            'Project Value','Spending','Profit','Margin %','Material Cost','Transportation Cost',
            'Installation Cost','Net meter cost','Vendor Margin','Remarks']
        rank_table=rank_source[[x for x in display_cols if x in rank_source]].copy()
        chart_labels=(rank_source['LEAD'].fillna('').astype(str) if 'LEAD' in rank_source
            else rank_source.index.astype(str))
        chart_labels=chart_labels.where(chart_labels.str.strip().ne(''),rank_source.index.astype(str))
        chart_labels=chart_labels+' · '+rank_source['Region'].astype(str)
    else:
        group_dims=category.split('|')
        rank_all=summary(rank_source,group_dims)
        if metric=='Margin %':
            rank_all=rank_all.dropna(subset=['Margin %'])
        rank_all=rank_all.sort_values(metric,ascending=ascending,kind='stable').head(rank_size).copy()
        rank_all.insert(0,'Rank',range(1,len(rank_all)+1))
        rank_table=rank_all
        chart_labels=rank_all[group_dims].astype(str).agg(' · '.join,axis=1)
    if rank_table.empty:
        st.info('No matching projects for this ranking and the selected filters.')
    else:
        k1,k2,k3=st.columns(3)
        k1.metric('Ranked entries',len(rank_table))
        if category in ('project','loss'):
            k2.metric('Ranked profit',money(rank_source['Profit'].sum()))
            k3.metric('Ranked spending',money(rank_source['Spending'].sum()))
        else:
            k2.metric('Ranked profit',money(rank_table['Profit'].sum()))
            k3.metric('Ranked revenue',money(rank_table['Revenue'].sum()))
        if chart_measure=='Ranked bars':
            chart_vals=rank_source[metric] if category in ('project','loss') else rank_table[metric]
            fig=go.Figure(go.Bar(
                x=chart_vals.iloc[::-1],y=chart_labels.iloc[::-1],orientation='h',
                marker_color=['#dc4960' if v<0 else '#16826b' for v in chart_vals.iloc[::-1]],
                text=[lbl(v,metric) for v in chart_vals.iloc[::-1]],
                textposition='outside',cliponaxis=False,textfont=dict(size=10)))
            fig_style(fig,min(570,max(250,120+len(rank_table)*39)))
            fig.update_layout(margin=dict(l=8,r=85,t=12,b=25),showlegend=False)
            fig.update_yaxes(automargin=True,gridcolor='rgba(0,0,0,0)')
            fig.update_xaxes(tickformat='.2s',zeroline=True,zerolinecolor='#9ca3af')
            category_axis(fig, chart_labels.iloc[::-1].tolist(), axis='y')
            st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})
        st.markdown('**Top ranking · Excel-filterable financial details**')
        pivot_table(rank_table,'v9_rank_table',height=290,export_name='FIN_top_rankings.csv')
        # Always show exact contributing raw projects, not just grouped totals.
        with st.expander('▸ View full raw projects behind this ranking',expanded=False):
            if category in ('project','loss'):
                underlying=rank_source
            else:
                keys=rank_table[group_dims].drop_duplicates()
                underlying=rank_source.merge(keys,on=group_dims,how='inner')
            st.caption(f'{len(underlying):,} underlying project rows · includes all available original Excel columns')
            raw_cols=[x for x in PROJECT_COLUMNS if x in underlying]
            other_cols=[x for x in underlying.columns if x not in raw_cols and x not in ['Cost Complete']]
            pivot_table(underlying[raw_cols+other_cols].reset_index(drop=True),
                'v9_rank_raw',height=400,export_name='FIN_top_rankings_raw_projects.csv')
    detail_box('top5_all',d,['Month','Region','KW/PHASE','Module'],
        'Full filtered project population for independent investigation and pivoting.')

if nav in ('All 11 reports','Custom Pivot Builder'):
    with st.expander('▸ Expand custom pivot builder', expanded=False):
        st.markdown('#### Custom Pivot Builder')
        st.caption('Choose any dimensions and metrics. Filter each pivot column like Excel.')
        a,b=st.columns(2)
        with a:dims=st.multiselect('Group by',DIM,default=['Month','Region'],key='pivot_dims')
        with b:metrics=st.multiselect('Measures',['Projects','Revenue','Spending','Profit','Margin %','Material Cost','Avg Spending','Avg Profit'],default=['Projects','Revenue','Spending','Profit','Margin %'],key='pivot_metrics')
        if dims and metrics:
            t=summary(d,dims)[dims+metrics].sort_values(dims)
            pivot_table(t,'v7_custom_pivot',height=380,export_name='FIN_custom_pivot.csv')
        else:st.info('Choose at least one dimension and one measure.')
st.caption('FIN V11 · Monthly chart axes use discrete calendar-month categories. Financial note: Each Excel row counts as one project. Missing cost components count as zero in spending; validate Gross Profit and Vendor Margin accounting before relying on profitability rankings.')
