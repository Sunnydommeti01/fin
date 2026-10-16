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
NUMS=['Project Value','Taxable Value','Outward GST','Dynamic Price','Gross Profit','Actual GP Percentage']+COSTS
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

def normalize_header(value):
    """Canonicalize Excel headers including NBSP, zero-width characters and line breaks."""
    import unicodedata
    value = unicodedata.normalize('NFKC', str(value))
    value = re.sub(r'[\u200b-\u200d\ufeff]', '', value)
    value = re.sub(r'\s+', ' ', value).strip()
    return value


def header_key(value):
    return re.sub(r'[^a-z0-9]', '', normalize_header(value).casefold())


CANONICAL_COLUMNS = [
    "RNG's", 'LEAD', 'CUS Name', 'INV DATE- BK', 'KW/PHASE',
    'DCR/NON-DCR', 'Module', 'Project Value', 'Taxable Value',
    'Outward GST', 'Dynamic Price', 'Vendor Margin', 'Region',
    'Material Cost', 'Dispatch Date', 'Transportation Cost',
    'Installation Cost', 'Net meter cost', 'Gross Profit',
    'Actual GP Percentage', 'Remarks'
]
COLUMN_ALIASES = {header_key(name): name for name in CANONICAL_COLUMNS}
COLUMN_ALIASES.update({
    'invdatebk': 'INV DATE- BK', 'invoicedate': 'INV DATE- BK',
    'kwphase': 'KW/PHASE', 'dcrnondcr': 'DCR/NON-DCR',
    'netmeteringcost': 'Net meter cost',
    'grossprofitpercentage': 'Actual GP Percentage',
    'actualgppercent': 'Actual GP Percentage',
})
REQUIRED_COLUMNS = ['Project Value','KW/PHASE','Module','DCR/NON-DCR','Region','INV DATE- BK']


def find_data_sheet(raw):
    """Search every sheet's first 25 rows for the real table header."""
    workbook = pd.ExcelFile(BytesIO(raw), engine='openpyxl')
    best = None
    for sheet in workbook.sheet_names:
        preview = pd.read_excel(workbook, sheet_name=sheet, header=None, nrows=25, dtype=object)
        for row_idx, row in preview.iterrows():
            keys = {COLUMN_ALIASES.get(header_key(v)) for v in row if pd.notna(v)}
            matched = len(keys.intersection(REQUIRED_COLUMNS))
            if best is None or matched > best[0]:
                best = (matched, sheet, int(row_idx))
            if matched == len(REQUIRED_COLUMNS):
                return workbook, sheet, int(row_idx)
    if best is None or best[0] < 4:
        raise ValueError('Could not locate the financial data header in the first 25 rows of any worksheet. Check that the selected Excel file contains your project table.')
    return workbook, best[1], best[2]


@st.cache_data(ttl=600,show_spinner=False,max_entries=4)
def read_excel(raw):
    workbook, sheet, header_row = find_data_sheet(raw)
    df = pd.read_excel(workbook, sheet_name=sheet, header=header_row, engine='openpyxl')
    df.columns = [COLUMN_ALIASES.get(header_key(c), normalize_header(c)) for c in df.columns]
    if df.columns.duplicated().any():
        duplicated = df.columns[df.columns.duplicated()].tolist()
        raise ValueError('Multiple Excel columns map to the same name: ' + ', '.join(duplicated))
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f'Sheet {sheet!r}, header row {header_row+1}: missing columns: {missing}. Detected headers: {df.columns.tolist()}')
    df = df.dropna(how='all').reset_index(drop=True)
    # Exclude footer and total rows, while retaining records that have a lead or financial data.
    df = df.loc[df[['INV DATE- BK','Project Value','LEAD'] if 'LEAD' in df else ['INV DATE- BK','Project Value']].notna().any(axis=1)].copy()
    for c in NUMS:
        if c not in df.columns:
            df[c] = np.nan
        series = df[c].astype(str).str.replace(',', '', regex=False).str.replace(r'[^0-9.\-]', '', regex=True)
        df[c] = pd.to_numeric(series, errors='coerce')
    for c in ['KW/PHASE','Module','DCR/NON-DCR','Region']:
        df[c] = df[c].fillna('Unknown').astype(str).str.strip().replace('', 'Unknown')
    df['DCR/NON-DCR'] = df['DCR/NON-DCR'].str.upper().str.replace('NON DCR','NON-DCR',regex=False)
    df['INV DATE- BK'] = parse_dates(df['INV DATE- BK'])
    df['Month'] = df['INV DATE- BK'].dt.strftime('%Y-%m').fillna('Unknown').astype(str)
    df['Spending'] = df[COSTS].fillna(0).sum(axis=1)
    df['Missing Cost Fields'] = df[COSTS].isna().sum(axis=1)
    # Profit is calculated on the taxable (ex-GST) revenue base.
    # Preserve Excel GP when present; fallback only when taxable value is available.
    df['Calculated GP'] = df['Taxable Value'] - df['Spending']
    df['Profit'] = df['Gross Profit'].where(df['Gross Profit'].notna(), df['Calculated GP'])
    df['GP Difference'] = df['Calculated GP'] - df['Gross Profit']
    df['Margin %'] = np.where(df['Project Value'].ne(0) & df['Project Value'].notna(),
                              100 * df['Profit'] / df['Project Value'], np.nan)
    df['Cost Complete'] = df[COSTS].notna().all(axis=1)
    df['Material Cost'] = df['Material Cost'].fillna(0)
    df['Profit Source'] = np.where(df['Gross Profit'].notna(), 'Excel Gross Profit', 'Taxable Value less listed costs')
    df.attrs['source_sheet'] = sheet
    df.attrs['header_row'] = header_row + 1
    return df

def summary(df,keys):
    t=df.groupby(keys,dropna=False,observed=True).agg(Projects=('Project Value','size'),**{'Project Value':('Project Value','sum')},Revenue=('Project Value','sum'),Spending=('Spending','sum'),Profit=('Profit','sum'),**{'Material Cost':('Material Cost','sum')}).reset_index()
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
PROJECT_COLUMNS=["RNG's",'LEAD','CUS Name','INV DATE- BK','Dispatch Date','Month','KW/PHASE','DCR/NON-DCR','Module','Region','Project Value','Taxable Value','Outward GST','Dynamic Price','Actual GP Percentage','Material Cost','Transportation Cost','Installation Cost','Net meter cost','Vendor Margin','Spending','Calculated GP','Profit','GP Difference','Margin %','Gross Profit','Profit Source','Missing Cost Fields','Remarks']


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

overview_tab,predictive_tab=st.tabs(['📊 Overview of Data','🔮 Predictive Analysis'])
with overview_tab:
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
        st.caption(f"Detected worksheet: {df.attrs.get('source_sheet', 'Unknown')} · Header row: {df.attrs.get('header_row', '?')}")
        st.markdown('### Global filters')
        picks={}
        for dim in DIM:
            values=sorted(df[dim].dropna().unique().tolist(),key=str)
            picks[dim]=st.multiselect(dim,values,default=values,key='global_'+dim)
        st.caption('Calendar months are grouped by invoice date (INV DATE- BK). Revenue = Project Value; GP margin = Profit / Project Value. Accounting GP reconciliation still uses Taxable Value minus costs.')
        st.caption(f'Source rows: {len(df):,} · Months: {df.Month.nunique()}')
    d=df.copy()
    for dim in DIM:d=d[d[dim].isin(picks[dim])]
    if d.empty:st.warning('No records match the filters.');st.stop()
    if len(d)>0 and d['INV DATE- BK'].isna().any():
        st.caption(f'ℹ️ {int(d["INV DATE- BK"].isna().sum()):,} rows have missing/unreadable invoice dates; they appear as Unknown Month and are excluded from monthly charts.')
    rev=d['Project Value'].sum();taxable_revenue=d['Taxable Value'].sum();sp=d['Spending'].sum();pr=d['Profit'].sum()
    cols=st.columns(6)
    for c,name,val in zip(cols,['Projects','Revenue · Project Value','Taxable Value','Spending','Gross profit','GP margin'],[f'{len(d):,}',money(rev),money(taxable_revenue),money(sp),money(pr),f'{pr/rev*100:.1f}%' if rev else '—']):c.metric(name,val)
    if d['Missing Cost Fields'].gt(0).any():
        st.caption(f'⚠️ {int(d["Missing Cost Fields"].gt(0).sum()):,} projects have missing cost fields. Missing values are provisionally treated as ₹0 in Spending.')
    # Compare only projects with an actual Excel GP, taxable value and complete cost fields.
    check=d.loc[d['Gross Profit'].notna() & d['Taxable Value'].notna() & d['Cost Complete']].copy()
    if not check.empty:
        tolerance=1.0
        exceptions=check.loc[check['GP Difference'].abs()>tolerance]
        if not exceptions.empty:
            st.warning(f'GP reconciliation: {len(exceptions):,} of {len(check):,} comparable projects differ by more than ₹{tolerance:,.0f}. Net difference: {money(exceptions["GP Difference"].sum())}; total absolute difference: {money(exceptions["GP Difference"].abs().sum())}. Formula: Taxable Value − all five costs − Excel Gross Profit.')
            with st.expander('▸ Inspect GP reconciliation differences',expanded=False):
                columns=[c for c in ['LEAD','CUS Name','Month','Region','Taxable Value']+COSTS+['Spending','Calculated GP','Gross Profit','GP Difference'] if c in exceptions]
                pivot_table(exceptions[columns].sort_values('GP Difference',key=lambda v:v.abs(),ascending=False),key='gp_reconciliation',height=330,export_name='FIN_GP_reconciliation.csv')
    missing_taxable=d['Taxable Value'].isna().sum()
    if missing_taxable:
        st.warning(f'{missing_taxable:,} projects have no Taxable Value; their taxable-based calculated GP cannot be determined unless an Excel Gross Profit exists.')

    with st.expander('▸ Data quality · missing cost fields and taxable revenue', expanded=False):
        issues=d.loc[d['Missing Cost Fields'].gt(0) | d['Taxable Value'].isna()]
        if issues.empty:
            st.success('No missing cost or taxable revenue fields in the selected records.')
        else:
            cols=[c for c in ['LEAD','CUS Name','Month','Region','Project Value','Taxable Value']+COSTS+['Missing Cost Fields','Profit','Profit Source'] if c in issues]
            pivot_table(issues[cols],key='missing_cost_audit',height=300,export_name='FIN_missing_financial_fields.csv')

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
        bars(month,'Month',['Margin %'],'GP margin on Project Value (%)',kind='line',height=290)
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
                'Project Value','Taxable Value','Spending','Calculated GP','Profit','GP Difference','Margin %','Material Cost','Transportation Cost',
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
    st.caption('FIN V11 · Monthly chart axes use discrete calendar-month categories. Financial note: Revenue and GP margin use Project Value; calculated GP and reconciliation use Taxable Value. Each Excel row counts as one project. Missing cost components count as zero in spending; validate Gross Profit and Vendor Margin accounting before relying on profitability rankings.')

# ============================================================
# FIN V15 — PREDICTIVE FINANCIAL INTELLIGENCE
# All projections are based on Project Value (not Taxable Value).
# No causal claims: driver decomposition is descriptive.
# ============================================================
with predictive_tab:
    st.title('FIN  /  Predictive Financial Intelligence')
    st.caption('Project Value forecasting • Region outlook • Profitability • Cost drivers • Loss exposure')
    st.info('Forecasts are statistical planning scenarios, not guaranteed outcomes. Six or fewer months of history can produce unstable estimates. Predictions use only the records matching the global filters on the Overview tab.')

    hist = d.loc[d['Month'].astype(str).str.fullmatch(r'\d{4}-\d{2}')].copy()
    if hist.empty:
        st.warning('No valid invoice months in the selected data. Choose records with valid INV DATE- BK dates.')
    else:
        hist['Period'] = pd.PeriodIndex(hist['Month'], freq='M')
        monthly_actual = hist.groupby('Period').agg(
            Projects=('Project Value','size'), Revenue=('Project Value','sum'),
            Profit=('Profit','sum'), Spending=('Spending','sum'),
            Loss_Projects=('Profit',lambda s:int((s<0).sum())),
            Loss_Amount=('Profit',lambda s:float(-s[s<0].sum()))
        ).sort_index()
        # Fill genuinely missing calendar months with zero recorded activity.
        all_periods = pd.period_range(monthly_actual.index.min(),monthly_actual.index.max(),freq='M')
        monthly_actual = monthly_actual.reindex(all_periods,fill_value=0)
        monthly_actual['Margin'] = np.where(monthly_actual.Revenue.ne(0),100*monthly_actual.Profit/monthly_actual.Revenue,np.nan)
        monthly_actual['Average_Ticket'] = np.where(monthly_actual.Projects.ne(0),monthly_actual.Revenue/monthly_actual.Projects,np.nan)
        months_count=len(monthly_actual)
        st.caption(f'Historical coverage: {months_count} calendar months ({all_periods[0]} to {all_periods[-1]}). Source: invoice dates; revenue is Project Value. Missing calendar months are treated as zero observed projects, not inferred sales.')
        if months_count<3:
            st.warning('Fewer than 3 calendar months available. Historical comparisons are shown, but forecasts are withheld until at least 3 months exist.')

        horizon = st.select_slider('Forecast horizon (months)',options=[1,2,3,4,5,6],value=3,key='pred_horizon')
        conf_col,model_col=st.columns([1,2])
        with conf_col:
            scenario=st.selectbox('Planning scenario',['Baseline','Conservative','Optimistic'],key='pred_scenario')
        with model_col:
            st.caption('Method: recency-weighted linear trend blended with recent-month average; negative forecasts are floored at zero. Scenario bands use historical one-step forecast errors when sufficient observations exist, otherwise a clearly labeled planning assumption.')

        def forecast_series(series, steps, nonnegative=True):
            """Blend trend and recent weighted mean; backtest one-step errors."""
            y=np.asarray(pd.to_numeric(series,errors='coerce').fillna(0),dtype=float)
            n=len(y)
            if n<3:
                return np.array([]),np.nan,np.array([])
            def predict(v,k=1):
                v=np.asarray(v,dtype=float);n0=len(v)
                window=v[-min(n0,6):]
                x=np.arange(len(window),dtype=float)
                weights=np.linspace(1,2,len(window))
                # Avoid unstable high-order fits; robust simple linear regression.
                xbar=np.average(x,weights=weights);ybar=np.average(window,weights=weights)
                den=np.sum(weights*(x-xbar)**2)
                slope=np.sum(weights*(x-xbar)*(window-ybar))/den if den else 0.0
                # Cap projected slope to 25% of average monthly level per step.
                cap=max(abs(ybar)*.25,1.0)
                slope=float(np.clip(slope,-cap,cap))
                smooth=np.average(window[-min(3,len(window)):],weights=np.arange(1,min(3,len(window))+1))
                return np.array([.55*(ybar+slope*(len(window)-xbar+j))+.45*smooth for j in range(k)])
            pred=predict(y,steps)
            errors=[]
            for i in range(3,n):
                prev=predict(y[:i],1)[0]
                errors.append(y[i]-prev)
            err=np.array(errors,dtype=float)
            mae=float(np.mean(np.abs(err))) if len(err) else np.nan
            return (np.maximum(pred,0) if nonnegative else pred),mae,err

        def scenario_values(pred,mae,history):
            base=np.asarray(pred,dtype=float)
            if not len(base):return base,base,base,'Insufficient history'
            # No model accuracy estimate before at least one genuine holdout month.
            fallback=max(float(np.mean(np.abs(history[-min(3,len(history)):]))) *.20, 1.)
            width=mae if np.isfinite(mae) else fallback
            step=np.sqrt(np.arange(1,len(base)+1))
            lower=np.maximum(base-1.28*width*step,0)
            upper=base+1.28*width*step
            selected={'Baseline':base,'Conservative':lower,'Optimistic':upper}[scenario]
            note='Backtest MAE' if np.isfinite(mae) else 'Assumed 20% recent-average error (not validated)'
            return selected,lower,upper,note

        def prediction_plot(actual, pred, lo, hi, title, ylabel='₹', money_axis=True):
            actual_labels=[str(v) for v in actual.index]
            future=[str(v) for v in pd.period_range(actual.index[-1]+1,periods=len(pred),freq='M')]
            fig=go.Figure()
            fig.add_trace(go.Scatter(x=actual_labels,y=actual.values,mode='lines+markers',name='Actual',line=dict(color='#2563eb',width=3)))
            if len(pred):
                fig.add_trace(go.Scatter(x=future,y=hi,mode='lines',line=dict(width=0),hoverinfo='skip',showlegend=False))
                fig.add_trace(go.Scatter(x=future,y=lo,mode='lines',fill='tonexty',fillcolor='rgba(37,99,235,.12)',line=dict(width=0),name='Scenario range'))
                fig.add_trace(go.Scatter(x=[actual_labels[-1]]+future,y=[float(actual.iloc[-1])]+list(pred),mode='lines+markers',name=scenario+' forecast',line=dict(color='#059669',width=3,dash='dash')))
            fig.update_layout(title=title,height=330,margin=dict(l=10,r=12,t=50,b=35),paper_bgcolor='white',plot_bgcolor='white',legend=dict(orientation='h',y=1.15),hovermode='x unified')
            fig.update_xaxes(type='category',tickangle=-25)
            fig.update_yaxes(title=ylabel,tickprefix='₹' if money_axis else '',gridcolor='#e9eef5')
            st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})

        # 1 — Revenue forecast
        st.markdown('### 01 · Project Value revenue outlook')
        st.caption('Compares actual booked project value with next-month expectations; distinguishes project count growth from average project size.')
        revenue_pred,rev_mae,_=forecast_series(monthly_actual.Revenue,horizon)
        chosen,low,high,rev_method=scenario_values(revenue_pred,rev_mae,monthly_actual.Revenue.values)
        prediction_plot(monthly_actual.Revenue,chosen,low,high,'Actual vs forecast · Project Value')
        a,b,c=st.columns(3)
        a.metric('Latest month revenue',money(monthly_actual.Revenue.iloc[-1]))
        b.metric('Next month · '+scenario,money(chosen[0]) if len(chosen) else 'Insufficient history')
        c.metric('One-step backtest MAE',money(rev_mae) if np.isfinite(rev_mae) else 'Not measurable')
        if len(chosen):
            delta=chosen[0]-monthly_actual.Revenue.iloc[-1]
            st.markdown(f'**Interpretation:** The {scenario.lower()} scenario projects **{money(chosen[0])}** next month, **{money(delta)}** versus the most recent recorded month. Forecast uncertainty method: {rev_method}. This is a planning estimate, not a committed order book.')
        if months_count>=2:
            last,prev=monthly_actual.iloc[-1],monthly_actual.iloc[-2]
            project_delta=last.Projects-prev.Projects
            ticket_delta=(last.Average_Ticket-prev.Average_Ticket) if np.isfinite(last.Average_Ticket) and np.isfinite(prev.Average_Ticket) else np.nan
            st.caption(f'Descriptive drivers, latest vs previous month: project count {project_delta:+,.0f}; average project value change {money(ticket_delta) if np.isfinite(ticket_delta) else "not comparable"}. These factors describe revenue movement but do not prove causation.')

        # 2 — Regional forecast
        st.markdown('### 02 · Region-wise growth and next-month outlook')
        st.caption('Compare regional contribution, latest-month momentum, and forecast next-month Project Value. Thin-history regions are not forecast.')
        regions=hist.groupby(['Region','Period']).agg(Revenue=('Project Value','sum'),Projects=('Project Value','size')).reset_index()
        region_rows=[]
        for reg,part in regions.groupby('Region',dropna=False):
            s=part.set_index('Period').Revenue.reindex(all_periods,fill_value=0)
            cnt=part.set_index('Period').Projects.reindex(all_periods,fill_value=0)
            fp,err,_=forecast_series(s,1)
            sel,l,u,_=scenario_values(fp,err,s.values)
            last=float(s.iloc[-1]);previous=float(s.iloc[-2]) if len(s)>1 else np.nan
            growth=100*(last-previous)/previous if np.isfinite(previous) and previous>0 else np.nan
            region_rows.append({'Region':reg,'Projects (all)':int(cnt.sum()),'Revenue (all)':float(s.sum()),'Latest month':last,'Prior month':previous,'MoM %':growth,'Next-month forecast':float(sel[0]) if len(sel) else np.nan,'Low scenario':float(l[0]) if len(l) else np.nan,'High scenario':float(u[0]) if len(u) else np.nan,'Forecast MAE':err})
        rframe=pd.DataFrame(region_rows).sort_values('Revenue (all)',ascending=False)
        if not rframe.empty:
            top=rframe.head(12)
            fig=go.Figure()
            fig.add_bar(x=top.Region,y=top['Latest month'],name='Latest actual',marker_color='#2563eb')
            fig.add_bar(x=top.Region,y=top['Next-month forecast'],name='Next forecast',marker_color='#10b981')
            fig.update_layout(barmode='group',height=340,margin=dict(l=10,r=10,t=35,b=85),legend=dict(orientation='h',y=1.15))
            fig.update_xaxes(tickangle=-35);fig.update_yaxes(tickprefix='₹')
            st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})
            st.dataframe(rframe,use_container_width=True,hide_index=True)
            valid=rframe.dropna(subset=['Next-month forecast'])
            if len(valid):
                leader=valid.sort_values('Next-month forecast',ascending=False).iloc[0]
                st.markdown(f'**Interpretation:** **{leader.Region}** has the highest modeled next-month Project Value (**{money(leader["Next-month forecast"])}**) among the selected regions. Review pipeline capacity and the region’s forecast error before allocating targets.')

        # 3 — Profitability trend and outlook
        st.markdown('### 03 · Gross profit and margin outlook')
        st.caption('Forecasts recorded/fallback Gross Profit, and compares weighted margins against Project Value. Margin is not based on Taxable Value.')
        profit_pred,profit_mae,_=forecast_series(monthly_actual.Profit,horizon,nonnegative=False)
        profit_chosen,profit_lo,profit_hi,profit_method=scenario_values(profit_pred,profit_mae,monthly_actual.Profit.values)
        # Profit can be negative; scenario bounds must not be zero-floored.
        if len(profit_pred):
            err_width=profit_mae if np.isfinite(profit_mae) else max(float(np.mean(np.abs(monthly_actual.Profit.values[-3:])))*.20,1.)
            step=np.sqrt(np.arange(1,len(profit_pred)+1))
            profit_lo=profit_pred-1.28*err_width*step
            profit_hi=profit_pred+1.28*err_width*step
            profit_chosen={'Baseline':profit_pred,'Conservative':profit_lo,'Optimistic':profit_hi}[scenario]
        prediction_plot(monthly_actual.Profit,profit_chosen,profit_lo,profit_hi,'Gross Profit · Actual vs projected')
        if len(profit_chosen) and len(chosen):
            margin_forecast=100*profit_chosen[0]/chosen[0] if chosen[0]>0 else np.nan
            a,b,c=st.columns(3)
            a.metric('Latest GP margin',f'{monthly_actual.Margin.iloc[-1]:.1f}%' if np.isfinite(monthly_actual.Margin.iloc[-1]) else '—')
            b.metric('Projected GP next month',money(profit_chosen[0]))
            c.metric('Implied projected GP margin',f'{margin_forecast:.1f}%' if np.isfinite(margin_forecast) else '—')
            st.markdown('**Interpretation:** Profit forecasts may diverge from revenue forecasts if cost mix or project mix changes. Implied margin is forecast GP divided by forecast Project Value; validate with actual invoiced costs before using as a target.')

        # 4 — Cost mix and drivers
        st.markdown('### 04 · Cost structure and spending pressure')
        st.caption('Compare each of the five spending categories and the observed change in cost-to-project-value ratio. This is driver analysis, not a causal prediction.')
        cost_by_month=hist.groupby('Period')[COSTS+['Project Value']].sum().reindex(all_periods,fill_value=0)
        cfig=go.Figure()
        for field in COSTS:
            cfig.add_bar(x=[str(x) for x in all_periods],y=cost_by_month[field],name=field)
        cfig.update_layout(barmode='stack',height=350,margin=dict(l=10,r=10,t=35,b=65),legend=dict(orientation='h',y=1.25),yaxis=dict(tickprefix='₹'))
        st.plotly_chart(cfig,use_container_width=True,config={'displayModeBar':False})
        cost_ratio=100*monthly_actual.Spending/monthly_actual.Revenue.replace(0,np.nan)
        current_ratio=cost_ratio.iloc[-1]
        prev_ratio=cost_ratio.iloc[-2] if len(cost_ratio)>1 else np.nan
        a,b,c=st.columns(3)
        a.metric('Latest spending / Project Value',f'{current_ratio:.1f}%' if np.isfinite(current_ratio) else '—')
        b.metric('Prior spending / Project Value',f'{prev_ratio:.1f}%' if np.isfinite(prev_ratio) else '—')
        c.metric('Projects with incomplete costs',f'{int(hist["Missing Cost Fields"].gt(0).sum()):,}')
        if len(cost_by_month)>1:
            change=(cost_by_month[COSTS].iloc[-1]-cost_by_month[COSTS].iloc[-2]).sort_values(ascending=False)
            lead=change.index[0]
            st.markdown(f'**Interpretation:** **{lead}** had the largest absolute month-on-month cost increase (**{money(change.iloc[0])}**). Check project volume and KW/module mix before interpreting this as a unit-cost increase.')
        st.caption('Warning: Missing cost fields are treated as zero in existing dashboard Spending; the chart may understate true costs. Gross Profit continues to prioritize Excel GP and uses Taxable Value minus costs only as fallback.')

        # 5 — Loss exposure and risk
        st.markdown('### 05 · Loss exposure and project risk')
        st.caption('Compare historical loss counts, loss amount and loss share by month and region. Future loss exposure is an illustrative rate-based scenario, not a trained probability model.')
        loss_chart=go.Figure()
        loss_chart.add_bar(x=[str(x) for x in all_periods],y=monthly_actual.Loss_Projects,name='Loss projects',marker_color='#e05a6a')
        loss_chart.add_trace(go.Scatter(x=[str(x) for x in all_periods],y=monthly_actual.Projects,mode='lines+markers',name='Total projects',line=dict(color='#2563eb'),yaxis='y'))
        loss_chart.update_layout(height=315,margin=dict(l=10,r=10,t=35,b=55),legend=dict(orientation='h',y=1.2),yaxis=dict(title='Projects'))
        st.plotly_chart(loss_chart,use_container_width=True,config={'displayModeBar':False})
        losses=hist[hist.Profit<0]
        risk_by_region=hist.groupby('Region').agg(Projects=('Profit','size'),Loss_Projects=('Profit',lambda s:int((s<0).sum())),Revenue=('Project Value','sum')).reset_index()
        loss_sums=losses.groupby('Region').Profit.sum()
        risk_by_region['Loss_Amount']=risk_by_region.Region.map(loss_sums).fillna(0).mul(-1)
        risk_by_region['Loss_Rate_%']=100*risk_by_region.Loss_Projects/risk_by_region.Projects.replace(0,np.nan)
        risk_by_region=risk_by_region.sort_values(['Loss_Amount','Loss_Projects'],ascending=False)
        a,b,c=st.columns(3)
        a.metric('Historical loss projects',f'{len(losses):,}')
        b.metric('Historical loss amount',money(-losses.Profit.sum()))
        c.metric('Loss project rate',f'{len(losses)/len(hist)*100:.1f}%' if len(hist) else '—')
        st.dataframe(risk_by_region,use_container_width=True,hide_index=True)
        if len(chosen) and len(monthly_actual)>=3:
            last3=monthly_actual.tail(3)
            avg_ticket=last3.Revenue.sum()/last3.Projects.sum() if last3.Projects.sum()>0 else np.nan
            recent_losses=last3.Loss_Projects.sum()/last3.Projects.sum() if last3.Projects.sum()>0 else np.nan
            if np.isfinite(avg_ticket) and avg_ticket>0 and np.isfinite(recent_losses):
                est_projects=chosen[0]/avg_ticket
                est_losses=est_projects*recent_losses
                st.markdown(f'**Forward scenario:** If the recent three-month average project value (**{money(avg_ticket)}**) and loss rate (**{recent_losses:.1%}**) persist, projected next-month revenue implies about **{est_projects:,.0f} projects**, of which approximately **{est_losses:,.1f}** may be loss-making. This is a rate-based scenario, not an individualized prediction.')
        if len(risk_by_region):
            worst=risk_by_region.iloc[0]
            st.markdown(f'**Management focus:** **{worst.Region}** has the highest observed total loss amount (**{money(worst.Loss_Amount)}**). Review the underlying projects, completeness of costs, and profitability before assigning a risk cause.')

        with st.expander('▸ Download predictive analysis data and assumptions',expanded=False):
            export=monthly_actual.reset_index().rename(columns={'index':'Month'})
            export['Month']=export['Month'].astype(str)
            st.download_button('Download monthly actuals CSV',export.to_csv(index=False).encode('utf-8-sig'),'FIN_V15_Predictive_Monthly_Actuals.csv','text/csv',key='pred_csv_month')
            st.download_button('Download regional outlook CSV',rframe.to_csv(index=False).encode('utf-8-sig'),'FIN_V15_Predictive_Regional_Outlook.csv','text/csv',key='pred_csv_reg')
            st.markdown('**Definitions:** Revenue = sum(Project Value); GP = Excel Gross Profit, otherwise Taxable Value minus listed costs; GP margin = sum(GP)/sum(Project Value); loss = GP < 0. Missing months = zero recorded projects. Forecasts use no external economic, market, pipeline, or seasonal indicators. Predictions are especially uncertain with short histories and incomplete months.')
