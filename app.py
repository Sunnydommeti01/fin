"""Rengy | Pending Amount Intelligence — Excel-powered Streamlit dashboard.
Run: streamlit run rengy_pending_dashboard.py
"""
from io import BytesIO
from pathlib import Path
import re
import unicodedata
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title='Rengy | Pending Intelligence', page_icon='📊', layout='wide')
st.markdown('''<style>
.block-container{max-width:1580px;padding:.7rem 1.4rem 2rem}
.stApp{background:#f5f8fd;color:#17304d}
section[data-testid="stSidebar"]{background:#edf4fb}
[data-testid="stMetric"]{background:#fff;border:1px solid #dce7f3;border-radius:15px;padding:14px;box-shadow:0 2px 10px #17304d09}
[data-testid="stMetricValue"]{color:#155caa;font-size:1.65rem!important}
[data-testid="stPlotlyChart"]{background:white;border:1px solid #e1e9f4;border-radius:14px;padding:9px}
[data-testid="stExpander"]{border:1px solid #e1e9f4;background:white;border-radius:12px}
h1,h2,h3{color:#183657;letter-spacing:-.025em}
</style>''',unsafe_allow_html=True)
ROOT=Path(__file__).resolve().parent
ALIASES={
 'Lead Number':['Lead Number','Lead','Lead Code','Lead ID','leadCode','LEAD'],
 'Pending Amount':['Pending Amount','Pending Project Amount','Pending Project Value','Pending Balance','Balance Amount','Outstanding Amount','Amount Pending','Pending Payment','Pending Payments','pendingAmount'],
 'Project Start Date':['Project Start Date','Project   Start Date','projectStartDate','Disbursed At','First Transaction At'],
 'Lead Substage':['Lead Substage','Lead Sub Stage','leadSubStage','Lead Sub-Stage','Sub Stage','Substage'],
 'Latest Problem':['Latest Problem','Latest Issue','Current Problem','Problem','Latest Remarks'],
 'Region':['Region','Vendor Region','Partner Region','Customer Region'],
 'Vendor':['Vendor','Vendor Name','Vendor / CP','Vendor/CP','Vendor CP','CP / Vendor','Channel Partner','Partner Name'],
 'Consultant':['Consultant','Consultant Name','Owner Name','Lead Owner','Sales Owner'],
 'Net Meter Installed At':['Net Meter Installed At','Net Meter   Installed At'],
 'Net Metering Aging (Days)':['Net Metering Aging (Days)','Net Metering   Aging (Days)'],
 'Completed NM TAT (Days)':['Completed NM TAT (Days)','Completed NM   TAT (Days)'],
}
def norm(s):
 s=unicodedata.normalize('NFKC',str(s));return re.sub(r'[^a-z0-9]','',s.casefold())
LOOKUP={norm(alias):name for name,aliases in ALIASES.items() for alias in aliases}
REQUIRED={'Lead Number','Pending Amount','Project Start Date'}

def dates(s):
 nums=pd.to_numeric(s,errors='coerce')
 dt=pd.to_datetime(s.where(~nums.between(20000,80000)),errors='coerce',dayfirst=True)
 excel=pd.to_datetime(nums.where(nums.between(20000,80000)),unit='D',origin='1899-12-30',errors='coerce')
 return dt.fillna(excel)

@st.cache_data(show_spinner=False,max_entries=4)
def load_excel(raw):
 book=pd.ExcelFile(BytesIO(raw),engine='openpyxl');best=(-1,None,0)
 for sheet in book.sheet_names:
  preview=pd.read_excel(book,sheet_name=sheet,header=None,nrows=30,dtype=object)
  for idx,row in preview.iterrows():
   found={LOOKUP.get(norm(v)) for v in row if pd.notna(v)}
   score=len(found & set(ALIASES)) + 3*len(found & REQUIRED)
   if score>best[0]:best=(score,sheet,int(idx))
 if best[0]<6:raise ValueError('Could not detect the data header. Required: Lead Number, Pending Amount, Project Start Date.')
 df=pd.read_excel(book,sheet_name=best[1],header=best[2],dtype=object)
 df.columns=[LOOKUP.get(norm(c),str(c).strip()) for c in df.columns]
 if df.columns.duplicated().any():raise ValueError('Duplicate Excel columns after header normalization. Rename duplicate columns in the workbook.')
 missing=REQUIRED-set(df.columns)
 if missing:raise ValueError('Missing required columns: '+', '.join(sorted(missing))+'. Available: '+', '.join(df.columns))
 df=df.dropna(how='all').copy()
 df['Lead Number']=df['Lead Number'].fillna('').astype(str).str.strip().str.replace(r'\.0$','',regex=True)
 df=df[df['Lead Number'].ne('') & ~df['Lead Number'].str.casefold().isin(['total','grand total'])].copy()
 money=df['Pending Amount'].astype(str).str.replace(',','',regex=False).str.replace(r'[^\d.\-()]','',regex=True).str.replace(r'^\((.*)\)$',r'-\1',regex=True)
 df['Pending Amount']=pd.to_numeric(money,errors='coerce')
 df['Project Start Date']=dates(df['Project Start Date'])
 for c in ['Lead Substage','Latest Problem','Region','Vendor','Consultant']:
  if c not in df:df[c]='Not Provided'
  df[c]=df[c].fillna('Not Provided').astype(str).str.strip().replace({'':'Not Provided','nan':'Not Provided','None':'Not Provided'})
 # Keep the most informative duplicate per lead; do not double-count financial values.
 df['_quality']=df[['Pending Amount','Project Start Date']].notna().sum(axis=1)+(df[['Lead Substage','Latest Problem','Region','Vendor']]!='Not Provided').sum(axis=1)
 duplicates=int(df['Lead Number'].duplicated().sum())
 df=df.sort_values('_quality',kind='stable').drop_duplicates('Lead Number',keep='last').drop(columns='_quality').reset_index(drop=True)
 df['Start Month']=df['Project Start Date'].dt.strftime('%Y-%m').fillna('Unknown')
 df['Pending Flag']=df['Pending Amount'].gt(0)
 return df,best[1],best[2]+1,duplicates

def inr(x):
 if pd.isna(x):return '—'
 x=float(x);sign='−' if x<0 else '';x=abs(x)
 if x>=1e7:return f'{sign}₹{x/1e7:,.2f} Cr'
 if x>=1e5:return f'{sign}₹{x/1e5:,.2f} L'
 return f'{sign}₹{x:,.0f}'

def to_excel(frame):
 b=BytesIO()
 with pd.ExcelWriter(b,engine='openpyxl') as w:
  frame.to_excel(w,index=False,sheet_name='Data')
  ws=w.sheets['Data'];ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
  for col in ws.columns:
   ws.column_dimensions[col[0].column_letter].width=min(40,max(16,max(len(str(cell.value or '')) for cell in list(col)[:150])+2))
 return b.getvalue()

def grid(frame,key,height=370):
 view=frame.copy().reset_index(drop=True)
 for c in view.select_dtypes(include=['datetime64[ns]']).columns:view[c]=view[c].dt.strftime('%d %b %Y').fillna('')
 view=view.replace([np.inf,-np.inf],np.nan)
 try:
  from st_aggrid import AgGrid,GridOptionsBuilder,GridUpdateMode,DataReturnMode
  gb=GridOptionsBuilder.from_dataframe(view)
  gb.configure_default_column(sortable=True,filter=True,resizable=True,minWidth=130)
  for c in view.columns:
   if pd.api.types.is_numeric_dtype(view[c]):gb.configure_column(c,filter='agNumberColumnFilter',filterParams={'buttons':['apply','clear']})
   else:gb.configure_column(c,filter='agSetColumnFilter',filterParams={'excelMode':'windows','buttons':['apply','clear'],'suppressMiniFilter':False})
  gb.configure_pagination(paginationAutoPageSize=False,paginationPageSize=25)
  result=AgGrid(view,gridOptions=gb.build(),height=height,theme='streamlit',enable_enterprise_modules=True,update_mode=GridUpdateMode.FILTERING_CHANGED,data_return_mode=DataReturnMode.FILTERED,key=key)
  shown=result.get('data',view)
  if not isinstance(shown,pd.DataFrame):shown=pd.DataFrame(shown if shown is not None else view)
 except ImportError:
  st.dataframe(view,use_container_width=True,hide_index=True,height=height)
  shown=view
 st.download_button('⬇ Download filtered Excel',to_excel(shown),file_name=f'{key}.xlsx',key='download_'+key)

DETAIL_COLS=['Lead Number','Pending Amount','Project Start Date','Start Month','Region','Lead Substage','Latest Problem','Vendor','Consultant']
def detail(frame,title,key):
 @st.dialog(title,width='large')
 def popup():
  a,b=st.columns(2);a.metric('Projects',f'{len(frame):,}');b.metric('Pending Amount',inr(frame['Pending Amount'].sum()))
  grid(frame[[c for c in DETAIL_COLS if c in frame]+[c for c in frame.columns if c not in DETAIL_COLS]],'popup_'+key)
 popup()

def aggregate(df,dims):
 t=df.groupby(dims,dropna=False,observed=True).agg(Projects=('Lead Number','nunique'),Pending_Projects=('Pending Flag','sum'),Pending_Amount=('Pending Amount','sum')).reset_index()
 return t.rename(columns={'Pending_Projects':'Pending Projects','Pending_Amount':'Pending Amount'})

def style_fig(fig,height=340):
 fig.update_layout(height=height,margin=dict(l=8,r=15,t=25,b=70),paper_bgcolor='white',plot_bgcolor='white',font=dict(color='#25415f',size=11),xaxis_title=None,yaxis_title=None,showlegend=False)
 fig.update_xaxes(tickangle=-24,automargin=True,type='category')
 fig.update_yaxes(gridcolor='#edf2f8',rangemode='tozero')
 return fig

def bar_report(df,dim,title,key):
 st.markdown('#### '+title)
 t=aggregate(df,[dim]).sort_values('Pending Amount',ascending=False)
 if t.empty:st.info('No data');return
 measure=st.radio('Chart measure',['Pending Amount','Pending Projects','Projects'],horizontal=True,key='measure_'+key)
 fig=go.Figure(go.Bar(x=t[dim].astype(str),y=t[measure],marker_color='#3289e5',text=[inr(v) if measure=='Pending Amount' else f'{int(v):,}' for v in t[measure]],textposition='outside',cliponaxis=False))
 style_fig(fig,360)
 event=st.plotly_chart(fig,use_container_width=True,on_select='rerun',selection_mode='points',key='plot_'+key,config={'displayModeBar':False})
 pts=event.selection.points if event else []
 if pts:
  val=str(pts[0].get('x',''))
  subset=df[df[dim].astype(str).eq(val)]
  if st.button(f'Open {len(subset):,} projects · {val}',key='open_'+key):detail(subset,f'{title} · {val}',key+'_'+norm(val))
 with st.expander('Pivot table · category totals and project details'):
  grid(t,key+'_pivot',280)
  choice=st.selectbox('Inspect category',t[dim].astype(str).tolist(),key='pick_'+key)
  grid(df[df[dim].astype(str).eq(choice)],key+'_rows',300)

def matrix_report(df,row,col,title,key):
 st.markdown('#### '+title)
 t=aggregate(df,[row,col]);metric=st.radio('Heatmap measure',['Pending Amount','Pending Projects','Projects'],horizontal=True,key='measure_'+key)
 pivot=t.pivot(index=row,columns=col,values=metric).fillna(0)
 if pivot.empty:st.info('No data');return
 z=pivot.to_numpy(dtype=float)
 texts=[[inr(v) if metric=='Pending Amount' else str(int(v)) for v in r] for r in z]
 fig=go.Figure(go.Heatmap(z=z,x=[str(x) for x in pivot.columns],y=[str(x) for x in pivot.index],colorscale='Blues',text=texts,texttemplate='%{text}' if z.size<=140 else None,hovertemplate='%{y} × %{x}<br>%{text}<extra></extra>',colorbar=dict(title='₹' if metric=='Pending Amount' else 'Count')))
 fig.update_layout(height=min(790,max(330,120+len(pivot)*36)),margin=dict(l=10,r=12,t=18,b=85),paper_bgcolor='white',font=dict(color='#25415f',size=11))
 fig.update_xaxes(tickangle=-30,automargin=True)
 event=st.plotly_chart(fig,use_container_width=True,on_select='rerun',selection_mode='points',key='plot_'+key,config={'displayModeBar':False})
 pts=event.selection.points if event else []
 if pts:
  x,y=str(pts[0].get('x','')),str(pts[0].get('y',''))
  subset=df[df[row].astype(str).eq(y)&df[col].astype(str).eq(x)]
  if st.button(f'Open {len(subset):,} projects · {y} × {x}',key='open_'+key):detail(subset,f'{title} · {y} × {x}',key+'_'+norm(y+x))
 with st.expander('Pivot table · pending amount and project counts'):
  tab1,tab2=st.tabs(['Pending amount pivot','Project count pivot'])
  with tab1:
   p=t.pivot(index=row,columns=col,values='Pending Amount').fillna(0).reset_index();p['Grand Total']=p.select_dtypes(include='number').sum(axis=1);grid(p,key+'_amount',330)
  with tab2:
   p=t.pivot(index=row,columns=col,values='Pending Projects').fillna(0).reset_index();p['Grand Total']=p.select_dtypes(include='number').sum(axis=1);grid(p,key+'_counts',330)
  a,b=st.columns(2)
  with a:r=st.selectbox(row,pivot.index.astype(str).tolist(),key='r_'+key)
  with b:c=st.selectbox(col,pivot.columns.astype(str).tolist(),key='c_'+key)
  grid(df[df[row].astype(str).eq(r)&df[col].astype(str).eq(c)],key+'_rows',310)

st.title('Rengy / Pending Amount Intelligence')
st.caption('Excel-powered outstanding payments · Project start month · Clickable lead-level analysis')
with st.sidebar:
 st.header('Data source')
 upload=st.file_uploader('Upload project Excel',type=['xlsx'])
 files=sorted((p for p in ROOT.glob('*.xlsx') if not p.name.startswith('~$')),key=lambda p:p.name)
 selected=st.selectbox('Or choose repository workbook',files,format_func=lambda p:p.name) if files and upload is None else None
if upload is None and selected is None:st.info('Upload a workbook containing Lead Number, Pending Amount and Project Start Date.');st.stop()
try:df,sheet,header,dupes=load_excel(upload.getvalue() if upload else selected.read_bytes())
except Exception as exc:st.error(f'Excel loading error: {exc}');st.stop()
with st.sidebar:
 st.caption(f'Sheet: {sheet} · Header row: {header} · {len(df):,} unique leads')
 if dupes:st.warning(f'{dupes:,} repeated lead rows were deduplicated; check the source if these represent separate payment records.')
 st.header('Global filters')
 scope=st.radio('Financial scope',['Positive pending only','All projects'],index=0)
 dims=['Start Month','Region','Lead Substage','Latest Problem','Vendor','Consultant']
 choices={}
 for d in dims:
  vals=sorted(df[d].astype(str).unique().tolist())
  choices[d]=st.multiselect(d,vals,default=vals,key='filter_'+norm(d))
filtered=df.copy()
if scope=='Positive pending only':filtered=filtered[filtered['Pending Flag']]
for dim,selected_values in choices.items():filtered=filtered[filtered[dim].astype(str).isin(selected_values)]
if filtered.empty:st.warning('No records match these filters.');st.stop()
if df['Pending Amount'].isna().all():st.error('Pending Amount has no valid numeric values. Please check your Excel data.');st.stop()
if df['Pending Amount'].isna().any():st.warning(f'{df["Pending Amount"].isna().sum():,} leads have blank or invalid pending amounts. They are excluded from positive-pending scope.')
if df['Project Start Date'].isna().any():st.caption(f'{df["Project Start Date"].isna().sum():,} leads have no readable project start date; grouped under Unknown month.')
for dim in ['Lead Substage','Latest Problem','Region','Vendor']:
 if (df[dim]=='Not Provided').all():st.warning(f'{dim} is missing from the Excel file. Its reports will show Not Provided until that column is supplied.')

k=st.columns(4)
k[0].metric('Total Pending Amount',inr(filtered['Pending Amount'].sum()))
k[1].metric('Projects in View',f'{len(filtered):,}')
k[2].metric('Projects with Pending Amount',f'{int(filtered["Pending Flag"].sum()):,}')
k[3].metric('Avg Pending / Pending Project',inr(filtered.loc[filtered['Pending Flag'],'Pending Amount'].mean()))
st.caption('Positive pending = Pending Amount > ₹0. All totals and charts follow the sidebar filters. Amounts are summed once per unique lead.')

reports=[('Month-wise Pending · Project Start Date','Start Month',None),('Lead Substage · Counts & Pending Amount','Lead Substage',None),('Region × Lead Substage','Region','Lead Substage'),('Region × Latest Problem','Region','Latest Problem'),('Vendor × Lead Substage','Vendor','Lead Substage'),('Vendor × Latest Problem','Vendor','Latest Problem')]
st.markdown('### Pending amount analysis')
options=['All reports']+[f'{i:02d} · {r[0]}' for i,r in enumerate(reports,1)]
selection=st.selectbox('Select a report',options,index=0)
for i,(title,dim1,dim2) in enumerate(reports,1):
 if selection not in ('All reports',options[i]):continue
 if dim2 is None:bar_report(filtered,dim1,title,str(i))
 else:matrix_report(filtered,dim1,dim2,title,str(i))

st.markdown('### Interactive Project Master')
search=st.text_input('Search any field',placeholder='Lead number, vendor, region, substage, problem...')
master=filtered
if search.strip():
 q=search.strip();mask=filtered.astype(str).apply(lambda col:col.str.contains(q,case=False,regex=False)).any(axis=1)
 master=filtered.loc[mask]
st.caption(f'{len(master):,} matching project rows · filter individual columns using the column menus')
ordered=[c for c in DETAIL_COLS if c in master.columns]+[c for c in master.columns if c not in DETAIL_COLS]
grid(master[ordered],'project_master',470)
