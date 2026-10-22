"""Rengy Net Metering Intelligence — Excel-upload Streamlit dashboard."""
from io import BytesIO
from pathlib import Path
import re
import unicodedata
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title='Rengy | Net Metering Intelligence', page_icon='⚡', layout='wide')
st.markdown('''<style>.block-container{max-width:1550px;padding-top:1rem}.stApp{background:#f6f9fd}div[data-testid="stMetric"]{background:white;border:1px solid #dce6f1;border-radius:13px;padding:12px}div[data-testid="stPlotlyChart"]{background:white;border:1px solid #e0e9f3;border-radius:12px;padding:8px}section[data-testid="stSidebar"]{background:#edf4fb}</style>''',unsafe_allow_html=True)
ROOT=Path(__file__).resolve().parent
DATE_COLS=['Project Start Date','Installation Completed At','Net Metering Created At','Document Prepared At','Document Dispatched At','Document Received At','Document Submitted At','Net Meter Installed At']
NUM_COLS=['Net Metering Aging (Days)','Completed NM TAT (Days)']
ALIASES={
 'Lead Number':['Lead Number','Lead','Lead Code','Lead ID','leadCode','LEAD'],
 'Consultant':['Consultant','Consultant Name','Owner Name','Sales Consultant'],
 'Vendor / CP':['Vendor / CP','Vendor / CP Name','Vendor/CP','Vendor CP','CP / Vendor','Channel Partner','Vendor Name','CP Name'],
 'Region':['Region','Customer Region','Vendor Region'],
 'Lead Substage':['Lead Substage','Lead Sub Stage','leadSubStage','Sub Stage','Substage'],
 'Latest Problem':['Latest Problem','Latest Issue','Current Problem','Problem'],
 'Project Start Date':['Project Start Date','Project   Start Date','Disbursed At'],
 'Installation Completed At':['Installation Completed At','Installation   Completed At'],
 'Net Metering Created At':['Net Metering Created At','Net Metering   Created At'],
 'Document Prepared At':['Document Prepared At','Document   Prepared At'],
 'Document Dispatched At':['Document Dispatched At','Document   Dispatched At'],
 'Document Received At':['Document Received At','Document   Received At'],
 'Document Submitted At':['Document Submitted At','Document   Submitted At'],
 'Net Meter Installed At':['Net Meter Installed At','Net Meter   Installed At'],
 'Net Metering Aging (Days)':['Net Metering Aging (Days)','Net Metering   Aging (Days)'],
 'Completed NM TAT (Days)':['Completed NM TAT (Days)','Completed NM   TAT (Days)'],
}
def key(s):
 s=unicodedata.normalize('NFKC',str(s));return re.sub('[^a-z0-9]','',s.casefold())
LOOKUP={key(a):name for name, aliases in ALIASES.items() for a in aliases}
@st.cache_data(show_spinner=False,max_entries=4)
def load(raw):
 xls=pd.ExcelFile(BytesIO(raw),engine='openpyxl')
 best=(-1,None,0)
 for sheet in xls.sheet_names:
  preview=pd.read_excel(xls,sheet_name=sheet,header=None,nrows=25)
  for idx,row in preview.iterrows():
   found={LOOKUP.get(key(v)) for v in row if pd.notna(v)}
   score=len(found.intersection(set(ALIASES)))
   if score>best[0]:best=(score,sheet,int(idx))
 if best[0]<2:raise ValueError('Could not detect a Net Metering header row. Check Excel column names.')
 df=pd.read_excel(xls,sheet_name=best[1],header=best[2],engine='openpyxl')
 df.columns=[LOOKUP.get(key(c),str(c).strip()) for c in df.columns]
 if df.columns.duplicated().any():raise ValueError('Duplicate mapped column names. Please remove duplicate Excel headers.')
 df=df.dropna(how='all').copy()
 if 'Lead Number' not in df:raise ValueError('Lead Number / Lead Code column is required to count unique projects.')
 df['Lead Number']=df['Lead Number'].fillna('').astype(str).str.strip().str.replace(r'\.0$','',regex=True)
 df=df[df['Lead Number'].ne('') & ~df['Lead Number'].str.casefold().isin(['total','grand total'])]
 for c in DATE_COLS:
  if c not in df:df[c]=pd.NaT
  s=df[c];numeric=pd.to_numeric(s,errors='coerce')
  dates=pd.to_datetime(s.where(~numeric.between(20000,80000)),errors='coerce',dayfirst=True)
  excel_dates=pd.to_datetime(numeric.where(numeric.between(20000,80000)),unit='D',origin='1899-12-30',errors='coerce')
  df[c]=dates.fillna(excel_dates)
 for c in NUM_COLS:
  if c not in df:df[c]=pd.NA
  df[c]=pd.to_numeric(df[c],errors='coerce')
 for c in ['Consultant','Vendor / CP','Region','Lead Substage','Latest Problem']:
  if c not in df:df[c]='Not Provided'
  df[c]=df[c].fillna('Not Provided').astype(str).str.strip().replace({'':'Not Provided','nan':'Not Provided'})
 # A lead is a project once, even if the source has repeated lead rows.
 df=df.drop_duplicates('Lead Number',keep='last').reset_index(drop=True)
 df['NM Status']=df['Net Meter Installed At'].notna().map({True:'Installed',False:'Pending'})
 df['Project Month']=df['Project Start Date'].dt.strftime('%Y-%m').fillna('Unknown')
 return df,best[1],best[2]+1

def export_excel(frame):
 out=BytesIO()
 with pd.ExcelWriter(out,engine='openpyxl') as writer:
  frame.to_excel(writer,index=False,sheet_name='Projects')
  ws=writer.sheets['Projects'];ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
  for col in ws.columns:
   letter=col[0].column_letter;ws.column_dimensions[letter].width=min(38,max(16,max(len(str(c.value or '')) for c in list(col)[:200])+2))
 return out.getvalue()

def show_grid(frame,grid_key):
 view=frame.copy()
 for c in view.select_dtypes(include=['datetime64[ns]']).columns:view[c]=view[c].dt.strftime('%d %b %Y').fillna('')
 try:
  from st_aggrid import AgGrid,GridOptionsBuilder,GridUpdateMode,DataReturnMode
  gb=GridOptionsBuilder.from_dataframe(view)
  gb.configure_default_column(sortable=True,filter='agSetColumnFilter',resizable=True,filterParams={'excelMode':'windows','buttons':['apply','clear']})
  for c in view.select_dtypes(include=['number']).columns:gb.configure_column(c,filter='agNumberColumnFilter')
  gb.configure_pagination(paginationAutoPageSize=False,paginationPageSize=25)
  response=AgGrid(view,gridOptions=gb.build(),height=370,theme='streamlit',enable_enterprise_modules=True,update_mode=GridUpdateMode.FILTERING_CHANGED,data_return_mode=DataReturnMode.FILTERED,key=grid_key)
  filtered=response.get('data',view)
  if not isinstance(filtered,pd.DataFrame):filtered=pd.DataFrame(filtered)
 except ImportError:
  st.dataframe(view,use_container_width=True,hide_index=True,height=370)
  filtered=view
 st.download_button('Download displayed rows (Excel)',export_excel(filtered),file_name=f'{grid_key}.xlsx',key='download_'+grid_key)

def detail_dialog(title,frame):
 @st.dialog(title,width='large')
 def popup():
  st.caption(f'{len(frame):,} matching projects. Use column menus for Excel-style filtering.')
  show_grid(frame,'popup_'+re.sub('[^a-z0-9]','_',title.lower()))
 popup()

def chart_report(frame,dimension,title,report_id):
 st.markdown('#### '+title)
 t=frame.groupby(dimension,dropna=False).size().reset_index(name='Projects').sort_values('Projects',ascending=False)
 if t.empty:st.info('No records');return
 fig=px.bar(t,x=dimension,y='Projects',text='Projects',color_discrete_sequence=['#2985d8'])
 fig.update_traces(textposition='outside',cliponaxis=False)
 fig.update_layout(height=360,margin=dict(l=10,r=12,t=25,b=80),xaxis_title=None,yaxis_title=None,plot_bgcolor='white',paper_bgcolor='white',showlegend=False)
 fig.update_xaxes(tickangle=-25,type='category')
 event=st.plotly_chart(fig,use_container_width=True,on_select='rerun',selection_mode='points',key='chart_'+report_id)
 points=event.selection.points if event else []
 if points:
  selected=points[0].get('x')
  if selected is not None:
   subset=frame[frame[dimension].astype(str)==str(selected)]
   if st.button(f'Open {len(subset)} projects · {selected}',key='open_'+report_id):detail_dialog(f'{title} · {selected}',subset)
 with st.expander('Interactive data details · '+title):
  choice=st.selectbox('View category',t[dimension].tolist(),key='select_'+report_id)
  show_grid(frame[frame[dimension]==choice],report_id)

def matrix_report(frame,row_dim,col_dim,title,report_id):
 st.markdown('#### '+title)
 pivot=pd.crosstab(frame[row_dim],frame[col_dim])
 if pivot.empty:st.info('No records');return
 fig=px.imshow(pivot,text_auto=(pivot.size<=150),aspect='auto',color_continuous_scale='Blues',labels={'color':'Projects'})
 fig.update_layout(height=min(730,max(330,90+len(pivot)*33)),margin=dict(l=10,r=15,t=15,b=90),paper_bgcolor='white')
 fig.update_xaxes(tickangle=-30)
 event=st.plotly_chart(fig,use_container_width=True,on_select='rerun',selection_mode='points',key='chart_'+report_id)
 points=event.selection.points if event else []
 if points:
  x,y=points[0].get('x'),points[0].get('y')
  if x is not None and y is not None:
   subset=frame[(frame[row_dim].astype(str)==str(y))&(frame[col_dim].astype(str)==str(x))]
   if st.button(f'Open {len(subset)} projects · {y} × {x}',key='open_'+report_id):detail_dialog(f'{title} · {y} × {x}',subset)
 with st.expander('Interactive data details · '+title):
  a,b=st.columns(2)
  with a:r=st.selectbox(row_dim,pivot.index.tolist(),key='row_'+report_id)
  with b:c=st.selectbox(col_dim,pivot.columns.tolist(),key='col_'+report_id)
  show_grid(frame[(frame[row_dim]==r)&(frame[col_dim]==c)],report_id)

st.title('⚡ Rengy / Net Metering Intelligence')
st.caption('Excel-powered project analysis · Click charts to inspect contributing leads · No comment extraction or API connection')
with st.sidebar:
 st.header('Data source')
 uploaded=st.file_uploader('Upload Net Metering Excel',type=['xlsx'])
 files=sorted([p for p in ROOT.glob('*.xlsx') if not p.name.startswith('~$')])
 selected=st.selectbox('Or use workbook beside app.py',files,format_func=lambda p:p.name) if files and uploaded is None else None
if uploaded is None and selected is None:st.info('Upload an Excel workbook to begin.');st.stop()
try:df,sheet,header=load(uploaded.getvalue() if uploaded else selected.read_bytes())
except Exception as exc:st.error(f'Excel loading error: {exc}');st.stop()
with st.sidebar:
 st.caption(f'Sheet: {sheet} · Header row: {header} · Unique leads: {len(df):,}')
 st.header('Global filters')
 status=st.multiselect('NM Status',sorted(df['NM Status'].unique()),default=sorted(df['NM Status'].unique()))
 dims=['Project Month','Region','Consultant','Vendor / CP','Lead Substage','Latest Problem']
 filters={}
 for d in dims:
  vals=sorted(df[d].astype(str).unique())
  filters[d]=st.multiselect(d,vals,default=vals,key='filter_'+key(d))
filtered=df[df['NM Status'].isin(status)].copy()
for dim,choices in filters.items():filtered=filtered[filtered[dim].astype(str).isin(choices)]
if filtered.empty:st.warning('No projects match the selected filters.');st.stop()
installed=filtered['NM Status'].eq('Installed');pending=~installed
k=st.columns(5)
k[0].metric('Total Projects',f'{len(filtered):,}')
k[1].metric('Net Meters Installed',f'{installed.sum():,}')
k[2].metric('Pending NM Installations',f'{pending.sum():,}')
k[3].metric('Avg Pending Aging',f'{filtered.loc[pending,"Net Metering Aging (Days)"].mean():.1f} days' if filtered.loc[pending,'Net Metering Aging (Days)'].notna().any() else '—')
k[4].metric('Avg Completed NM TAT',f'{filtered.loc[installed,"Completed NM TAT (Days)"].mean():.1f} days' if filtered.loc[installed,'Completed NM TAT (Days)'].notna().any() else '—')
st.caption('Installed = Net Meter Installed At has a valid date; Pending = date missing. Averages use supplied Excel TAT/aging values only.')
if (df['Latest Problem']=='Not Provided').all():st.warning('Latest Problem column is missing or blank; problem-based charts will show Not Provided until you supply that column.')
if (df['Consultant']=='Not Provided').all():st.warning('Consultant column is missing or blank; consultant-based charts will show Not Provided.')
REPORTS=[('Lead Substage','Lead Substage',None),('Latest Problem','Latest Problem',None),('Region × Lead Substage','Region','Lead Substage'),('Region × Latest Problem','Region','Latest Problem'),('Consultant × Lead Substage','Consultant','Lead Substage'),('Consultant × Latest Problem','Consultant','Latest Problem'),('Vendor / CP × Lead Substage','Vendor / CP','Lead Substage'),('Vendor / CP × Latest Problem','Vendor / CP','Latest Problem')]
st.markdown('### Analysis reports')
options=['All 8 reports']+[f'{i:02d} · {name}' for i,(name,_,__) in enumerate(REPORTS,1)]
choice=st.selectbox('Choose report',options,index=1)
for i,(name,axis1,axis2) in enumerate(REPORTS,1):
 if choice!='All 8 reports' and choice!=options[i]:continue
 if axis2 is None:chart_report(filtered,axis1,name,str(i))
 else:matrix_report(filtered,axis1,axis2,name,str(i))
st.markdown('### Interactive Project Master')
search=st.text_input('Search any project field',placeholder='Lead, region, consultant, substage, problem...')
master=filtered
if search.strip():
 match=filtered.astype(str).apply(lambda s:s.str.contains(search.strip(),case=False,regex=False)).any(axis=1)
 master=filtered.loc[match]
st.caption(f'{len(master):,} matching projects')
show_grid(master,'full_project_master')
