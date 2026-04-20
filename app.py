import streamlit as st,pandas as pd,requests as r,json,io

st.title("Практическая №10 (Colab)")
u="http://127.0.0.1:8000/p"

dx=st.sidebar.number_input("dx",0.0)
dy=st.sidebar.number_input("dy",0.0)
dz=st.sidebar.number_input("dz",0.0)
m=st.sidebar.number_input("m",0.0)

f=st.file_uploader("Excel",type=['xlsx'])

if f:
  df=pd.read_excel(f)
  st.write(df.head())
  if st.button("Расчет"):
    f.seek(0)
    ps={"dx":dx,"dy":dy,"dz":dz,"m":m}
    res=r.post(u,files={"file":f},data={"params":json.dumps(ps)})
    if res.status_code==200:
      st.download_button("Скачать DOCX",res.content,"report.docx")
