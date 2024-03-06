import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
import altair as alt
import plotly.graph_objs as go
import plotly.express as px


Logo=Image.open('SIM-LOGO-02.jpg')
st.image(Logo,width=680)
#########################################################
def formatted_display(label, value, unit):
    formatted_value = "<span style='color:yellow'>{:,.0f}</span>".format(value)  # Format value with comma separator and apply green color
    display_text = f"{formatted_value} {unit}"  # Combine formatted value and unit
    st.write(label, display_text, unsafe_allow_html=True)
st.header('NG Report and Analysis 2024')
StartWeek = st.sidebar.selectbox('StratWeek',['1','2','3','4','5','6','7','8','9'] )
EndWeek = st.sidebar.selectbox('EndWeek',['1','2','3','4','5','6','7','8','9'] )
# ProdWeek = st.sidebar.selectbox('ProdWeek',['1','2','3','4','5','6','7','8','9'] )
NG_Type = st.sidebar.selectbox('NG-Type',[
                                        'NG - เกลียวล้มเกลียวแตก',
                                        'NG - ตามด (Porosity)',
                                        'NG - ชิ้นงานรูตื้นหรือรูลึก ',
                                        'NG - รอยกระแทกหลัง MC',
                                        'NG - รอยตะไบหลัง MC',
                                        'NG - รอยครูดลึกบนผิว MC',
                                        'NG - รอยแตกหลัง MC',
                                        'NG - ผิวไม่เรียบ มีรอยกดทับ',
                                        'NG - เชื้อรา คราบสกปรกหลัง MC',
                                        'NG - ชิ้นงานเสียรูป ชิ้นงานไม่ได้ขนาด MC',
                                        'NG - รอยแตกร้าว',
                                        'NG - ฉีดไม่เต็ม',
                                        'NG - Flowline ',
                                        'NG - รอยหักเกจ',
                                        'NG - รอยครูด',
                                        'NG - เสียรูปจากการฉีด',
                                        'NG - ไม่ได้ใส่ Steel bush',
                                        'NG - รอยตะไบ',
                                        'NG - รอยกระแทก',
                                        'NG - ผิวงานลอก',
                                        'NG - เจียรลึก เจียรแหว่ง',
                                        'NG - ปาดไม่หมด',
                                        'NG - เชื้อรา คราบสกปรก'] )

############### 2024 #####################
@st.cache_data 
def load_data_from_drive():
    url="https://docs.google.com/spreadsheets/d/1g57LMLyDlGkHpyfikvoXwO0fyzz-P_r62qV9M4GssOM/export?format=xlsx"
    data2024=pd.read_excel(url)
    return data2024
data2024 = load_data_from_drive()
NG_2024=data2024

# Week=str(Week)
NG_2024['Weeknum']=NG_2024['Weeknum'].astype(str)
filtered = NG_2024[
    (NG_2024['Weeknum'] >= StartWeek) &
    (NG_2024['Weeknum'] <= EndWeek)]
############# Files ##############
chart=filtered
###############
ALLNG=filtered
###############
TrendNG=filtered
################
NG_WK=filtered
NG_WK=filtered
NG_WK=NG_WK[['Part No.',NG_Type]]
NG_WK=NG_WK.groupby('Part No.').sum()
NG_WK=NG_WK.fillna(0)
NG_WK=NG_WK[NG_WK[NG_Type]!=0]
SUM_NG=NG_WK[NG_Type].sum()
# NG_WK
# formatted_display('NG SUM Pcs:',round(SUM_NG),'Pcs')
Top5=NG_WK.nlargest(5,NG_Type)
NG_SUM=filtered
NG_SUM=NG_SUM[['Weeknum','Part No.','ยอดตรวจงาน NG']]

Top5['NG-%']=(Top5[NG_Type]/NG_SUM['ยอดตรวจงาน NG'].sum())*100
# Top5
# formatted_display('NG SUM Pcs:',round(SUM_NG),'Pcs')

TTNG=NG_SUM['ยอดตรวจงาน NG'].sum()
############ Chart ##############################
chart=chart[['Part No.',NG_Type]]
chart=chart.groupby('Part No.').sum()
chart=chart.fillna(0)
chart=chart[chart[NG_Type]!=0]
chart=chart.nlargest(5,NG_Type)            
# Correcting the DataFrame creation
df = pd.DataFrame({
    "x": chart.index,
    "y": chart[NG_Type],
    "category": ["Part_No"] * len(chart)  # Assigning 'Part_No' to all entries as 'category'
})
fig = px.bar(df, x='x', y='y', color='category', text=df['y'].apply(lambda x: f'{x:,.0f}'),color_discrete_sequence=['#C8FF1B'])

# Updating layout to display text on top of bars
fig.update_traces(textposition='outside')
fig.update_layout(
    title=f"NG Cases by Part No on Case: {NG_Type}",
    xaxis_title="Part No",
    yaxis_title="NG_Type"
)

# st.plotly_chart(fig)
################ All NG ##########################
st.write("---")
NG=ALLNG.columns.str.startswith('NG')
NGColumns=ALLNG.loc[:,NG]
ALLNG= ALLNG[['Part No.'] + NGColumns.columns.tolist()]
ALLNG=ALLNG.groupby('Part No.').sum()
# ALLNG=ALLNG.fillna(0)
# ALLNG=ALLNG[ALLNG[NG_Type]!=0]
ALLNG['SUM-NG']=ALLNG.sum(axis=1)
ALLNGTop5=ALLNG.nlargest(10,'SUM-NG')
# ALLNGTop5=ALLNGTop5[ALLNGTop5[NG_Type]!=0]
# ALLNGTop5=ALLNGTop5[ALLNGTop5[ALLNGTop5.columns]!=0]
ALLNGTopNG=ALLNGTop5.T
ALLNGTopNG=ALLNGTopNG.drop('SUM-NG')
ALLNGTopNG['NG-Top']=ALLNGTopNG.sum(axis=1)
ALLNGTopNG=ALLNGTopNG.nlargest(10,'NG-Top')

st.subheader(f'Over All NG by week range: {StartWeek} - {EndWeek}')
st.write('NG-Part_No Top 10')
ALLNGTop5
formatted_display('TT-NG SUM Pcs:',round(TTNG),'Pcs')
st.write('NG-Types Top 10')
ALLNGTopNG
st.write("---")
st.write('NG Top5 on type:',NG_Type)
Top5
st.write("---")
st.plotly_chart(fig)
######### Trending ###################
st.write("---")
st.write('Trending Total NG by Weekly')
TrendNG.columns.str.startswith('NG')
NGColumns=TrendNG.loc[:,NG]
TrendNG= TrendNG[['Weeknum','ยอดตรวจงาน NG'] + NGColumns.columns.tolist()]
TrendNG = TrendNG.groupby('Weeknum').sum()
TrendNG['NG-PCT']=(TrendNG['ยอดตรวจงาน NG']/TrendNG['ยอดตรวจงาน NG'].sum())*100
TrendNG
############ Chart-Trend ##############################

df = pd.DataFrame({
    "x": TrendNG.index,
    "y": TrendNG['ยอดตรวจงาน NG'],
    "category": ["Week"] * len(TrendNG)  # Assigning 'Part_No' to all entries as 'category'
})
fig = px.bar(df, x='x', y='y', color='category', text=df['y'].apply(lambda x: f'{x:,.0f}'), color_discrete_sequence=['#FF451B'])

# Updating layout to display text on top of bars
fig.update_traces(textposition='outside')
fig.update_layout(
    title="NG Cases Trend by Week",
    xaxis_title="Week",
    yaxis_title="NG_Pcs"
)

st.plotly_chart(fig)
