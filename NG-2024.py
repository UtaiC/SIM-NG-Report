import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image
import altair as alt
import plotly.express as px


Logo=Image.open('SIM-LOGO-02.jpg')
st.image(Logo,width=680)
#########################################################
def formatted_display(label, value, unit):
    formatted_value = "<span style='color:yellow'>{:,.0f}</span>".format(value)  # Format value with comma separator and apply green color
    display_text = f"{formatted_value} {unit}"  # Combine formatted value and unit
    st.write(label, display_text, unsafe_allow_html=True)
st.header('NG Report and Analysis 2024')
StartWeek = st.sidebar.selectbox('StratWeek',['2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20'] )
EndWeek = st.sidebar.selectbox('EndWeek',['2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20'] )
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

############### Files Read  #####################
@st.cache_data 
def load_dataframes(sheet_names):
    url = "https://docs.google.com/spreadsheets/d/1pbzO4YI-TkW3AO6yssJgHO9F3FwWb9Rs/export?format=xlsx"
    dataframes = {}
    for sheet_name in sheet_names:
        df = pd.read_excel(url, header=7, engine='openpyxl', sheet_name=sheet_name)
        dataframes[sheet_name] = df
    return dataframes

# Load dataframes for the specified range of weeks
StartWeek = int(StartWeek)
EndWeek = int(EndWeek)
all_sheet_names = [str(week) for week in range(StartWeek, EndWeek + 1)]
dataframes = load_dataframes(all_sheet_names)
#########################
# Concatenate dataframes into a single dataframe
DataMerges = pd.concat(dataframes.values(), ignore_index=True)

# Convert columns to numeric
DataMerges['Total.4'] = pd.to_numeric(DataMerges['Total.4'], errors='coerce')
DataMerges['Beginning Balance.6'] = pd.to_numeric(DataMerges['Beginning Balance.6'], errors='coerce')

# Calculate 'QC-Prod' column
DataMerges['QC-Prod'] = DataMerges['Total.4'] - DataMerges['Beginning Balance.6']

# Group by 'Part no.' and calculate the sum of 'Beginning Balance.6', 'Total.4', and 'QC-Prod'
QC_Prod = DataMerges.groupby('Part no.')[['Beginning Balance.6', 'Total.4', 'QC-Prod']].sum()
#################################################
@st.cache_data 
def load_data_from_drive():
    url="https://docs.google.com/spreadsheets/d/1g57LMLyDlGkHpyfikvoXwO0fyzz-P_r62qV9M4GssOM/export?format=xlsx"
    data2024=pd.read_excel(url)
    return data2024
data2024 = load_data_from_drive()
NG_2024=data2024
###############
StartWeek=str(StartWeek)
EndWeek=str(EndWeek)
##############
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
########## Merge QC-Prod ###############
ALLNG=pd.merge(ALLNG,QC_Prod['QC-Prod'],left_index=True,right_index=True,how='left')
#######################################

ALLNG['SUM-NG']=ALLNG.sum(axis=1)
ALLNG['SUM-NG']=ALLNG['SUM-NG']-ALLNG['QC-Prod']
ALLNG['QC-Prod']=ALLNG['QC-Prod'].where(ALLNG['QC-Prod']>0)+ALLNG['SUM-NG']
ALLNG['NG-%']=(ALLNG['SUM-NG']/ALLNG['QC-Prod'])*100
ALLNGTop5=ALLNG.nlargest(10,'SUM-NG')

st.subheader(f'Over All NG by week range: {StartWeek} - {EndWeek}')
st.write('NG-Part_No Top 10')
ALLNGTop5
formatted_display('TT-NG SUM Pcs:',round(TTNG),'Pcs')
st.write('NG-Types Top 10')
ALLNGTopNG=ALLNGTop5.T
ALLNGTopNG['NG-Top']=ALLNGTopNG.sum(axis=1)
ALLNGTopNG=ALLNGTopNG.drop(index=['QC-Prod','NG-%','SUM-NG'])
ALLNGTopNG=ALLNGTopNG.nlargest(12,'NG-Top')
ALLNGTopNG
################## Top 10 Part NG ###################
PartTOP10=ALLNGTop5.nlargest(10,'NG-%')
df = pd.DataFrame({
    "x": PartTOP10.index,
    "y": PartTOP10['NG-%'],
    "category": ["Part NG-%"] * len(PartTOP10)  # Assigning 'Part_No' to all entries as 'category'
})
figPCT = px.bar(df, x='x', y='y', color='category', text=df['y'].apply(lambda x: f'{x:,.2f}%'), color_discrete_sequence=['#F97500'])

# Updating layout to display text on top of bars
figPCT.update_traces(textposition='outside')
figPCT.update_layout(
    title=f'Part No NG-% TOP10@ range WK0{StartWeek} - WK0{EndWeek}',
    xaxis_title="Part No",
    yaxis_title="NG_%"
)

st.plotly_chart(figPCT)
################## Top 10 Cases NG ###################
st.write("---")
df = pd.DataFrame({
    "x": ALLNGTopNG.index,
    "y": ALLNGTopNG['NG-Top'],
    "category": ["Cases NG"] * len(ALLNGTopNG)  # Assigning 'Part_No' to all entries as 'category'
})
figCases = px.bar(df, x='x', y='y', color='category', text=df['y'].apply(lambda x: f'{x:,.0f}'), color_discrete_sequence=['#F9C000'])

# Updating layout to display text on top of bars
figCases.update_traces(textposition='outside')
figCases.update_layout(
    title=f'Cases NG TOP10@ range WK0{StartWeek} - WK0{EndWeek}',
    xaxis_title="NG_Cases",
    yaxis_title="NG_Pcs",
    height=500
)

st.plotly_chart(figCases)
################################################
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
st.write("---")
df = pd.DataFrame({
    "x": TrendNG.index,
    "y": TrendNG['ยอดตรวจงาน NG'],
    "category": ["Week"] * len(TrendNG)  # Assigning 'Part_No' to all entries as 'category'
})
fig = px.bar(df, x='x', y='y', color='category', text=df['y'].apply(lambda x: f'{x:,.0f}'), color_discrete_sequence=['#FF451B'])

# Updating layout to display text on top of bars
fig.update_traces(textposition='outside')
fig.update_layout(
    title=f"NG Cases Trend by Week range WK0{StartWeek} - WK0{EndWeek}",
    xaxis_title="Week",
    yaxis_title="NG_Pcs"
)

st.plotly_chart(fig)

