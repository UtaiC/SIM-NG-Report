import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image
import altair as alt
import plotly.express as px
import sys


Logo=Image.open('SIM-LOGO-02.jpg')
st.image(Logo,width=900)
#########################################################
def formatted_display0(label, value, unit):
    formatted_value = "<span style='color:yellow'>{:,.0f}</span>".format(value)  # Format value with comma separator and apply green color
    display_text = f"{formatted_value} {unit}"  # Combine formatted value and unit
    st.write(label, display_text, unsafe_allow_html=True)
######################################################
def formatted_display(label, value, unit):
    formatted_value = "<span style='color:yellow'>{:,.2f}</span>".format(value)  # Format value with comma separator and apply green color
    display_text = f"{formatted_value} {unit}"  # Combine formatted value and unit
    st.write(label, display_text, unsafe_allow_html=True)
######################################################
st.header('NG Report and Analysis 2024')
StartWeek = st.sidebar.selectbox('StratWeek',['2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30'] )
EndWeek = st.sidebar.selectbox('EndWeek',['2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30'] )
# ProdWeek = st.sidebar.selectbox('ProdWeek',['1','2','3','4','5','6','7','8','9'] )
NG_Type = st.sidebar.selectbox('NG-Type',[
                                        'NG - ตามด (Porosity)',
                                        'NG - เกลียวล้มเกลียวแตก',
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
Std_Cost=pd.read_excel('STD-Cost-2023.xlsx')
# Std_Cost
################################################
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
FilesMO = load_dataframes(all_sheet_names)
#########################
# Concatenate dataframes into a single dataframe
DataMerges = pd.concat(dataframes.values(), ignore_index=True)
# DataMerges
# Convert columns to numeric
DataMerges[['Total.1','Total.3', 'Total.4','ACT.7']] = DataMerges[['Total.1','Total.3', 'Total.4','ACT.7']].apply(pd.to_numeric, errors='coerce')
DataMerges[['Beginning Balance.3','Beginning Balance.5','Beginning Balance.6']] = DataMerges[['Beginning Balance.3','Beginning Balance.5','Beginning Balance.6']].apply(pd.to_numeric, errors='coerce')
DataMerges['FN-Prod'] = DataMerges['Total.1'] - DataMerges['Beginning Balance.3']
DataMerges['MC-Prod'] = DataMerges['Total.3'] - DataMerges['Beginning Balance.5']
DataMerges['QC-Prod'] = DataMerges['Total.4'] - DataMerges['Beginning Balance.6']
DataMerges['Sales-Pcs'] = DataMerges['ACT.7']

# DataMerges[['Part no.','Total.1','Beginning Balance.3','Total.3','Beginning Balance.5','Total.4','Beginning Balance.6']]

QC_Prod = DataMerges.groupby('Part no.')[['Beginning Balance.6', 'Total.4', 'QC-Prod','ACT.7']].sum()
Prod_Pcs = DataMerges.groupby('Part no.')[['FN-Prod','MC-Prod','QC-Prod','Sales-Pcs']].sum()
################################################
@st.cache_data 
def load_data_from_drive():
    url="https://docs.google.com/spreadsheets/d/1g57LMLyDlGkHpyfikvoXwO0fyzz-P_r62qV9M4GssOM/export?format=xlsx"
    data2024=pd.read_excel(url,sheet_name='inputdata')
    return data2024
data2024 = load_data_from_drive()
NG_2024=data2024
# NG_2024
###############
StartWeek=str(StartWeek)
EndWeek=str(EndWeek)
##################
missing_values = NG_2024['Weeknum'].isna().sum()
non_finite_values = ~np.isfinite(NG_2024['Weeknum']).sum()

if missing_values > 0:
    # Handle missing values (e.g., fill with a specific value or drop rows)
    NG_2024['Weeknum'].fillna(0, inplace=True)  # Example: fill with 0
if non_finite_values > 0:
    # Handle non-finite values (e.g., replace with a specific value or drop rows)
    NG_2024['Weeknum'].replace([np.inf, -np.inf], np.nan, inplace=True)  # Example: replace with NaN

# Convert to integer type
NG_2024['Weeknum'] = NG_2024['Weeknum'].astype(int)
#################
# NG_2024['Weeknum'] = NG_2024['Weeknum'].astype(int)
filtered = NG_2024[
    (NG_2024['Weeknum'] >= int(StartWeek)) &
    (NG_2024['Weeknum'] <= int(EndWeek))]
############# Files ##############
chart=filtered
###############
ALLNG=filtered
# summ=ALLNG['ยอดตรวจงาน NG'].sum()
# ALLNG
# summ
###############
TrendNG=filtered
################
TrendNG_B=filtered
################
TrendPartNG=filtered
###############
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
ALLNG=pd.merge(ALLNG,Prod_Pcs[['FN-Prod','MC-Prod','QC-Prod','Sales-Pcs']],left_index=True,right_index=True,how='left')
#######################################

ALLNG['SUM-NG']=ALLNG.sum(axis=1)
ALLNG['SUM-NG']=ALLNG['SUM-NG']-(ALLNG['FN-Prod']+ALLNG['MC-Prod']+ALLNG['QC-Prod']+ALLNG['Sales-Pcs'])
ALLNG['NG-%']=(ALLNG['SUM-NG']/(ALLNG['FN-Prod']+ALLNG['MC-Prod']+ALLNG['QC-Prod']+ALLNG['Sales-Pcs']))*100
ALLNGTop5=ALLNG#.nlargest(10,'SUM-NG')

st.subheader(f'Over All NG by week range: {StartWeek} - {EndWeek}')
st.write('NG-Part_No All NG')
################## DF Display ###############
# ALLNGTop5=ALLNGTop5.loc[:,(ALLNGTop5!=0).any(axis=0)]
ALLNGTop5=ALLNGTop5.fillna(0)
ALLNGTop5
##################################
formatted_display0('TT-NG SUM Pcs:',round(TTNG),'Pcs')
TTQC=ALLNGTop5['QC-Prod'].sum()
formatted_display0('TT-QC Check SUM Pcs:',round(TTQC),'Pcs')
ALLNG['NG-%'] = pd.to_numeric(ALLNG['NG-%'], errors='coerce')  # Convert to numeric, coerce errors to NaN
ALLNG.replace([np.inf, -np.inf], np.nan, inplace=True)
TTNGPCT=ALLNG['NG-%'].mean()
formatted_display('TT-NG-%:',round(TTNGPCT,2),'%')
#################################
st.write('NG-Types Top 10')
ALLNGTopNG=ALLNGTop5.T
ALLNGTopNG['NG-Top']=ALLNGTopNG.sum(axis=1)
ALLNGTopNG=ALLNGTopNG.drop(index=['QC-Prod','NG-%','SUM-NG'])
ALLNGTopNG=ALLNGTopNG.nlargest(12,'NG-Top')
excluded=['FN-Prod','Sales-Pcs','MC-Prod']
mask = ~ALLNGTopNG.index.isin(excluded)
ALLNGTopNG=ALLNGTopNG[mask]
ALLNGTopNG
################## Top 10 Part NG-% ###################
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
################## Top 10 Part NG-Pcs ###################
PartTOP10=ALLNGTop5.nlargest(10,'SUM-NG')
df = pd.DataFrame({
    "x": PartTOP10.index,
    "y": PartTOP10['SUM-NG'],
    "category": ["Part NG"] * len(PartTOP10)  # Assigning 'Part_No' to all entries as 'category'
})
figPCT = px.bar(df, x='x', y='y', color='category', text=df['y'].apply(lambda x: f'{x:,.0f} Pcs'), color_discrete_sequence=['#F97500'])

# Updating layout to display text on top of bars
figPCT.update_traces(textposition='outside')
figPCT.update_layout(
    title=f'Part No NG-Pcs TOP10@ range WK0{StartWeek} - WK0{EndWeek}',
    xaxis_title="Part No",
    yaxis_title="NG_Pcs"
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
TrendNG= TrendNG[['Part No.','Weeknum','ยอดตรวจงาน NG'] + NGColumns.columns.tolist()]
TrendNG['NG-PCT']=(TrendNG['ยอดตรวจงาน NG']/TrendNG['ยอดตรวจงาน NG'].sum())*100
TrendNG = TrendNG.groupby('Weeknum').agg({'ยอดตรวจงาน NG':'sum'})
# TrendNG
###########################
TrendNG_B=pd.merge(TrendNG_B,Std_Cost,left_on='Part No.',right_on='Part_No',how='outer')
TrendNG_B['NG-Cost']=TrendNG_B['ยอดตรวจงาน NG']*TrendNG_B['FG1Cost']
TrendNG_B=TrendNG_B[['Part No.','Weeknum','ยอดตรวจงาน NG','FG1Cost','NG-Cost']]
# TrendNG_B = TrendNG_B.groupby('Part_No').agg({'Weeknum':'first','ยอดตรวจงาน NG':'first','FG1Cost':'mean','NG-Cost':'sum'})
TrendNG_B = TrendNG_B.groupby('Weeknum').agg({'ยอดตรวจงาน NG':'sum','NG-Cost':'sum'})

TrendNG_B
##################################
st.write("---")
df = pd.DataFrame({
    "x": TrendNG.index,
    "y": TrendNG['ยอดตรวจงาน NG'],
    "category": ["Week-NG"] * len(TrendNG)  # Assigning 'Part_No' to all entries as 'category'
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

############ Chart-Trend-Cost ##############################

st.write("---")
df = pd.DataFrame({
    "x": TrendNG_B.index,
    "y": TrendNG_B['NG-Cost'],
    "category": ["Week-Cost"] * len(TrendNG)  # Assigning 'Part_No' to all entries as 'category'
})
fig = px.bar(df, x='x', y='y', color='category', text=df['y'].apply(lambda x: f'{x:,.0f}.B'), color_discrete_sequence=['#FF451B'])

# Updating layout to display text on top of bars
fig.update_traces(textposition='outside')
fig.update_layout(
    title=f"NG Cost Trend by Week range WK0{StartWeek} - WK0{EndWeek}",
    xaxis_title="Week",
    yaxis_title="NG_Cost"
)
st.plotly_chart(fig)
############ Sales and NG Cost ###############
st.write('SUMARIZE  NG-Cost loss VS Sales AMT')
Std_Cost.set_index('Part_No',inplace=True)
Sales=QC_Prod['ACT.7']
Sales=pd.merge(Sales,Std_Cost['Prices'],left_index=True,right_index=True,how='outer')
Sales.drop(index=['Part no.','Part no.','KOSHIN'],inplace=True)
Sales['Sales-AMT']=Sales['ACT.7']*Sales['Prices']
SalesAMT=Sales['Sales-AMT'].sum()
NGCOST=TrendNG_B['NG-Cost'].sum()
formatted_display('Sales AMT:',round(SalesAMT,2),'B.')
formatted_display('NG Cost:',round(NGCOST,2),'B.')
LossPCT=(NGCOST/SalesAMT)*100
formatted_display('NG Cost-%:',round(LossPCT,2),'%')
st.write("---")
############# Part Monitoring ##########################
st.subheader('Trending NG by Part No')
TrendPartNG=TrendPartNG= TrendPartNG[['Part No.','Weeknum','ยอดตรวจงาน NG'] + NGColumns.columns.tolist()]
TrendPartNG['Weeknum'] = TrendPartNG['Weeknum'].astype(str)
# TrendNG
##############################
st.write('**Checking Part No NG by Weekly Trending**')
    # Get the user input for the 4-digit Part No
PartNo = st.text_input('Input 4-digit Part No')
####################################
if len(PartNo) == 4:
    PartMASS = TrendPartNG
    # PartNo
    PartMASS['Part No.']=PartMASS['Part No.'].astype(str)
    mask = PartMASS['Part No.'].str.contains(PartNo)
    matching_rows = PartMASS[mask]

    ############################
    TTPCS=matching_rows['ยอดตรวจงาน NG'].sum()
    if len(matching_rows) > 0:
        ###################
        Other_Column = ['ยอดตรวจงาน NG'] + list(NGColumns.columns)
        matching_rows = matching_rows.groupby('Weeknum').agg({'Part No.': 'first', **{col: 'sum' for col in Other_Column}})
        ########################
        matching_rows = matching_rows.loc[:, (matching_rows != 0).any(axis=0)]
        matching_rows 
        formatted_display0('Total Pcs:',round(TTPCS,2),'Pcs')
        ########################
    else:
        st.write("No matching data found.")
        st.write("Pls Input 4-digit of the end of Part No")
        sys.exit()  # Terminate the program
if len(PartNo) == 0:
    st.write("Pls Input 4-digit of the end of Part No")
    sys.exit()
############ Chart-Trend-Part No ##############################
st.write("---")
# Ensure 'matching_rows' is not empty before creating the DataFrame 'df'
if not matching_rows.empty:
    df = pd.DataFrame({
        "x": matching_rows.index,
        "y": matching_rows['ยอดตรวจงาน NG'],
        "category": ["Week-NG"] * len(matching_rows)  # Assigning 'Week-NG' to all entries as 'category'
    })
    
    fig = px.bar(df, x='x', y='y', color='category', text=df['y'].apply(lambda x: f'{x:,.0f} Pcs'), color_discrete_sequence=['#FF451B'])

    # Updating layout to display text on top of bars
    fig.update_traces(textposition='outside')
    fig.update_layout(
        title=f"Part No #{PartNo} ALL NG Trend by Week range WK0{StartWeek} - WK0{EndWeek}",
        xaxis_title="Week",
        yaxis_title="NG_Pcs"
    )
    
    st.plotly_chart(fig)
    #####################

###################### Part Trend-% ##################################
matching_rows = matching_rows.reset_index()
all_week_data = pd.DataFrame()

# Iterate over each row in matching_rows
for index, row in matching_rows.iterrows():
    # Get the week number from the current row
    week_num = row['Weeknum']
    
    # Check if the week sheet exists in FilesMO keys
    if str(week_num) in FilesMO.keys():
        # Access the data of the corresponding week sheet
        week_data = FilesMO[str(week_num)]
        week_data = week_data[['Part no.','Beginning Balance.6','Total.4']]
        week_data['Total.4'] = pd.to_numeric(week_data['Total.4'], errors='coerce')
        week_data['Beginning Balance.6'] = pd.to_numeric(week_data['Beginning Balance.6'], errors='coerce')

        # Calculate 'QC-Prod' column
        week_data['QC-Prod'] = week_data['Total.4'] - week_data['Beginning Balance.6']
        week_data['Week-Num'] = week_num
        
        # Display the DataFrame for the current week
        # st.write(f"Week {week_num} data:")
        week_data=week_data[['Part no.','QC-Prod','Week-Num']]
        all_week_data = pd.concat([all_week_data, week_data], ignore_index=True)
        # all_week_data
    else:
        # Display a message indicating that the week sheet was not found
        st.write(f"Week {week_num} sheet not found in datagram.")

# Filter out rows with 'QC-Prod' equal to 0
all_week_data = all_week_data[all_week_data['QC-Prod'] != 0]

# Group by 'Part no.' and aggregate 'QC-Prod' for each part number
# all_week_data = all_week_data.groupby('Part no.').agg({'QC-Prod': 'sum', 'Week-Num': 'first'}).reset_index()

# Fill NaN values with 0
all_week_data = all_week_data.fillna(0)

# Display the concatenated DataFrame

all_week_data['Part no.']=all_week_data['Part no.'].replace(0,'NoN')
all_week_data=all_week_data[all_week_data['Part no.'].str.contains(PartNo)]
# all_week_data
#####################
matching_PCS=pd.merge(matching_rows,all_week_data[['Week-Num','QC-Prod']],left_on='Weeknum',right_on='Week-Num',how='outer')
matching_PCS['QC-Pcs']=matching_PCS['QC-Prod']+matching_PCS['ยอดตรวจงาน NG']
matching_PCS['NG-%']=(matching_PCS['ยอดตรวจงาน NG']/matching_PCS['QC-Pcs'])*100
############ Chart-Trend-Part No ##############################
st.write("---")
matching_PCS.drop_duplicates('Weeknum',inplace=True)
matching_PCS.set_index('Weeknum',inplace=True)
# matching_PCS
# Ensure 'matching_rows' is not empty before creating the DataFrame 'df'
if not matching_rows.empty:
    df = pd.DataFrame({
        "x": matching_PCS.index,
        "y": matching_PCS['NG-%'],
        "category": ["Week-NG"] * len(matching_rows)  # Assigning 'Week-NG' to all entries as 'category'
    })
    
    fig = px.bar(df, x='x', y='y', color='category', text=df['y'].apply(lambda x: f'{x:,.2f}%'), color_discrete_sequence=['#FF451B'])

    # Updating layout to display text on top of bars
    fig.update_traces(textposition='outside')
    fig.update_layout(
        title=f"Part No #{PartNo} ALL NG-% Trend by Week range WK0{StartWeek} - WK0{EndWeek}",
        xaxis_title="Week",
        yaxis_title="NG_%"
    )
    
    # st.plotly_chart(fig)
    #####################
    NG_Case=matching_PCS[NG_Type]
    # NG_Case
if not matching_rows.empty:
    df = pd.DataFrame({
        "x": NG_Case.index,
        "y": NG_Case,
        "category": ["Week-NG"] * len(matching_rows)  # Assigning 'Week-NG' to all entries as 'category'
    })
    
    fig = px.bar(df, x='x', y='y', color='category', text=df['y'].apply(lambda x: f'{x:,.0f} Pcs'), color_discrete_sequence=['#FF451B'])

    # Updating layout to display text on top of bars
    fig.update_traces(textposition='outside')
    fig.update_layout(
        title=f"Part No #{PartNo} -{NG_Type} Trend by Week range WK0{StartWeek} - WK0{EndWeek}",
        xaxis_title="Week",
        yaxis_title="NG_Pcs"
    )
    
    st.plotly_chart(fig)

    #####################
# Prod_Pcs
# matching_PCS
# # matching_PCS=pd.merge(matching_PCS,Prod_Pcs,left_on='Part No.',right_index=True,how='left')
# # matching_PCS
# matching_PCS['Check-Pcs']=(matching_PCS['FN-Prod']+matching_PCS['MC-Prod']+matching_PCS['QC-Prod']+matching_PCS['Sales-Pcs'])
# NG_CasePCT = matching_PCS[[NG_Type, 'QC-Pcs']]
# NG_CasePCT['Case-%'] = (NG_CasePCT[NG_Type] / NG_CasePCT['QC-Pcs']) * 100
# # NG_CasePCT
# if not matching_rows.empty:
#     df = pd.DataFrame({
#         "x": NG_CasePCT.index,
#         "y": NG_CasePCT['Case-%'],
#         "category": ["Week-NG"] * len(matching_rows)  # Assigning 'Week-NG' to all entries as 'category'
#     })
    
#     fig = px.bar(df, x='x', y='y', color='category', text=df['y'].apply(lambda x: f'{x:,.2f} %'), color_discrete_sequence=['#FF451B'])

#     # Updating layout to display text on top of bars
#     fig.update_traces(textposition='outside')
#     fig.update_layout(
#         title=f"Part No #{PartNo} -{NG_Type}-% Trend by Week range WK0{StartWeek} - WK0{EndWeek}",
#         xaxis_title="Week",
#         yaxis_title="NG_%"
#     )
    
#     st.plotly_chart(fig)
st.write("---End Application---")
