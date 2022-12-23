from ast import Str
from email.mime import image
from email.policy import default
from operator import concat
from os import renames
from re import A
from select import select
from ssl import Options
from tkinter import Menu
from turtle import width
import streamlit as st
from streamlit_option_menu import option_menu
import pandas as pd
from PIL import Image
import numpy as np
#######################################################
Logo=Image.open('SIM-LOGO-02.jpg')
st.image(Logo,width=700)
st.subheader("Costing Analysis Report:")

#######################################################
################# Main DATA FILES ##########################
db=pd.read_excel('Database-2022.xlsx')
Sales=pd.read_excel(r"C:\Users\utaie\Desktop\SIM-Main-App\inv-4.xlsx",header=4)
Sales[['วันที่','ลูกค้า','ชื่อสินค้า','รหัสสินค้า']]=Sales[['วันที่','ลูกค้า','ชื่อสินค้า','รหัสสินค้า']].astype(str)
Sales=Sales[['วันที่','ลูกค้า','ชื่อสินค้า','จำนวน','มูลค่าสินค้า','รหัสสินค้า']]
filt=(Sales['ลูกค้า'].str.contains('VALEO')|
Sales['ชื่อสินค้า'].str.contains('STEEL')|
Sales['ลูกค้า'].str.contains('แครทโค')|
Sales['ลูกค้า'].str.contains('เซนทรัล เมทัล')|
Sales['ลูกค้า'].str.contains('ทีบีเคเค')&
~Sales['ชื่อสินค้า'].str.contains('Mold')&
~Sales['รหัสสินค้า'].str.contains('PART')&
~Sales['รหัสสินค้า'].str.contains('DENSE'))
Sales=Sales.loc[filt]
################# MENU ################################
MENU=st.sidebar.radio("Select MENU",['MENU Process Cost','MENU Unit Cost'])
Main=st.sidebar.radio("Select Process",['Melting-Cost','Mat-Cost','DC-Cost','FN-Cost','SB-Cost','MC-Cost','QC-Cost','SUB-Cost'])
Main2=st.sidebar.radio("Select Cost and Sales by Month",['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])
#######################################################
############## Month Selected #########################
if MENU=='MENU Process Cost':
    Minput=st.selectbox('Input-Month',['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])
    MAPYM={'Month':['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
    YMAP={'Jan':'2022-01','Feb':'2022-02','Mar':'2022-03','Apr':'2022-04','May':'2022-05','Jun':'2022-06','Jul':'2022-07',
    'Aug':'2022-08','Sep':'2022-09','Oct':'2022-10','Nov':'2022-12','Dec':'2022-12'}
    MAPYM=pd.DataFrame(MAPYM)
    MAPYM['Year']=MAPYM['Month'].map(YMAP)
    MAPYM=MAPYM[MAPYM['Month']==Minput]
    Y=MAPYM['Year'].to_string(index=False)
    YMInput=Y
    ######################## SIM Production Data ###############################
    DCproddata=pd.read_excel('DC-Report-Jan-Sep-2022.xlsx',sheet_name=Minput)
    DCproddata.rename(columns={'Good-Pcs':'Good Parts'},inplace=True)
    DCproddata=pd.merge(DCproddata,db,on='Part_No',how='left')
    FNproddata=pd.read_excel('FN-Record-Jan-Sep-2022.xlsx')
    FNproddata['Date']=FNproddata['Date'].astype(str)
    FNproddata=pd.merge(FNproddata,db,on='Part_No',how='left')
    Filt=(FNproddata['Date'].str.contains(YMInput))
    FNproddata=FNproddata.loc[Filt]
    SBproddata=pd.read_excel('Shot Blasting Record-Jan-Sep-2022.xlsx')
    SBproddata['Date']=SBproddata['Date'].astype(str)
    SBproddata=pd.merge(SBproddata,db,on='Part_No',how='left')
    Filt=(SBproddata['Date'].str.contains(YMInput))
    SBproddata=SBproddata.loc[Filt]
    MCproddata=pd.read_excel('MC Record-Jan-Sep-2022.xlsx')
    MCproddata['Date']=MCproddata['Date'].astype(str)
    MCproddata=pd.merge(MCproddata,db,on='Part_No',how='left')
    Filt=(MCproddata['Date'].str.contains(YMInput))
    MCproddata=MCproddata.loc[Filt]
    QCproddata=pd.read_excel('QC-Record-Jan-Sep-2022.xlsx')
    QCproddata=pd.merge(QCproddata,db,on='Part_No',how='left')
    #################### Cost Data ###############################
    Budgetdata=pd.read_excel('Budget-Control-Jan-Sep-2022.xlsx')
    Budgetdata['Total Budget']=Budgetdata.sum(axis=1)
    Budgetdata=Budgetdata.fillna(0)
    ############## New Cost Data ##################################
    COST=pd.read_excel('Report-Costing-Dec-08-2022.xlsx')
    COST[['DEPT','ACC-CODE']]=COST[['DEPT','ACC-CODE']].astype(str)
    ###################### Melting-Cost ###################
    if Main=='Melting-Cost':
        st.write('New MT Cost Data',Minput)
        MTCOST=COST[COST['ACC-CODE'].str.contains('611')]
        MTCOST=MTCOST[['ACC-CODE','ACC-NAME','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']]
        MTCOST.set_index('ACC-CODE',inplace=True)
        MTCOST=MTCOST[['ACC-NAME',Minput]]
        MTCOST
        MTDATATT=MTCOST[Minput].sum()
        st.write('Melting Cosr SUM:',round(MTDATATT,2),'B')
    ####################################################################
        st.write('Melting Cost Breakdown:',Minput)
        DC=DCproddata.fillna(0)
        DC=DC[DC['Good Parts']!=0]
        DC['Good Parts']=(DC['Good Parts'])
        DC['MT-Cost']=MTDATATT
        DC['TT-SH-Weight']=(DC['Shot-Weight']/DC['Part-Cavity'])*DC['Good Parts']
        DC['TT-SH-%']=(DC['TT-SH-Weight']/DC['TT-SH-Weight'].sum())*100
        DC['Part-Cost']=(DC['MT-Cost']*DC['TT-SH-%'])/100
        DC['MT-Pcs-Cost']=DC['Part-Cost']/DC['Good Parts']
        DC.set_index('Part_No',inplace=True)
        DC[['Good Parts','MT-Cost','Shot-Weight','TT-SH-Weight','TT-SH-%','Part-Cost','MT-Pcs-Cost']]

        SUMMT=DC.groupby('Part_No').agg({'MT-Pcs-Cost':np.mean})
        SUMMT.to_excel('Melting.xlsx')
        AvgMTCost=DC['MT-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((DC['Good Parts'].sum())),'Pcs')
        st.write('Avg Melting Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total Melting Cost',round(DC['Part-Cost'].sum(),2),'B')
        ReportMT=DC[['Good Parts','MT-Pcs-Cost']].groupby('Part_No').agg({'Good Parts':'sum','MT-Pcs-Cost':'mean'})
        ############### To Excel ##################################
        MonthList={'Month':['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
        MAPNAME={'Jan':'\Melting-Rev031-Jan.xlsx','Feb':'\Melting-Rev031-Feb.xlsx','Mar':'\Melting-Rev031-Mar.xlsx','Apr':'\Melting-Rev031-Apr.xlsx',
        'May':'\Melting-Rev031-May.xlsx','Jun':'\Melting-Rev031-Jun.xlsx','Jul':'\Melting-Rev031-Jul.xlsx','Aug':'\Melting-Rev031-Aug.xlsx',
        'Sep':'\Melting-Rev031-Sep.xlsx','Oct':'\Melting-Rev031-Oct.xlsx','Nov':'\Melting-Rev031-Nov.xlsx','Dec':'\Melting-Rev031-Dec.xlsx'}
        MonthList=pd.DataFrame(MonthList)
        MonthList['File-Name']=MonthList['Month'].map(MAPNAME)
        MonthList=MonthList[MonthList['Month']==Minput]
        EXCELNAME=MonthList['File-Name'].to_string(index=False)
        st.write("---")
        Path=r"C:\Users\utaie\Desktop\Costing\DATA-Rev031"
        NAME=Path+EXCELNAME
        ReportMT.to_excel(NAME,sheet_name=Minput,engine='xlsxwriter')
        st.write('Data had export to excel:',NAME)
        st.success('End of Melting Cost Analysis Report')
    ###################### Material-Cost ###################
    if Main=='Mat-Cost':
        st.write('ADC-12 Cost:',Minput)
        DC12COST=COST[COST['ACC-CODE'].str.contains('61701-01')]
        DC12COST[['ACC-NAME',Minput]]
        DC12COST['Total Actual']=DC12COST[Minput]

        DC=DCproddata.fillna(0)
        filt12=DC['Ingot-Type']=='ADC-12'
        DC12=DC[filt12]
        DC12=DC12.fillna(0)
        # DC12=pd.concat([DC12],axis=0)
        DC12=DC12[DC12['Good Parts']!=0]
        DC12['Good Parts']=(DC12['Good Parts'])
        DC12['Mat12-Cost']=(DC12COST['Total Actual']).sum()
        DC12['Mat12-Cost']=(DC12['Mat12-Cost'])
        DC12['TT-P-Weight']=(DC12['Part-Weight'])*(DC12['Good Parts'])
        DC12['TT-P-%']=(DC12['TT-P-Weight']/DC12['TT-P-Weight'].sum())*100
        DC12['Part-Cost']=(DC12['Mat12-Cost']*DC12['TT-P-%'])/100
        DC12['ADC12-Pcs-Cost']=(DC12['Part-Cost'])/(DC12['Good Parts'])
        DC12.set_index('Part_No',inplace=True)
        DC12[['Good Parts','Mat12-Cost','Part-Weight','TT-P-Weight','TT-P-%','Part-Cost','ADC12-Pcs-Cost','Ingot-Type']]
        SUMAD12=DC12.groupby('Part_No').agg({'ADC12-Pcs-Cost':np.mean,'Ingot-Type':'first'})
        AvgMTCost=DC12['ADC12-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((DC12['Good Parts'].sum())),'Pcs')
        st.write('Avg ADC-12 Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total ADC-12 Cost',round(DC12['Part-Cost'].sum(),2),'B')
        ##################### ADC 14 Cost #####################################
        st.write('ADC-14 Cost:',Minput)

        DC14COST=COST[COST['ACC-CODE'].str.contains('61701-02')]
        DC14COST[['ACC-NAME',Minput]]
        DC14COST['Total Actual']=DC14COST[Minput]
        DC=DCproddata.fillna(0)
        filt14=DC['Ingot-Type']=='ADC-14'
        DC14=DC[filt14]
        DC14=DC14.fillna(0)
        # DC14=pd.concat([DC14],axis=0)
        DC14=DC14[DC14['Good Parts']!=0]
        DC14['Good Parts']=(DC14['Good Parts'])
        DC14['Mat14-Cost']=(DC14COST['Total Actual']).sum()
        DC14['Mat14-Cost']=(DC14['Mat14-Cost'])
        DC14['TT-P-Weight']=(DC14['Part-Weight'])*(DC14['Good Parts'])
        DC14['TT-P-%']=(DC14['TT-P-Weight']/DC14['TT-P-Weight'].sum())*100
        DC14['Part-Cost']=(DC14['Mat14-Cost']*DC14['TT-P-%'])/100
        DC14['ADC14-Pcs-Cost']=(DC14['Part-Cost'])/(DC14['Good Parts'])
        DC14.set_index('Part_No',inplace=True)
        DC14[['Good Parts','Mat14-Cost','Part-Weight','TT-P-Weight','TT-P-%','Part-Cost','ADC14-Pcs-Cost','Ingot-Type']]
        SUMAD14=DC14.groupby('Part_No').agg({'ADC14-Pcs-Cost':np.mean,'Ingot-Type':'first'})
        AvgMTCost=DC14['ADC14-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((DC14['Good Parts'].sum())),'Pcs')
        st.write('Avg ADC-14 Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total ADC-14 Cost',round(DC14['Part-Cost'].sum(),2),'B')
        ######### Material Cost #########################
        st.write('SUM Material Pcs Cost:',Minput)
        SUMMaterial=pd.concat([SUMAD12,SUMAD14],axis=0)
        SUMMaterial=SUMMaterial.fillna(0)
        SUMMaterial['Mat-Pcs-Cost']=SUMMaterial['ADC12-Pcs-Cost']+SUMMaterial['ADC14-Pcs-Cost']
        SUMMaterial=SUMMaterial[['Mat-Pcs-Cost','Ingot-Type']]
        
        ############### To Excel ##################################
        MonthList={'Month':['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
        MAPNAME={'Jan':'\Material-Rev031-Jan.xlsx','Feb':'\Material-Rev031-Feb.xlsx','Mar':'\Material-Rev031-Mar.xlsx','Apr':'\Material-Rev031-Apr.xlsx',
        'May':'\Material-Rev031-May.xlsx','Jun':'\Material-Rev031-Jun.xlsx','Jul':'\Material-Rev031-Jul.xlsx','Aug':'\Material-Rev031-Aug.xlsx',
        'Sep':'\Material-Rev031-Sep.xlsx','Oct':'\Material-Rev031-Oct.xlsx','Nov':'\Material-Rev031-Nov.xlsx','Dec':'\Material-Rev031-Dec.xlsx'}
        MonthList=pd.DataFrame(MonthList)
        MonthList['File-Name']=MonthList['Month'].map(MAPNAME)
        MonthList=MonthList[MonthList['Month']==Minput]
        EXCELNAME=MonthList['File-Name'].to_string(index=False)
        st.write("---")
        Path=r"C:\Users\utaie\Desktop\Costing\DATA-Rev031"
        NAME=Path+EXCELNAME
        SUMMaterial.to_excel(NAME,sheet_name=Minput,engine='xlsxwriter')
        st.write('Data had export to excel:',NAME)
        st.success('End of Material Cost Analysis Report')
        ###################### DC-Cost ###################
    if Main=='DC-Cost':
        st.write('Diecasting Cost 2022:',Minput)
        st.write('New DC Cost Data',Minput)
        DCCOST=COST[COST['ACC-CODE'].str.contains('612')]

        DCSUB=DCCOST[DCCOST['ACC-NAME'].str.contains('ค่าจ้างผลิตชิ้นงาน')]
        DCSUB=DCSUB[Minput].sum()
        DCCOST=DCCOST[['ACC-CODE','ACC-NAME','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']]
        
        DCCOST.set_index('ACC-CODE',inplace=True)
        DCCOST[['ACC-NAME',Minput]]
        CASTDATATT=(DCCOST[Minput].sum())
        st.write('Total DC Cost',Minput,round(CASTDATATT,2),'B')
        DCDPCOST=DCCOST[DCCOST['ACC-NAME'].str.contains('350T')|DCCOST['ACC-NAME'].str.contains('400T')|DCCOST['ACC-NAME'].str.contains('500T')|
        DCCOST['ACC-NAME'].str.contains('650T')]
        DCDPCOST.set_index('ACC-NAME',inplace=True)
        DCDPCOST.rename(index={'ค่าเช่าเครื่องจักร Die casting 350T':'350T-02','ค่าเสื่อมราคา-เครื่องมือ Die casting 350T':'350T-02','ค่าเสื่อมราคา-เครื่องจักร Die casting 400T-01':'400T-01',
        'ค่าเสื่อมราคา-เครื่องมือ Die casting 400T-01':'400T-01','ค่าเสื่อมราคา-เครื่องจักร Die casting 400T-02':'400T-02','ค่าเสื่อมราคา-เครื่องมือ Die casting 400T-02':'400T-02',
        'ค่าเสื่อมราคา-อุปกรณ์อายุ<1ปี-Die casting 400T-02':'400T-02','ค่าเสื่อมราคา-เครื่องจักร Die casting 500T':'500T','ค่าเสื่อมราคา-เครื่องมือ Die casting 500T':'500T',
        'ค่าเสื่อมราคา-ปรับปรุงเครื่องจักร Die casting 650T':'650T','ค่าเสื่อมราคา-เครื่องมือ Die casting 650T':'650T','ค่าเช่าเครื่องจักร Die casting 650T':'650T'},inplace=True)
        # HDMC=st.selectbox('Input-HDMC',['350T-02','400T-01','400T-02','500T','650T'])
        # st.write('DC Cost:',Minput,HDMC)
        # CASTDATATT
        ALLDP=DCDPCOST[Minput].groupby('ACC-NAME').sum()
        # ALLDP
        DP350T=DCDPCOST.loc['350T-02',Minput].sum()
        # DP350T
        DP400T01=DCDPCOST.loc['400T-01',Minput].sum()
        # DP400T01
        DP400T02=DCDPCOST.loc['400T-02',Minput].sum()
        DP500T=DCDPCOST.loc['500T',Minput].sum()
        # DP500T
        DP650T=DCDPCOST.loc['650T',Minput].sum()
        # DP650T
        Count=DCproddata.groupby('HDMC')['Good Parts'].sum()
        Count=pd.merge(Count,ALLDP,left_index=True,right_index=True,how='left')
        Count
        COUNTDP=Count[Minput].sum()
        # COUNTDP
        st.write('Total Production:',Minput,round(DCproddata['Good Parts'].sum()),'Pcs')
        # Count=Count.count()
        ################# PCS PCT% ##################################
        PCT350T=DCproddata[DCproddata['HDMC']=='350T-02']['Good Parts']
        PCT350T=PCT350T.sum()
        PCT350T=PCT350T/DCproddata['Good Parts'].sum()
        # PCT350T
        PCT400T01=DCproddata[DCproddata['HDMC']=='400T-01']['Good Parts']
        PCT400T01=PCT400T01.sum()
        PCT400T01=PCT400T01/DCproddata['Good Parts'].sum()
        # PCT400T01
        PCT400T02=DCproddata[DCproddata['HDMC']=='400T-02']['Good Parts']
        PCT400T02=PCT400T02.sum()
        PCT400T02=PCT400T02/DCproddata['Good Parts'].sum()
        # PCT400T02
        PCT500T=DCproddata[DCproddata['HDMC']=='500T']['Good Parts']
        PCT500T=PCT500T.sum()
        PCT500T=PCT500T/DCproddata['Good Parts'].sum()
        # PCT500T
        PCT650T=DCproddata[DCproddata['HDMC']=='650T']['Good Parts']
        PCT650T=PCT650T.sum()
        PCT650T=PCT650T/DCproddata['Good Parts'].sum()
        # PCT650T
        # ############# 350T ###################################
        st.write('350T-02 Cost Analysis')
        filtHDMC=DCproddata['HDMC']=='350T-02'
        DCPROD=DCproddata[filtHDMC]
        DCPROD['DC-Total-Cost']=(CASTDATATT-COUNTDP)*PCT350T
        DCPROD['DP-Total-Cost']=DP350T
        DCPROD['DC|DP-Cost']=DCPROD['DC-Total-Cost']+DCPROD['DP-Total-Cost']
        DCPROD['TT-P-Weight']=DCPROD['Part-Weight']*DCPROD['Good Parts']
        DCPROD['TT-P-%']=((DCPROD['Part-Weight']*DCPROD['Good Parts'])/DCPROD['TT-P-Weight'].sum())*100
        DCPROD['Part-Cost']=(DCPROD['DC|DP-Cost']*DCPROD['TT-P-%'])/100
        DCPROD['DCPROD-Pcs-Cost']=(DCPROD['Part-Cost']/DCPROD['Part-Cavity'])/DCPROD['Good Parts']
        DCPROD=DCPROD[['Part_No','DC-Total-Cost','DP-Total-Cost','DC|DP-Cost','Good Parts','TT-P-%','Part-Cost','DCPROD-Pcs-Cost','HDMC']]
        # DCPROD
        DCPROD=DCPROD.groupby('Part_No').agg({'DC|DP-Cost':np.mean,'Good Parts':np.sum,'DCPROD-Pcs-Cost':np.mean,'HDMC':"first"})
        DCPROD
        AvgMTCost=DCPROD['DCPROD-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((DCPROD['Good Parts'].sum())),'Pcs')
        st.write('Avg DC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total DC Cost',round(DCPROD['DC|DP-Cost'].mean(),2),'B')
        DC350T02=DCPROD[['Good Parts','DCPROD-Pcs-Cost','HDMC']].groupby('Part_No').agg({'Good Parts':'sum','DCPROD-Pcs-Cost':'mean','HDMC':'first'})
            # ############# 400T-01 ###################################
        st.write('400T-01 Cost Analysis')
        filtHDMC=DCproddata['HDMC']=='400T-01'
        DCPROD=DCproddata[filtHDMC]
        DCPROD['DC-Total-Cost']=(CASTDATATT-COUNTDP)*PCT400T01
        DCPROD['DP-Total-Cost']=DP400T01
        DCPROD['DC|DP-Cost']=DCPROD['DC-Total-Cost']+DCPROD['DP-Total-Cost']
        DCPROD['TT-P-Weight']=DCPROD['Part-Weight']*DCPROD['Good Parts']
        DCPROD['TT-P-%']=((DCPROD['Part-Weight']*DCPROD['Good Parts'])/DCPROD['TT-P-Weight'].sum())*100
        DCPROD['Part-Cost']=(DCPROD['DC|DP-Cost']*DCPROD['TT-P-%'])/100
        DCPROD['DCPROD-Pcs-Cost']=(DCPROD['Part-Cost']/DCPROD['Part-Cavity'])/DCPROD['Good Parts']
        DCPROD=DCPROD[['Part_No','DC-Total-Cost','DP-Total-Cost','DC|DP-Cost','Good Parts','TT-P-%','Part-Cost','DCPROD-Pcs-Cost','HDMC']]
        # DCPROD
        DCPROD=DCPROD.groupby('Part_No').agg({'DC|DP-Cost':np.mean,'Good Parts':np.sum,'DCPROD-Pcs-Cost':np.mean,'HDMC':"first"})
        DCPROD
        AvgMTCost=DCPROD['DCPROD-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((DCPROD['Good Parts'].sum())),'Pcs')
        st.write('Avg DC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total DC Cost',round(DCPROD['DC|DP-Cost'].mean(),2),'B')
        DC400T01=DCPROD[['Good Parts','DCPROD-Pcs-Cost','HDMC']].groupby('Part_No').agg({'Good Parts':'sum','DCPROD-Pcs-Cost':'mean','HDMC':'first'})
            # ############# 400T-02 ###################################
        st.write('400T-02 Cost Analysis')
        filtHDMC=DCproddata['HDMC']=='400T-02'
        DCPROD=DCproddata[filtHDMC]
        DCPROD['DC-Total-Cost']=(CASTDATATT-COUNTDP)*PCT400T02
        DCPROD['DP-Total-Cost']=DP400T02
        DCPROD['DC|DP-Cost']=DCPROD['DC-Total-Cost']+DCPROD['DP-Total-Cost']
        DCPROD['TT-P-Weight']=DCPROD['Part-Weight']*DCPROD['Good Parts']
        DCPROD['TT-P-%']=((DCPROD['Part-Weight']*DCPROD['Good Parts'])/DCPROD['TT-P-Weight'].sum())*100
        DCPROD['Part-Cost']=(DCPROD['DC|DP-Cost']*DCPROD['TT-P-%'])/100
        DCPROD['DCPROD-Pcs-Cost']=(DCPROD['Part-Cost']/DCPROD['Part-Cavity'])/DCPROD['Good Parts']
        DCPROD=DCPROD[['Part_No','DC-Total-Cost','DP-Total-Cost','DC|DP-Cost','Good Parts','TT-P-%','Part-Cost','DCPROD-Pcs-Cost','HDMC']]
        # DCPROD
        DCPROD=DCPROD.groupby('Part_No').agg({'DC|DP-Cost':np.mean,'Good Parts':np.sum,'DCPROD-Pcs-Cost':np.mean,'HDMC':"first"})
        DCPROD
        AvgMTCost=DCPROD['DCPROD-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((DCPROD['Good Parts'].sum())),'Pcs')
        st.write('Avg DC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total DC Cost',round(DCPROD['DC|DP-Cost'].mean(),2),'B')
        DC400T02=DCPROD[['Good Parts','DCPROD-Pcs-Cost','HDMC']].groupby('Part_No').agg({'Good Parts':'sum','DCPROD-Pcs-Cost':'mean','HDMC':'first'})
            # ############# 500T ###################################
        st.write('500T  Cost Analysis')
        filtHDMC=DCproddata['HDMC']=='500T'
        DCPROD=DCproddata[filtHDMC]
        DCPROD['DC-Total-Cost']=(CASTDATATT-COUNTDP)*PCT500T
        DCPROD['DP-Total-Cost']=DP500T
        DCPROD['DC|DP-Cost']=DCPROD['DC-Total-Cost']+DCPROD['DP-Total-Cost']
        DCPROD['TT-P-Weight']=DCPROD['Part-Weight']*DCPROD['Good Parts']
        DCPROD['TT-P-%']=((DCPROD['Part-Weight']*DCPROD['Good Parts'])/DCPROD['TT-P-Weight'].sum())*100
        DCPROD['Part-Cost']=(DCPROD['DC|DP-Cost']*DCPROD['TT-P-%'])/100
        DCPROD['DCPROD-Pcs-Cost']=(DCPROD['Part-Cost']/DCPROD['Part-Cavity'])/DCPROD['Good Parts']
        DCPROD=DCPROD[['Part_No','DC-Total-Cost','DP-Total-Cost','DC|DP-Cost','Good Parts','TT-P-%','Part-Cost','DCPROD-Pcs-Cost','HDMC']]
        # DCPROD
        DCPROD=DCPROD.groupby('Part_No').agg({'DC|DP-Cost':np.mean,'Good Parts':np.sum,'DCPROD-Pcs-Cost':np.mean,'HDMC':"first"})
        DCPROD
        AvgMTCost=DCPROD['DCPROD-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((DCPROD['Good Parts'].sum())),'Pcs')
        st.write('Avg DC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total DC Cost',round(DCPROD['DC|DP-Cost'].mean(),2),'B')
        DC500T=DCPROD[['Good Parts','DCPROD-Pcs-Cost','HDMC']].groupby('Part_No').agg({'Good Parts':'sum','DCPROD-Pcs-Cost':'mean','HDMC':'first'})
            # ############# 650T ###################################
        st.write('650T  Cost Analysis')
        filtHDMC=DCproddata['HDMC']=='650T'
        DCPROD=DCproddata[filtHDMC]
        DCPROD['DC-Total-Cost']=(CASTDATATT-COUNTDP)*PCT650T
        DCPROD['DP-Total-Cost']=DP650T
        DCPROD['DC|DP-Cost']=DCPROD['DC-Total-Cost']+DCPROD['DP-Total-Cost']
        DCPROD['TT-P-Weight']=DCPROD['Part-Weight']*DCPROD['Good Parts']
        DCPROD['TT-P-%']=((DCPROD['Part-Weight']*DCPROD['Good Parts'])/DCPROD['TT-P-Weight'].sum())*100
        DCPROD['Part-Cost']=(DCPROD['DC|DP-Cost']*DCPROD['TT-P-%'])/100
        DCPROD['DCPROD-Pcs-Cost']=(DCPROD['Part-Cost']/DCPROD['Part-Cavity'])/DCPROD['Good Parts']
        DCPROD=DCPROD[['Part_No','DC-Total-Cost','DP-Total-Cost','DC|DP-Cost','Good Parts','TT-P-%','Part-Cost','DCPROD-Pcs-Cost','HDMC']]
        # DCPROD
        DCPROD=DCPROD.groupby('Part_No').agg({'DC|DP-Cost':np.mean,'Good Parts':np.sum,'DCPROD-Pcs-Cost':np.mean,'HDMC':"first"})
        DCPROD
        AvgMTCost=DCPROD['DCPROD-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((DCPROD['Good Parts'].sum())),'Pcs')
        st.write('Avg DC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total DC Cost',round(DCPROD['DC|DP-Cost'].mean(),2),'B')
        DC650T=DCPROD[['Good Parts','DCPROD-Pcs-Cost','HDMC']].groupby('Part_No').agg({'Good Parts':'sum','DCPROD-Pcs-Cost':'mean','HDMC':'first'})
        ############## SUM DC Cost ############################
        st.write('DC Cost Pcs SUM:',Minput)
        SUMPCSCOSTDC=pd.concat([DC350T02,DC400T01,DC400T02,DC500T,DC650T],axis=0)
        SUMPCSCOSTDC
        st.write('Total Prod Pcs',round((SUMPCSCOSTDC['Good Parts'].sum())),'Pcs')
        st.write('Total DC Cost',round(SUMPCSCOSTDC['DCPROD-Pcs-Cost'].mean(),2),'B')
        ############### To Excel ##################################
        MonthList={'Month':['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
        MAPNAME={'Jan':'\Diecasting-Rev031-Jan.xlsx','Feb':'\Diecasting-Rev031-Feb.xlsx','Mar':'\Diecasting-Rev031-Mar.xlsx','Apr':'\Diecasting-Rev031-Apr.xlsx',
        'May':'\Diecasting-Rev031-May.xlsx','Jun':'\Diecasting-Rev031-Jun.xlsx','Jul':'\Diecasting-Rev031-Jul.xlsx','Aug':'\Diecasting-Rev031-Aug.xlsx',
        'Sep':'\Diecasting-Rev031-Sep.xlsx','Oct':'\Diecasting-Rev031-Oct.xlsx','Nov':'\Diecasting-Rev031-Nov.xlsx','Dec':'\Diecasting-Rev031-Dec.xlsx'}
        MonthList=pd.DataFrame(MonthList)
        MonthList['File-Name']=MonthList['Month'].map(MAPNAME)
        MonthList=MonthList[MonthList['Month']==Minput]
        EXCELNAME=MonthList['File-Name'].to_string(index=False)
        st.write("---")
        Path=r"C:\Users\utaie\Desktop\Costing\DATA-Rev031"
        NAME=Path+EXCELNAME
        SUMPCSCOSTDC.to_excel(NAME,sheet_name=Minput,engine='xlsxwriter')
        st.write('Data had export to excel:',NAME)
        st.success('End of Diecasting Cost Analysis Report')
            ###################### FN-Cost ###################
    if Main=='FN-Cost':
        st.write('FN Cost 2022:',Minput)
        FNCOST=COST[COST['ACC-CODE'].str.contains('613')]
        FNCOST=FNCOST[['ACC-CODE','ACC-NAME','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']]
        # FNCOST=FNCOST[~FNCOST['ACC-CODE'].str.contains('61304')]
        FNCOST.set_index('ACC-CODE',inplace=True)
        FNCOST[['ACC-NAME',Minput]]
        FNDATATT=FNCOST[Minput].sum()
        st.write('FN Cost SUM:',round(FNDATATT,2),'B')
        FN=FNproddata
        FN['Good Parts']=FN['Good Parts']
        FN['FN-Cost']=FNDATATT
        FN['FN-Pcs-Cost']=FN['FN-Cost']/(FN['Good Parts'].sum())
        FNSUMM=FN.groupby('Part_No').agg({'Good Parts':np.sum,'FN-Cost':np.mean,'Prices-Q1-22':np.mean,'Part-Weight':np.mean})
        FNSUMM['FN-Value-%']=((FNSUMM['Part-Weight']*FNSUMM['Good Parts'])/(FNSUMM['Part-Weight']*FNSUMM['Good Parts']).sum())*100
        FNSUMM['FN-SUM-Cost']=(FNSUMM['FN-Cost']*FNSUMM['FN-Value-%'])/100
        FNSUMM['FN-Pcs-Cost']=FNSUMM['FN-SUM-Cost']/(FNSUMM['Good Parts'])
        FNSUMM[['Good Parts','FN-Cost','FN-Value-%','FN-SUM-Cost','FN-Pcs-Cost']]
        FNCost=FNSUMM['FN-Pcs-Cost']
        AvgFNCost=FNSUMM['FN-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((FNSUMM['Good Parts'].sum())),'Pcs')
        st.write('Avg FN Cost/Pcs',round(AvgFNCost,2),'B')
        st.write('Total FN Cost',round(FNSUMM['FN-SUM-Cost'].sum(),2),'B')
        SUMPCSCOSTFN=FNSUMM[['Good Parts','FN-Pcs-Cost']]
        ############### To Excel ##################################
        MonthList={'Month':['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
        MAPNAME={'Jan':'\Finishing-Rev031-Jan.xlsx','Feb':'\Finishing-Rev031-Feb.xlsx','Mar':'\Finishing-Rev031-Mar.xlsx','Apr':'\Finishing-Rev031-Apr.xlsx',
        'May':'\Finishing-Rev031-May.xlsx','Jun':'\Finishing-Rev031-Jun.xlsx','Jul':'\Finishing-Rev031-Jul.xlsx','Aug':'\Finishing-Rev031-Aug.xlsx',
        'Sep':'\Finishing-Rev031-Sep.xlsx','Oct':'\Finishing-Rev031-Oct.xlsx','Nov':'\Finishing-Rev031-Nov.xlsx','Dec':'\Finishing-Rev031-Dec.xlsx'}
        MonthList=pd.DataFrame(MonthList)
        MonthList['File-Name']=MonthList['Month'].map(MAPNAME)
        MonthList=MonthList[MonthList['Month']==Minput]
        EXCELNAME=MonthList['File-Name'].to_string(index=False)
        st.write("---")
        Path=r"C:\Users\utaie\Desktop\Costing\DATA-Rev031"
        NAME=Path+EXCELNAME
        SUMPCSCOSTFN.to_excel(NAME,sheet_name=Minput,engine='xlsxwriter')
        st.write('Data had export to excel:',NAME)
        st.success('End of Finishing Cost Analysis Report')
            ###################### SB-Cost ###################
    if Main=='SB-Cost':
        st.write('SB Cost:',Minput)
        SBCOST=COST[COST['ACC-CODE'].str.contains('614')]
        SBCOST=SBCOST[['ACC-CODE','ACC-NAME','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']]
        # SBCOST=SBCOST[~SBCOST['ACC-CODE'].str.contains('61403')]
        SBCOST.set_index('ACC-CODE',inplace=True)
        SBCOST[['ACC-NAME',Minput]]
        SBDATATT=SBCOST[Minput].sum()
        st.write('SB Cost SUM:',round(SBDATATT,2),'B')
        SB=SBproddata
        SB=SB.fillna(0)
        SB=SB[SB['Good Parts']!=0]
        SB['SB-Cost']=SBDATATT
        SB['SB-Pcs-Cost']=SB['SB-Cost']/(SB['Good Parts'].sum())
        SBSUMM=SB.groupby('Part_No').agg({'Good Parts':np.sum,'SB-Cost':np.mean,'Prices-Q1-22':np.mean,'Part-Weight':np.mean})
        SBSUMM['SB-Value-%']=((SBSUMM['Part-Weight']*SBSUMM['Good Parts'])/(SBSUMM['Part-Weight']*SBSUMM['Good Parts']).sum())*100
        SBSUMM['SB-SUM-Cost']=(SBSUMM['SB-Cost']*SBSUMM['SB-Value-%'])/100
        SBSUMM['SB-Pcs-Cost']=SBSUMM['SB-SUM-Cost']/SBSUMM['Good Parts']
        SBSUMM[['Good Parts','SB-Cost','SB-Value-%','SB-SUM-Cost','SB-Pcs-Cost']]
        SBCost=SBSUMM['SB-Pcs-Cost']
        AvgSBCost=SBSUMM['SB-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((SBSUMM['Good Parts'].sum())),'Pcs')
        st.write('Avg SB Cost/Pcs',round(AvgSBCost,2),'B')
        st.write('Total SB Cost',round(SBSUMM['SB-SUM-Cost'].sum(),2),'B')
        SUMPCSCOSTSB=SBSUMM[['Good Parts','SB-Pcs-Cost']]
        ############### To Excel ##################################
        MonthList={'Month':['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
        MAPNAME={'Jan':'\ShotBlasting-Rev031-Jan.xlsx','Feb':'\ShotBlasting-Rev031-Feb.xlsx','Mar':'\ShotBlasting-Rev031-Mar.xlsx','Apr':'\ShotBlasting-Rev031-Apr.xlsx',
        'May':'\ShotBlasting-Rev031-May.xlsx','Jun':'\ShotBlasting-Rev031-Jun.xlsx','Jul':'\ShotBlasting-Rev031-Jul.xlsx','Aug':'\ShotBlasting-Rev031-Aug.xlsx',
        'Sep':'\ShotBlasting-Rev031-Sep.xlsx','Oct':'\ShotBlasting-Rev031-Oct.xlsx','Nov':'\ShotBlasting-Rev031-Nov.xlsx','Dec':'\ShotBlasting-Rev031-Dec.xlsx'}
        MonthList=pd.DataFrame(MonthList)
        MonthList['File-Name']=MonthList['Month'].map(MAPNAME)
        MonthList=MonthList[MonthList['Month']==Minput]
        EXCELNAME=MonthList['File-Name'].to_string(index=False)
        st.write("---")
        Path=r"C:\Users\utaie\Desktop\Costing\DATA-Rev031"
        NAME=Path+EXCELNAME
        SUMPCSCOSTSB.to_excel(NAME,sheet_name=Minput,engine='xlsxwriter')
        st.write('Data had export to excel:',NAME)
        st.success('End of Shot Blasting Cost Analysis Report')
    #  ###################### MC-Cost ###################
    # if Main=='MC-Cost':
    #     st.write('MC Cost')
    #     MCCOST=COST[COST['ACC-CODE'].str.contains('615')]
    #     MCCOST=MCCOST[['ACC-CODE','ACC-NAME','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']]
    #     MCCOST=MCCOST[~MCCOST['ACC-CODE'].str.contains('61505')]
    #     MCCOST.set_index('ACC-CODE',inplace=True)
    #     MCCOST[['ACC-NAME',Minput]]
    #     MCCOSTALL=MCCOST[Minput].sum()
    #     st.write('MC ALL Cost',Minput,round(MCCOST[Minput].sum(),2),'B')

    #     MCDP=MCCOST[MCCOST['ACC-NAME'].str.contains('1771')|MCCOST['ACC-NAME'].str.contains('3000')|MCCOST['ACC-NAME'].str.contains('3000')|
    #     MCCOST['ACC-NAME'].str.contains('3100')|MCCOST['ACC-NAME'].str.contains('493C')|MCCOST['ACC-NAME'].str.contains('6496')|
    #     MCCOST['ACC-NAME'].str.contains('6497')|MCCOST['ACC-NAME'].str.contains('4900')|MCCOST['ACC-NAME'].str.contains('Sleeve')|
    #     MCCOST['ACC-NAME'].str.contains('Ele')]
    #     MCDP.set_index('ACC-NAME',inplace=True)
    #     # MCDP=MCDP.set_flags(allows_duplicate_labels=True)
    #     # MCDP=MCDP.rename(index={'ค่าเช่าเครื่องจักร Machine 1771':'Z0021771A','ค่าเช่าเครื่องจักร Machine 3000':'5612603000A','ค่าเสื่อมราคา-เครื่องจักร Machine 3000':'5612603000A',
    #     # 'ค่าเช่าเครื่องจักร Machine 3100':'5612603100A','ค่าเสื่อมราคา-เครื่องจักร Machine 3100':'5612603100A','ค่าเช่าเครื่องจักร Machine 493C':'T96493CA',
    #     # 'ค่าเช่าเครื่องจักร Machine 6496':'T46496AA','ค่าเช่าเครื่องจักร Machine 6497':'T46497AA','ค่าเช่าเครื่องจักร Machine 51-4900':'5611514900A',
    #     # 'ค่าเช่าเครื่องจักร Machine Sleeve':'1050B375-RM','ค่าเช่าเครื่องจักร Machine Eletrolux':'220-00331','ค่าเช่าเครื่องจักร Machine Eletrolux':'220-00016-1',
    #     # 'ค่าเช่าเครื่องจักร Machine Eletrolux':'220-00016-2','ค่าเช่าเครื่องจักร Machine Sleeve':'5612604900A-SIM'})
    #     ELECNAME=['220-00331','220-00016-1','220-00016-2']
    #     MCDP=MCDP.rename(index={'ค่าเช่าเครื่องจักร Machine 1771':'Z0021771A','ค่าเช่าเครื่องจักร Machine 3000':'5612603000A','ค่าเสื่อมราคา-เครื่องจักร Machine 3000':'5612603000A',
    #     'ค่าเช่าเครื่องจักร Machine 3100':'5612603100A','ค่าเสื่อมราคา-เครื่องจักร Machine 3100':'5612603100A','ค่าเช่าเครื่องจักร Machine 493C':'T96493CA',
    #     'ค่าเช่าเครื่องจักร Machine 6496':'T46496AA','ค่าเช่าเครื่องจักร Machine 6497':'T46497AA','ค่าเช่าเครื่องจักร Machine 51-4900':'5611514900A',
    #     'ค่าเช่าเครื่องจักร Machine Sleeve':'1050B375-RM','ค่าเช่าเครื่องจักร Machine Eletrolux':'Electrolux-Part'})
    #     MCDP=MCDP[Minput].groupby('ACC-NAME').sum()
    #     MCDP
    #     MCDPALL=MCDP.sum()
    #     st.write('MC Parts-DP Cost',Minput,round(MCDP.sum(),2),'B')
    #     ##########################
    #     
    #         MCCOST=MCproddata.fillna(0)
    #         MCPart=st.selectbox('Input-MC-Parts',['5612603000A','5612603100A','T96493CA','Z0021771A','T46496AA','T46497AA','5611514900A',
    #         '5612604400A','5612604500A','5612604900A-SIM','220-00331','220-00016-2','220-00016-1','1050B375-RM'])
    #         #################### Elec Count DP #########################
    #         Count=MCCOST.groupby('Part_No')[['Good Parts','MC-CT']].agg({'Good Parts':np.sum,'MC-CT':np.mean})
    #         CountPCS=Count
    #         CountELEC=Count.filter(items=ELECNAME,axis=0).count()
    #         CountELEC=CountELEC['Good Parts']
    #         ELECALLDP=MCDP.filter(items=['Electrolux-Part',Minput]).sum()
    #         # ELECALLDP
    #         ELECDP=ELECALLDP/CountELEC
    #         # ELECDP
    #         ########################################################
    #         # DPONPART=MCDP[MCPart]
    #         DPMCPROD=pd.merge(CountPCS,MCDP,left_index=True,right_index=True,how='left')
    #         DPMCPROD.rename(columns={Minput:'DP'},inplace=True)
    #         ELECPRODP=DPMCPROD.filter(items=ELECNAME,axis=0).fillna(ELECDP)
    #         # ELECPRODP
    #         DPMCPROD=pd.merge(DPMCPROD,ELECPRODP,on='Part_No',suffixes=('','_ADD'),how='left')
    #         DPMCPROD.fillna(0,inplace=True)
    #         DPMCPROD['DP']=DPMCPROD['DP']+DPMCPROD['DP_ADD']
    #         st.write('MC Roduction:',Minput)
    #         DPMCPROD[['Good Parts','DP']]
    #         DPONPART=DPMCPROD.loc[MCPart,'DP']
    #         TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
    
    #         MCproddata=MCproddata[MCproddata['Part_No'].str.contains(MCPart)]
    #         MCproddata.set_index('Part_No',inplace=True)
    #         st.write('MC Cost Analysis:',MCPart)
    #         MCPROD=MCproddata
    #         MCPROD['MC-CT-%']=((MCPROD['Good Parts']*MCPROD['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
    #         MCPROD['MC-Total-Cost']=(TTCost*MCPROD['MC-CT-%'])/100
    #         MCPROD['DP-Total-Cost']=DPONPART
    #         MCPROD['MC|DP-Cost']=MCPROD['MC-Total-Cost']+MCPROD['DP-Total-Cost']
    #         MCPROD['TT-P-Weight']=MCPROD['Part-Weight']*MCPROD['Good Parts']
    #         MCPROD['TT-P-%']=((MCPROD['Part-Weight']*MCPROD['Good Parts'])/MCPROD['TT-P-Weight'].sum())*100
    #         MCPROD['Part-Cost']=(MCPROD['MC|DP-Cost']*MCPROD['TT-P-%'])/100
    #         MCPROD['MC-Pcs-Cost']=(MCPROD['Part-Cost']/MCPROD['Part-Cavity'])/MCPROD['Good Parts']
    #         MCPROD=MCPROD[['Date','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
    #         # MCPROD.set_index('Part_No',inplace=True)
    #         # MCPROD.set_index('Date',inplace=True)
    #         # MCPROD
    #         SUMMC=MCPROD.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
    #         SUMMC
    #         AvgMTCost=MCPROD['MC-Pcs-Cost'].mean()
    #         st.write('Total Prod Pcs',round((MCPROD['Good Parts'].sum())),'Pcs')
    #         st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
    #         st.write('Total MC Cost',round((MCPROD['Part-Cost'].sum()),2),'B')
    #     
    #         st.warning('No Part Production your have selected, Pls,Pick up from the list above table!')
    ############################################################
    ############## MC Rev0312 ##################################
        ###################### MC-Cost ###################
    if Main=='MC-Cost':
        st.write('MC Cost')
        MCCOST=COST[COST['ACC-CODE'].str.contains('615')]
        MCCOST=MCCOST[['ACC-CODE','ACC-NAME','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']]
        # MCCOST=MCCOST[~MCCOST['ACC-CODE'].str.contains('61505')]
        MCCOST.set_index('ACC-CODE',inplace=True)
        MCCOST[['ACC-NAME',Minput]]
        MCCOSTALL=MCCOST[Minput].sum()
        st.write('MC ALL Cost',Minput,round(MCCOST[Minput].sum(),2),'B')

        MCDP=MCCOST[MCCOST['ACC-NAME'].str.contains('1771')|MCCOST['ACC-NAME'].str.contains('3000')|MCCOST['ACC-NAME'].str.contains('3000')|
        MCCOST['ACC-NAME'].str.contains('3100')|MCCOST['ACC-NAME'].str.contains('493C')|MCCOST['ACC-NAME'].str.contains('6496')|
        MCCOST['ACC-NAME'].str.contains('6497')|MCCOST['ACC-NAME'].str.contains('4900')|MCCOST['ACC-NAME'].str.contains('Sleeve')|
        MCCOST['ACC-NAME'].str.contains('Ele')]
        MCDP.set_index('ACC-NAME',inplace=True)
        # MCDP=MCDP.set_flags(allows_duplicate_labels=True)
        # MCDP=MCDP.rename(index={'ค่าเช่าเครื่องจักร Machine 1771':'Z0021771A','ค่าเช่าเครื่องจักร Machine 3000':'5612603000A','ค่าเสื่อมราคา-เครื่องจักร Machine 3000':'5612603000A',
        # 'ค่าเช่าเครื่องจักร Machine 3100':'5612603100A','ค่าเสื่อมราคา-เครื่องจักร Machine 3100':'5612603100A','ค่าเช่าเครื่องจักร Machine 493C':'T96493CA',
        # 'ค่าเช่าเครื่องจักร Machine 6496':'T46496AA','ค่าเช่าเครื่องจักร Machine 6497':'T46497AA','ค่าเช่าเครื่องจักร Machine 51-4900':'5611514900A',
        # 'ค่าเช่าเครื่องจักร Machine Sleeve':'1050B375-RM','ค่าเช่าเครื่องจักร Machine Eletrolux':'220-00331','ค่าเช่าเครื่องจักร Machine Eletrolux':'220-00016-1',
        # 'ค่าเช่าเครื่องจักร Machine Eletrolux':'220-00016-2','ค่าเช่าเครื่องจักร Machine Sleeve':'5612604900A-SIM'})
        ELECNAME=['220-00331','220-00016-1','220-00016-2']
        MCDP=MCDP.rename(index={'ค่าเช่าเครื่องจักร Machine 1771':'Z0021771A','ค่าเช่าเครื่องจักร Machine 3000':'5612603000A','ค่าเสื่อมราคา-เครื่องจักร Machine 3000':'5612603000A',
        'ค่าเช่าเครื่องจักร Machine 3100':'5612603100A','ค่าเสื่อมราคา-เครื่องจักร Machine 3100':'5612603100A','ค่าเช่าเครื่องจักร Machine 493C':'T96493CA',
        'ค่าเช่าเครื่องจักร Machine 6496':'T46496AA','ค่าเช่าเครื่องจักร Machine 6497':'T46497AA','ค่าเช่าเครื่องจักร Machine 51-4900':'5611514900A',
        'ค่าเช่าเครื่องจักร Machine Sleeve':'1050B375-RM','ค่าเช่าเครื่องจักร Machine Eletrolux':'Electrolux-Part'})
        MCDP=MCDP[Minput].groupby('ACC-NAME').sum()
        MCDP
        MCDPALL=MCDP.sum()
        st.write('MC Parts-DP Cost',Minput,round(MCDP.sum(),2),'B')
        ##########################

        MCCOST=MCproddata.fillna(0)
        # MCPart=st.selectbox('Input-MC-Parts',['5612603000A','5612603100A','T96493CA','Z0021771A','T46496AA','T46497AA','5611514900A',
        # '5612604400A','5612604500A','5612604900A-SIM','220-00331','220-00016-2','220-00016-1','1050B375-RM'])
        #################### Elec Count DP #########################
        Count=MCCOST.groupby('Part_No')[['Good Parts','MC-CT']].agg({'Good Parts':np.sum,'MC-CT':np.mean})
        CountPCS=Count
        CountELEC=Count.filter(items=ELECNAME,axis=0).count()
        CountELEC=CountELEC['Good Parts']
        ELECALLDP=MCDP.filter(items=['Electrolux-Part',Minput]).sum()
        ELECDP=ELECALLDP/CountELEC
        ########################################################
        DPMCPROD=pd.merge(CountPCS,MCDP,left_index=True,right_index=True,how='left')
        DPMCPROD.rename(columns={Minput:'DP'},inplace=True)
        ELECPRODP=DPMCPROD.filter(items=ELECNAME,axis=0).fillna(ELECDP)
    
        DPMCPROD=pd.merge(DPMCPROD,ELECPRODP,on='Part_No',suffixes=('','_ADD'),how='left')
        DPMCPROD.fillna(0,inplace=True)
        DPMCPROD['DP']=DPMCPROD['DP']+DPMCPROD['DP_ADD']
        st.write('MC Roduction:',Minput)
        MCONProd=DPMCPROD[['Good Parts','DP']]
        MCONProd

        ############################# Part 3000 ############################
        try:
            DPONPART=DPMCPROD.loc['5612603000A','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MC3000=MCproddata[MCproddata['Part_No'].str.contains('5612603000A')]
        st.write('MC Cost Analysis:','5612603000A')
        MC3000['MC-CT-%']=((MC3000['Good Parts']*MC3000['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MC3000['MC-Total-Cost']=(TTCost*MC3000['MC-CT-%'])/100
        MC3000['DP-Total-Cost']=DPONPART
        MC3000['MC|DP-Cost']=MC3000['MC-Total-Cost']+MC3000['DP-Total-Cost']
        MC3000['TT-P-Weight']=MC3000['Part-Weight']*MC3000['Good Parts']
        MC3000['TT-P-%']=((MC3000['Part-Weight']*MC3000['Good Parts'])/MC3000['TT-P-Weight'].sum())*100
        MC3000['Part-Cost']=(MC3000['MC|DP-Cost']*MC3000['TT-P-%'])/100
        MC3000['MC-Pcs-Cost']=(MC3000['Part-Cost']/MC3000['Part-Cavity'])/MC3000['Good Parts']
        MC3000=MC3000[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MC3000.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MC3000['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MC3000['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MC3000['Part-Cost'].sum()),2),'B')
        SUM3000=SUMMC[['Good Parts','MC-Pcs-Cost']]

        

        ############################# Part 3100 ############################
        try:
            DPONPART=DPMCPROD.loc['5612603100A','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MC3100=MCproddata[MCproddata['Part_No'].str.contains('5612603100A')]
        st.write('MC Cost Analysis:','5612603100A')
        MC3100['MC-CT-%']=((MC3100['Good Parts']*MC3100['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MC3100['MC-Total-Cost']=(TTCost*MC3100['MC-CT-%'])/100
        MC3100['DP-Total-Cost']=DPONPART
        MC3100['MC|DP-Cost']=MC3100['MC-Total-Cost']+MC3100['DP-Total-Cost']
        MC3100['TT-P-Weight']=MC3100['Part-Weight']*MC3100['Good Parts']
        MC3100['TT-P-%']=((MC3100['Part-Weight']*MC3100['Good Parts'])/MC3100['TT-P-Weight'].sum())*100
        MC3100['Part-Cost']=(MC3100['MC|DP-Cost']*MC3100['TT-P-%'])/100
        MC3100['MC-Pcs-Cost']=(MC3100['Part-Cost']/MC3100['Part-Cavity'])/MC3100['Good Parts']
        MC3100=MC3100[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MC3100.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MC3100['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MC3100['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MC3100['Part-Cost'].sum()),2),'B')
        SUM3100=SUMMC[['Good Parts','MC-Pcs-Cost']]

        

    ############################# Part 493C ############################
        try:
            DPONPART=DPMCPROD.loc['T96493CA','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MC493C=MCproddata[MCproddata['Part_No'].str.contains('T96493CA')]
        st.write('MC Cost Analysis:','T96493CA')
        MC493C['MC-CT-%']=((MC493C['Good Parts']*MC493C['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MC493C['MC-Total-Cost']=(TTCost*MC493C['MC-CT-%'])/100
        MC493C['DP-Total-Cost']=DPONPART
        MC493C['MC|DP-Cost']=MC493C['MC-Total-Cost']+MC493C['DP-Total-Cost']
        MC493C['TT-P-Weight']=MC493C['Part-Weight']*MC493C['Good Parts']
        MC493C['TT-P-%']=((MC493C['Part-Weight']*MC493C['Good Parts'])/MC493C['TT-P-Weight'].sum())*100
        MC493C['Part-Cost']=(MC493C['MC|DP-Cost']*MC493C['TT-P-%'])/100
        MC493C['MC-Pcs-Cost']=(MC493C['Part-Cost']/MC493C['Part-Cavity'])/MC493C['Good Parts']
        MC493C=MC493C[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MC493C.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MC493C['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MC493C['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MC493C['Part-Cost'].sum()),2),'B')
        SUM493C=SUMMC[['Good Parts','MC-Pcs-Cost']]

        

        ############################# Part1771 ############################

        try:
            DPONPART=DPMCPROD.loc['Z0021771A','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MC1771=MCproddata[MCproddata['Part_No'].str.contains('Z0021771A')]
        st.write('MC Cost Analysis:','Z0021771A')
        MC1771['MC-CT-%']=((MC1771['Good Parts']*MC1771['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MC1771['MC-Total-Cost']=(TTCost*MC1771['MC-CT-%'])/100
        MC1771['DP-Total-Cost']=DPONPART
        MC1771['MC|DP-Cost']=MC1771['MC-Total-Cost']+MC1771['DP-Total-Cost']
        MC1771['TT-P-Weight']=MC1771['Part-Weight']*MC1771['Good Parts']
        MC1771['TT-P-%']=((MC1771['Part-Weight']*MC1771['Good Parts'])/MC1771['TT-P-Weight'].sum())*100
        MC1771['Part-Cost']=(MC1771['MC|DP-Cost']*MC1771['TT-P-%'])/100
        MC1771['MC-Pcs-Cost']=(MC1771['Part-Cost']/MC1771['Part-Cavity'])/MC1771['Good Parts']
        MC1771=MC1771[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MC1771.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MC1771['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MC1771['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MC1771['Part-Cost'].sum()),2),'B')
        SUM1771=SUMMC[['Good Parts','MC-Pcs-Cost']]

        
            ############################# Part6496 ############################
        try:
            DPONPART=DPMCPROD.loc['T46496AA','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MC6496=MCproddata[MCproddata['Part_No'].str.contains('T46496AA')]
        st.write('MC Cost Analysis:','T46496AA')
        MC6496['MC-CT-%']=((MC6496['Good Parts']*MC6496['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MC6496['MC-Total-Cost']=(TTCost*MC6496['MC-CT-%'])/100
        MC6496['DP-Total-Cost']=DPONPART
        MC6496['MC|DP-Cost']=MC6496['MC-Total-Cost']+MC6496['DP-Total-Cost']
        MC6496['TT-P-Weight']=MC6496['Part-Weight']*MC6496['Good Parts']
        MC6496['TT-P-%']=((MC6496['Part-Weight']*MC6496['Good Parts'])/MC6496['TT-P-Weight'].sum())*100
        MC6496['Part-Cost']=(MC6496['MC|DP-Cost']*MC6496['TT-P-%'])/100
        MC6496['MC-Pcs-Cost']=(MC6496['Part-Cost']/MC6496['Part-Cavity'])/MC6496['Good Parts']
        MC6496=MC6496[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MC6496.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MC6496['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MC6496['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MC6496['Part-Cost'].sum()),2),'B')
        SUM6496=SUMMC[['Good Parts','MC-Pcs-Cost']]

        
        ############################# Part6497 ############################
        try:
            DPONPART=DPMCPROD.loc['T46497AA','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MC6497=MCproddata[MCproddata['Part_No'].str.contains('T46497AA')]
        st.write('MC Cost Analysis:','T46497AA')
        MC6497['MC-CT-%']=((MC6497['Good Parts']*MC6497['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MC6497['MC-Total-Cost']=(TTCost*MC6497['MC-CT-%'])/100
        MC6497['DP-Total-Cost']=DPONPART
        MC6497['MC|DP-Cost']=MC6497['MC-Total-Cost']+MC6497['DP-Total-Cost']
        MC6497['TT-P-Weight']=MC6497['Part-Weight']*MC6497['Good Parts']
        MC6497['TT-P-%']=((MC6497['Part-Weight']*MC6497['Good Parts'])/MC6497['TT-P-Weight'].sum())*100
        MC6497['Part-Cost']=(MC6497['MC|DP-Cost']*MC6497['TT-P-%'])/100
        MC6497['MC-Pcs-Cost']=(MC6497['Part-Cost']/MC6497['Part-Cavity'])/MC6497['Good Parts']
        MC6497=MC6497[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MC6497.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MC6497['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MC6497['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MC6497['Part-Cost'].sum()),2),'B')
        SUM6497=SUMMC[['Good Parts','MC-Pcs-Cost']]

            
        ############################# Part4900 ############################

        try:
            DPONPART=DPMCPROD.loc['5611514900A','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MC4900=MCproddata[MCproddata['Part_No'].str.contains('5611514900A')]
        st.write('MC Cost Analysis:','5611514900A')
        MC4900['MC-CT-%']=((MC4900['Good Parts']*MC4900['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MC4900['MC-Total-Cost']=(TTCost*MC4900['MC-CT-%'])/100
        MC4900['DP-Total-Cost']=DPONPART
        MC4900['MC|DP-Cost']=MC4900['MC-Total-Cost']+MC4900['DP-Total-Cost']
        MC4900['TT-P-Weight']=MC4900['Part-Weight']*MC4900['Good Parts']
        MC4900['TT-P-%']=((MC4900['Part-Weight']*MC4900['Good Parts'])/MC4900['TT-P-Weight'].sum())*100
        MC4900['Part-Cost']=(MC4900['MC|DP-Cost']*MC4900['TT-P-%'])/100
        MC4900['MC-Pcs-Cost']=(MC4900['Part-Cost']/MC4900['Part-Cavity'])/MC4900['Good Parts']
        MC4900=MC4900[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MC4900.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MC4900['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MC4900['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MC4900['Part-Cost'].sum()),2),'B')
        SUM4900=SUMMC[['Good Parts','MC-Pcs-Cost']]

        
            ############################# Part4400 ############################

        try:
            DPONPART=DPMCPROD.loc['5612604400A','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MC4400=MCproddata[MCproddata['Part_No'].str.contains('5612604400A')]
        st.write('MC Cost Analysis:','5612604400A')
        MC4400['MC-CT-%']=((MC4400['Good Parts']*MC4400['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MC4400['MC-Total-Cost']=(TTCost*MC4400['MC-CT-%'])/100
        MC4400['DP-Total-Cost']=DPONPART
        MC4400['MC|DP-Cost']=MC4400['MC-Total-Cost']+MC4400['DP-Total-Cost']
        MC4400['TT-P-Weight']=MC4400['Part-Weight']*MC4400['Good Parts']
        MC4400['TT-P-%']=((MC4400['Part-Weight']*MC4400['Good Parts'])/MC4400['TT-P-Weight'].sum())*100
        MC4400['Part-Cost']=(MC4400['MC|DP-Cost']*MC4400['TT-P-%'])/100
        MC4400['MC-Pcs-Cost']=(MC4400['Part-Cost']/MC4400['Part-Cavity'])/MC4400['Good Parts']
        MC4400=MC4400[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MC4400.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MC4400['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MC4400['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MC4400['Part-Cost'].sum()),2),'B')
        SUM4400=SUMMC[['Good Parts','MC-Pcs-Cost']]

        
            ############################# Part4500 ############################

        try:
            DPONPART=DPMCPROD.loc['5612604500A','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MC4500=MCproddata[MCproddata['Part_No'].str.contains('5612604500A')]
        st.write('MC Cost Analysis:','5612604500A')
        MC4500['MC-CT-%']=((MC4500['Good Parts']*MC4500['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MC4500['MC-Total-Cost']=(TTCost*MC4500['MC-CT-%'])/100
        MC4500['DP-Total-Cost']=DPONPART
        MC4500['MC|DP-Cost']=MC4500['MC-Total-Cost']+MC4500['DP-Total-Cost']
        MC4500['TT-P-Weight']=MC4500['Part-Weight']*MC4500['Good Parts']
        MC4500['TT-P-%']=((MC4500['Part-Weight']*MC4500['Good Parts'])/MC4500['TT-P-Weight'].sum())*100
        MC4500['Part-Cost']=(MC4500['MC|DP-Cost']*MC4500['TT-P-%'])/100
        MC4500['MC-Pcs-Cost']=(MC4500['Part-Cost']/MC4500['Part-Cavity'])/MC4500['Good Parts']
        MC4500=MC4500[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MC4500.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MC4500['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MC4500['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MC4500['Part-Cost'].sum()),2),'B')
        SUM4500=SUMMC[['Good Parts','MC-Pcs-Cost']]

            
            ############################# Part4900SIM ############################

        try:
            DPONPART=DPMCPROD.loc['5612604900A','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MC4900SIM=MCproddata[MCproddata['Part_No'].str.contains('5612604900A-SIM')]
        st.write('MC Cost Analysis:','5612604900A-SIM')
        MC4900SIM['MC-CT-%']=((MC4900SIM['Good Parts']*MC4900SIM['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MC4900SIM['MC-Total-Cost']=(TTCost*MC4900SIM['MC-CT-%'])/100
        MC4900SIM['DP-Total-Cost']=DPONPART
        MC4900SIM['MC|DP-Cost']=MC4900SIM['MC-Total-Cost']+MC4900SIM['DP-Total-Cost']
        MC4900SIM['TT-P-Weight']=MC4900SIM['Part-Weight']*MC4900SIM['Good Parts']
        MC4900SIM['TT-P-%']=((MC4900SIM['Part-Weight']*MC4900SIM['Good Parts'])/MC4900SIM['TT-P-Weight'].sum())*100
        MC4900SIM['Part-Cost']=(MC4900SIM['MC|DP-Cost']*MC4900SIM['TT-P-%'])/100
        MC4900SIM['MC-Pcs-Cost']=(MC4900SIM['Part-Cost']/MC4900SIM['Part-Cavity'])/MC4900SIM['Good Parts']
        MC4900SIM=MC4900SIM[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MC4900SIM.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MC4900SIM['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MC4900SIM['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MC4900SIM['Part-Cost'].sum()),2),'B')
        SUM4900SIM=SUMMC[['Good Parts','MC-Pcs-Cost']]
        

        
            ############################# Part331 ############################

        try:
            DPONPART=DPMCPROD.loc['220-00331','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MC331=MCproddata[MCproddata['Part_No'].str.contains('220-00331')]
        st.write('MC Cost Analysis:','220-00331')
        MC331['MC-CT-%']=((MC331['Good Parts']*MC331['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MC331['MC-Total-Cost']=(TTCost*MC331['MC-CT-%'])/100
        MC331['DP-Total-Cost']=DPONPART
        MC331['MC|DP-Cost']=MC331['MC-Total-Cost']+MC331['DP-Total-Cost']
        MC331['TT-P-Weight']=MC331['Part-Weight']*MC331['Good Parts']
        MC331['TT-P-%']=((MC331['Part-Weight']*MC331['Good Parts'])/MC331['TT-P-Weight'].sum())*100
        MC331['Part-Cost']=(MC331['MC|DP-Cost']*MC331['TT-P-%'])/100
        MC331['MC-Pcs-Cost']=(MC331['Part-Cost']/MC331['Part-Cavity'])/MC331['Good Parts']
        MC331=MC331[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MC331.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MC331['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MC331['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MC331['Part-Cost'].sum()),2),'B')
        SUM331=SUMMC[['Good Parts','MC-Pcs-Cost']]

                

            ############################# Part16_1 ############################

        try:
            DPONPART=DPMCPROD.loc['220-00016-1','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MC16_1=MCproddata[MCproddata['Part_No'].str.contains('220-00016-1')]
        st.write('MC Cost Analysis:','220-00016-1')
        MC16_1['MC-CT-%']=((MC16_1['Good Parts']*MC16_1['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MC16_1['MC-Total-Cost']=(TTCost*MC16_1['MC-CT-%'])/100
        MC16_1['DP-Total-Cost']=DPONPART
        MC16_1['MC|DP-Cost']=MC16_1['MC-Total-Cost']+MC16_1['DP-Total-Cost']
        MC16_1['TT-P-Weight']=MC16_1['Part-Weight']*MC16_1['Good Parts']
        MC16_1['TT-P-%']=((MC16_1['Part-Weight']*MC16_1['Good Parts'])/MC16_1['TT-P-Weight'].sum())*100
        MC16_1['Part-Cost']=(MC16_1['MC|DP-Cost']*MC16_1['TT-P-%'])/100
        MC16_1['MC-Pcs-Cost']=(MC16_1['Part-Cost']/MC16_1['Part-Cavity'])/MC16_1['Good Parts']
        MC16_1=MC16_1[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MC16_1.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MC16_1['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MC16_1['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MC16_1['Part-Cost'].sum()),2),'B')
        SUM16_1=SUMMC[['Good Parts','MC-Pcs-Cost']]

            
            ############################# Part16_2 ############################

        try:
            DPONPART=DPMCPROD.loc['220-00016-2','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MC16_2=MCproddata[MCproddata['Part_No'].str.contains('220-00016-2')]
        st.write('MC Cost Analysis:','220-00016-2')
        MC16_2['MC-CT-%']=((MC16_2['Good Parts']*MC16_2['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MC16_2['MC-Total-Cost']=(TTCost*MC16_2['MC-CT-%'])/100
        MC16_2['DP-Total-Cost']=DPONPART
        MC16_2['MC|DP-Cost']=MC16_2['MC-Total-Cost']+MC16_2['DP-Total-Cost']
        MC16_2['TT-P-Weight']=MC16_2['Part-Weight']*MC16_2['Good Parts']
        MC16_2['TT-P-%']=((MC16_2['Part-Weight']*MC16_2['Good Parts'])/MC16_2['TT-P-Weight'].sum())*100
        MC16_2['Part-Cost']=(MC16_2['MC|DP-Cost']*MC16_2['TT-P-%'])/100
        MC16_2['MC-Pcs-Cost']=(MC16_2['Part-Cost']/MC16_2['Part-Cavity'])/MC16_2['Good Parts']
        MC16_2=MC16_2[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MC16_2.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MC16_2['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MC16_2['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MC16_2['Part-Cost'].sum()),2),'B')
        SUM16_2=SUMMC[['Good Parts','MC-Pcs-Cost']]

        
        ############################# PartSleeve ############################


        try:
            DPONPART=DPMCPROD.loc['1050B375-RM','DP']
        except:
            DPMCPROD['DP']=0
        TTCost=MCCOSTALL-DPMCPROD['DP'].sum()
        MCSLEEVE=MCproddata[MCproddata['Part_No'].str.contains('1050B375-RM')]
        st.write('MC Cost Analysis:','1050B375-RM')
        MCSLEEVE['MC-CT-%']=((MCSLEEVE['Good Parts']*MCSLEEVE['MC-CT']).sum())/((MCCOST['Good Parts']*MCCOST['MC-CT']).sum())*100
        MCSLEEVE['MC-Total-Cost']=(TTCost*MCSLEEVE['MC-CT-%'])/100
        MCSLEEVE['DP-Total-Cost']=DPONPART
        MCSLEEVE['MC|DP-Cost']=MCSLEEVE['MC-Total-Cost']+MCSLEEVE['DP-Total-Cost']
        MCSLEEVE['TT-P-Weight']=MCSLEEVE['Part-Weight']*MCSLEEVE['Good Parts']
        MCSLEEVE['TT-P-%']=((MCSLEEVE['Part-Weight']*MCSLEEVE['Good Parts'])/MCSLEEVE['TT-P-Weight'].sum())*100
        MCSLEEVE['Part-Cost']=(MCSLEEVE['MC|DP-Cost']*MCSLEEVE['TT-P-%'])/100
        MCSLEEVE['MC-Pcs-Cost']=(MCSLEEVE['Part-Cost']/MCSLEEVE['Part-Cavity'])/MCSLEEVE['Good Parts']
        MCSLEEVE=MCSLEEVE[['Date','Part_No','MC-CT-%','MC-Total-Cost','DP-Total-Cost','MC|DP-Cost','Good Parts','TT-P-%','Part-Cost','MC-Pcs-Cost']]
        SUMMC=MCSLEEVE.groupby('Part_No').agg({'Good Parts':np.sum,'MC|DP-Cost':np.mean,'MC-Pcs-Cost':np.mean})
        SUMMC
        AvgMTCost=MCSLEEVE['MC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((MCSLEEVE['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(AvgMTCost,2),'B')
        st.write('Total MC Cost',round((MCSLEEVE['Part-Cost'].sum()),2),'B')
        SUMSLEEVE=SUMMC[['Good Parts','MC-Pcs-Cost']]
        SUMSLEEVE=SUMMC[['Good Parts','MC-Pcs-Cost']]

        st.warning('End of Machining Cost Report / The Cost Data show only part production')

    ############################ SUM MC Csot #####################################
        st.write('MC SUM Cost')
        SUMALLMC=pd.concat([SUM3000,SUM3100,SUM1771,SUM493C,SUM4900,SUM4900SIM,SUM16_1,SUM16_2,SUM331,SUM6496,SUM6497,SUMSLEEVE,SUM4400,SUM4500],axis=0)
        SUMALLMC
        st.write('Total Prod Pcs',round((SUMALLMC['Good Parts'].sum())),'Pcs')
        st.write('Avg MC Cost/Pcs',round(SUMALLMC['MC-Pcs-Cost'].mean(),2),'B')
        ############### To Excel ##################################
        MonthList={'Month':['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
        MAPNAME={'Jan':'\Machining-Rev031-Jan.xlsx','Feb':'\Machining-Rev031-Feb.xlsx','Mar':'\Machining-Rev031-Mar.xlsx','Apr':'\Machining-Rev031-Apr.xlsx',
        'May':'\Machining-Rev031-May.xlsx','Jun':'\Machining-Rev031-Jun.xlsx','Jul':'\Machining-Rev031-Jul.xlsx','Aug':'\Machining-Rev031-Aug.xlsx',
        'Sep':'\Machining-Rev031-Sep.xlsx','Oct':'\Machining-Rev031-Oct.xlsx','Nov':'\Machining-Rev031-Nov.xlsx','Dec':'\Machining-Rev031-Dec.xlsx'}
        MonthList=pd.DataFrame(MonthList)
        MonthList['File-Name']=MonthList['Month'].map(MAPNAME)
        MonthList=MonthList[MonthList['Month']==Minput]
        EXCELNAME=MonthList['File-Name'].to_string(index=False)
        st.write("---")
        Path=r"C:\Users\utaie\Desktop\Costing\DATA-Rev031"
        NAME=Path+EXCELNAME
        SUMALLMC.to_excel(NAME,sheet_name=Minput,engine='xlsxwriter')
        st.write('Data had export to excel:',NAME)
        st.success('End of MC Cost Analysis Report')
    
    ###################### QC-Cost ###################
    if Main=='QC-Cost':
        st.write('QC ALL Cost')
        QCCOST=COST[COST['ACC-CODE'].str.contains('616')]
        QCCOST=QCCOST[['ACC-CODE','ACC-NAME','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']]
        QCCOST.set_index('ACC-CODE',inplace=True)
        QCCOST[['ACC-NAME',Minput]]
        QCDATATT=QCCOST[Minput].sum()
        QCDATATT
        ##############################
        st.write('QC RM Cost')
        QCRM=QCCOST[~QCCOST['ACC-NAME'].str.contains('LAB')]
        QCRM[['ACC-NAME',Minput]]
        QCDATARM=QCRM[Minput].sum()
        QCDATARM
        ################
        st.write('QC MC Cost')
        QCMC=QCCOST[QCCOST['ACC-NAME'].str.contains('LAB')]
        QCMC[['ACC-NAME',Minput]]
        QCDATAMC=QCMC[Minput].sum()
        QCDATAMC
        ##########################
        st.write('RM-QC-Pcs Cost')
        QC=QCproddata
        QC['Good Parts']=QC["Sorting- Q'TY (Pcs)"]-QC['Total NG (Pcs)']
        QC=QC.fillna(0)
        QC['QC-Cost']=QCDATARM
        QC['QC-Pcs-Cost']=QC['QC-Cost']/(QC['Good Parts'].sum())
        QCSUMMRM=QC.groupby('Part_No').agg({'Good Parts':np.sum,'QC-Cost':np.mean,'Prices-Q1-22':np.mean,'Part-Weight':np.mean})
        QCSUMMRM['QC-Value-%']=(((QCSUMMRM['Part-Weight']*QCSUMMRM['Good Parts'])/(QCSUMMRM['Part-Weight']*QCSUMMRM['Good Parts']).sum())*100)
        QCSUMMRM['QC-SUM-Cost']=(QCSUMMRM['QC-Cost']*QCSUMMRM['QC-Value-%'])/100
        QCSUMMRM['QC-Pcs-Cost']=QCSUMMRM['QC-SUM-Cost']/QCSUMMRM['Good Parts']
        QCSUMMRM=QCSUMMRM[['Good Parts','QC-Cost','QC-Value-%','QC-SUM-Cost','QC-Pcs-Cost']]
        QCCost=QCSUMMRM['QC-Pcs-Cost']
        AvgQCCost=QCSUMMRM['QC-Pcs-Cost'].mean()
        QCRMDATA=pd.merge(QCSUMMRM,db[['Part_No','Part-Type']],on='Part_No',how='left')
        QCRMDATA=QCRMDATA[~QCRMDATA['Part-Type'].str.contains('MC')]
        QCRMDATA
        st.write('Total Prod Pcs',round((QCSUMMRM['Good Parts'].sum())),'Pcs')
        st.write('Avg QC Cost/Pcs',round(AvgQCCost,2),'B')
        st.write('Total QC Cost',round(QCSUMMRM['QC-SUM-Cost'].sum(),2),'B')
        ##################################
        st.write('MC-QC-Pcs Cost')
        QC=QCproddata[QCproddata['Part-Type'].str.contains('MC')|~QCproddata['Part-Type'].str.contains('SCK')]
        QC['Good Parts']=QC["Sorting- Q'TY (Pcs)"]-QC['Total NG (Pcs)']
        QC=QC.fillna(0)
        QC['QC-Cost']=QCDATAMC
        QC['QC-Pcs-Cost']=QC['QC-Cost']/(QC['Good Parts'].sum())
        QCSUMMMC=pd.merge(QC,db[['Part_No','QC-CT']],on='Part_No',how='left')
        QCSUMMMC=QC.groupby('Part_No').agg({'Good Parts':np.sum,'QC-Cost':np.mean,'Prices-Q1-22':np.mean,'Part-Weight':np.mean,'QC-CT':np.mean})
        QCSUMMMC['QC-Value-%']=(((QCSUMMMC['Part-Weight']*QCSUMMMC['Good Parts'])/(QCSUMMMC['Part-Weight']*QCSUMMMC['Good Parts']).sum())*100)*QCSUMMMC['QC-CT']
        QCSUMMMC['QC-SUM-Cost']=(QCSUMMMC['QC-Cost']*QCSUMMMC['QC-Value-%'])/100
        QCSUMMMC['QC-Pcs-Cost']=(QCSUMMMC['QC-SUM-Cost']/QCSUMMMC['Good Parts'])+QCSUMMRM['QC-Pcs-Cost']
        QCSUMMMC[['Good Parts','QC-Cost','QC-Value-%','QC-SUM-Cost','QC-Pcs-Cost']]
        QCCost=QCSUMMMC['QC-Pcs-Cost']
        AvgQCCost=QCSUMMMC['QC-Pcs-Cost'].mean()
        st.write('Total Prod Pcs',round((QCSUMMMC['Good Parts'].sum())),'Pcs')
        st.write('Avg QC Cost/Pcs',round(AvgQCCost,2),'B')
        st.write('Total QC Cost',round(QCSUMMMC['QC-SUM-Cost'].sum(),2),'B')
        ################### SUM QC Cost #####################################
        SUMQC=pd.concat([QCSUMMMC,QCSUMMRM],axis=0)
        SUMPCSCOSTQC=SUMQC[['Good Parts','QC-Pcs-Cost']]
        ############### To Excel ##################################
        MonthList={'Month':['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
        MAPNAME={'Jan':'\QC-Rev031-Jan.xlsx','Feb':'\QC-Rev031-Feb.xlsx','Mar':'\QC-Rev031-Mar.xlsx','Apr':'\QC-Rev031-Apr.xlsx',
        'May':'\QC-Rev031-May.xlsx','Jun':'\QC-Rev031-Jun.xlsx','Jul':'\QC-Rev031-Jul.xlsx','Aug':'\QC-Rev031-Aug.xlsx',
        'Sep':'\QC-Rev031-Sep.xlsx','Oct':'\QC-Rev031-Oct.xlsx','Nov':'\QC-Rev031-Nov.xlsx','Dec':'\QC-Rev031-Dec.xlsx'}
        MonthList=pd.DataFrame(MonthList)
        MonthList['File-Name']=MonthList['Month'].map(MAPNAME)
        MonthList=MonthList[MonthList['Month']==Minput]
        EXCELNAME=MonthList['File-Name'].to_string(index=False)
        st.write("---")
        Path=r"C:\Users\utaie\Desktop\Costing\DATA-Rev031"
        NAME=Path+EXCELNAME
        SUMPCSCOSTQC.to_excel(NAME,sheet_name=Minput,engine='xlsxwriter')
        st.write('Data had export to excel:',NAME)
        st.success('End of QC Cost Analysis Report')

        ############################### Outsource-Casting ################################################
    if Main=='SUB-Cost':
        st.write('**Outsource Cost**')
        ############################### RR ##########################
        SIMRR=pd.read_excel("2022 RR.xlsx",header=5)
        SIMRR['Part_No']=SIMRR['Part_No'].astype(str)
        SIMRR['วันที่']=SIMRR['วันที่'].astype(str)
        SIMRR=(SIMRR[SIMRR['วันที่'].str.contains(YMInput)])
        ################## SCK ######################################
        st.write('SCK-Cost')
        SCKRR=(SIMRR[SIMRR['ผู้จำหน่าย'].str.contains('ศรจินดา')])

        SCKRR['SCK-Pcs-Cost']=SCKRR['มูลค่าสินค้า']/SCKRR['จำนวน']
        SCKRR=SCKRR[['วันที่','Part_No','ผู้จำหน่าย','จำนวน','มูลค่าสินค้า','SCK-Pcs-Cost']]
        SCKRR=SCKRR.fillna(0)
        SCKRR
        SCK=SCKRR[['Part_No','SCK-Pcs-Cost']]
        st.write('Total Cost:',round(SCKRR['มูลค่าสินค้า'].sum(),2),'B')
        ################## KK ######################################
        st.write('KK-Cost')
        KKRR=(SIMRR[SIMRR['ผู้จำหน่าย'].str.contains('เค เค')])

        KKRR['KK-Pcs-Cost']=KKRR['มูลค่าสินค้า']/KKRR['จำนวน']
        KKRR=KKRR[['วันที่','Part_No','ผู้จำหน่าย','จำนวน','มูลค่าสินค้า','KK-Pcs-Cost']]
        KKRR=KKRR.fillna(0)
        KKRR
        KK=KKRR[['Part_No','KK-Pcs-Cost']]
        st.write('Total Cost:',round(KKRR['มูลค่าสินค้า'].sum(),2),'B')
        ################## NYS ######################################
        st.write('NYS-Cost')
        NYSRR=(SIMRR[SIMRR['ผู้จำหน่าย'].str.contains('เอ็น วาย เอส')])

        NYSRR['NYS-Pcs-Cost']=NYSRR['มูลค่าสินค้า']/NYSRR['จำนวน']
        NYSRR=NYSRR[['วันที่','Part_No','ผู้จำหน่าย','จำนวน','มูลค่าสินค้า','NYS-Pcs-Cost']]
        NYSRR=NYSRR.fillna(0)
        NYSRR
        NYS=NYSRR[['Part_No','NYS-Pcs-Cost']]
        st.write('Total Cost:',round(NYSRR['มูลค่าสินค้า'].sum(),2),'B')
        ################## Hot|Cool ######################################
        st.write('Hot|Cool-Cost')
        HotandCool=(SIMRR[SIMRR['ผู้จำหน่าย'].str.contains('ฮ็อทเเอนด์')])

        HotandCool['Hot|Cool-Pcs-Cost']=HotandCool['มูลค่าสินค้า']/HotandCool['จำนวน']
        HotandCool=HotandCool[['วันที่','Part_No','ผู้จำหน่าย','จำนวน','มูลค่าสินค้า','Hot|Cool-Pcs-Cost']]
        HotandCool=HotandCool.fillna(0)
        HotandCool
        st.write('Total Cost:',round(HotandCool['มูลค่าสินค้า'].sum(),2),'B')
        HC=HotandCool[['Part_No','Hot|Cool-Pcs-Cost']]
        st.write('KVS-Cost:',Minput)
        ######################## KVS ######################################
        KVS=(SIMRR[SIMRR['ผู้จำหน่าย'].str.contains('กฤษณะ')])
        KVSCOST=KVS['มูลค่าสินค้า'].sum()
        KVS['KVS-Pcs-Cost']=KVS['มูลค่าสินค้า']/KVS['จำนวน']
        KVS=KVS[['วันที่','Part_No','ผู้จำหน่าย','จำนวน','มูลค่าสินค้า','KVS-Pcs-Cost']]
        KVS=KVS.rename(columns={'จำนวน':'KVS-Pcs'})
        KVS=KVS.fillna(0)
        KVS
        st.write('Total Cost:',round(KVS['มูลค่าสินค้า'].sum(),2),'B')
        KVS=KVS[['Part_No','KVS-Pcs-Cost']]
        KVS.to_excel('KVS-Unit-Cost.xlsx')
        ################## Thai Inter ######################################
        st.write('Tin-Cost:',Minput)
        Tin=(SIMRR[SIMRR['ผู้จำหน่าย'].str.contains('ไทยอินเตอร์')])
        TinCOST=Tin['มูลค่าสินค้า'].sum()
        Tin['Tin-Pcs-Cost']=Tin['มูลค่าสินค้า']/Tin['จำนวน']
        Tin=Tin[['วันที่','Part_No','ผู้จำหน่าย','จำนวน','มูลค่าสินค้า','Tin-Pcs-Cost']]
        Tin=Tin.rename(columns={'จำนวน':'Tin-Pcs'})
        Tin=Tin.fillna(0)
        Tin
        st.write('Total Cost:',round(Tin['มูลค่าสินค้า'].sum(),2),'B')
        Tin=Tin[['Part_No','Tin-Pcs-Cost']]
        Tin.to_excel('ThaiInter-Unit-Cost.xlsx')
        ################## Coating ######################################
        st.write('Coating-Cost:',Minput)
        Coating=(SIMRR[SIMRR['ผู้จำหน่าย'].str.contains('ยูซีพี แอดวานซ์ ')])
        Coating['Coating-Pcs-Cost']=Coating['มูลค่าสินค้า']/Coating['จำนวน']
        Coating=Coating[['วันที่','Part_No','ผู้จำหน่าย','จำนวน','มูลค่าสินค้า','Coating-Pcs-Cost']]
        Coating=Coating.rename(columns={'จำนวน':'Coating-Pcs'})
        Coating=Coating.fillna(0)
        Coating
        st.write('Total Cost:',round(Coating['มูลค่าสินค้า'].sum(),2),'B')
        Coat=Coating[['Part_No','Coating-Pcs-Cost']]
        Coat.to_excel('Coating-Unit-Cost.xlsx')
        ######################## SUM RR ###############################
        SCK=SCK.groupby('Part_No').mean()
        KK=KK.groupby('Part_No').mean()
        NYS=NYS.groupby('Part_No').mean()
        HC=HC.groupby('Part_No').mean()
        KVS=KVS.groupby('Part_No').mean()
        Tin=Tin.groupby('Part_No').mean()
        Coating=Coating.groupby('Part_No').mean()
        SUMRR=pd.concat([SCK,KK,NYS,HC,KVS,Tin,Coating],axis=1)
        SUMRR=SUMRR.fillna(0)
        def color_endmonth(val):
            color = 'Black' if val <=0 else '#1D2B9D'
            return f'background-color: {color}'
        st.dataframe(SUMRR.style.applymap(color_endmonth, subset=['SCK-Pcs-Cost','KK-Pcs-Cost','NYS-Pcs-Cost',
        'Hot|Cool-Pcs-Cost','KVS-Pcs-Cost','Tin-Pcs-Cost','Coating-Pcs-Cost']))
        
        SUMRR.to_excel('SUMRR-Unit-Cost.xlsx')
####################### Unit Cost Analysis Apr ########################################################
if MENU=='MENU Unit Cost' :
    if Main2=='Apr':
        st.write('**Unit Cost Analysis-Apr**')
        MT=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Melting-Rev031-Apr.xlsx')
        MAT=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Material-Rev031-Apr.xlsx')
        DC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Diecasting-Rev031-Apr.xlsx')
        FN=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Finishing-Rev031-Apr.xlsx')
        SB=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\ShotBlasting-Rev031-Apr.xlsx')
        MC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Machining-Rev031-Apr.xlsx')
        QC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\QC-Rev031-Apr.xlsx')

        UnitCostBF=pd.merge(DC,pd.merge(MT,MAT,on='Part_No'),on='Part_No',how='left',suffixes={'',"_ADD"})
        UnitCostBF['BF-Pcs-Cost']=UnitCostBF[['Part_No','MT-Pcs-Cost','Mat-Pcs-Cost','DCPROD-Pcs-Cost']].sum(axis=1)
        # UnitCostBF[['Part_No','MT-Pcs-Cost','Mat-Pcs-Cost','DCPROD-Pcs-Cost','BF-Pcs-Cost']]
        BFCOST=UnitCostBF[['Part_No','BF-Pcs-Cost']]
        BFCOST.set_index('Part_No',inplace=True)
        FNCOST=FN[['Part_No','FN-Pcs-Cost']]
        FNCOST.set_index('Part_No',inplace=True)
        SBCOST=SB[['Part_No','SB-Pcs-Cost']]
        SBCOST.set_index('Part_No',inplace=True)
        MCCOST=MC[['Part_No','MC-Pcs-Cost']]
        MCCOST.set_index('Part_No',inplace=True)
        QCCOST=QC[['Part_No','QC-Pcs-Cost']]
        QCCOST.set_index('Part_No',inplace=True)
        BSCOST=pd.concat([BFCOST,FNCOST.reindex(BFCOST.index)],axis=1)
        BMCOST=pd.concat([BSCOST,SBCOST.reindex(BSCOST.index)],axis=1)
        FG0COST=pd.concat([BMCOST,MCCOST.reindex(BMCOST.index)],axis=1)
        FGCOST=pd.merge(FG0COST,QCCOST,on='Part_No')
        FGCOST=FGCOST.fillna(0)
        FGCOST['FG1-Pcs-Cost']=FGCOST.sum(axis=1)
        FGCOST=pd.merge(FGCOST,db[['Part_No','Prices-Q2-22']],on='Part_No',how='left')
        FGCOST['BL-(Bht)']=FGCOST['Prices-Q2-22']-FGCOST['FG1-Pcs-Cost']
        FGCOST['BL-(%)']=(FGCOST['BL-(Bht)']/FGCOST['Prices-Q2-22'])*100
        FGCOST=FGCOST.groupby('Part_No').mean()
        FGCOST
        st.write('Balance Unit Cost %:',round(FGCOST['BL-(%)'].mean(),2),'%')
    ####################################################
    ############## Sales Apr ###########################
        st.write('Sales: Apr-2022')
        filtApr=(Sales['วันที่'].str.contains('2022-04'))
        SalesM=Sales.loc[filtApr]
        SalesM.rename(columns={'รหัสสินค้า':'Part_No', 'จำนวน':'Quty', 'มูลค่าสินค้า':'AMT'},inplace=True)
        SalesM.set_index('Part_No',inplace=True)
    # ################### SUM Cost ##################################
        SUMCost=pd.read_excel('Apr-FG1-Final.xlsx')
        SUMCost.set_index('Part_No',inplace=True)
        SalesM=SalesM.fillna(0)
        SalesM=pd.merge(SalesM,SUMCost['FG1-Sales'],left_index=True,right_index=True,how='left')
        SalesM=SalesM[['Quty','AMT','FG1-Sales']].groupby('Part_No').agg({'Quty':'sum','AMT':'sum','FG1-Sales':'mean'})
        SalesM['Sales-Cost']=SalesM['Quty']*SalesM['FG1-Sales']
        SalesM['Sales-BL']=SalesM['AMT']-SalesM['Sales-Cost']
        SalesM['BL-%']=(SalesM['Sales-BL']/SalesM['AMT'])*100
        SalesM
        st.write('Sales AMT:',round(SalesM['AMT'].sum(),2),'B')
        st.write('Sales Balance:',round(SalesM['Sales-BL'].sum(),2),'B')
        st.write('Sales Balance %:',round(SalesM['BL-%'].mean(),2),'B')
        ####################### End Stock #########################################
        st.write('SUM Stock End Month')
        Stock=pd.read_excel('Stock-End-Month-2022.xlsx',sheet_name='Apr')
        Stock[['Part_No','TOTAL - BM' ,'TOTAL - FG0','TOTAL - FG1']]

    ##################################################################
    ####################### Unit Cost Analysis May ########################################################
    if Main2=='May':
        st.write('**Unit Cost Analysis-May**')
        MT=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Melting-Rev031-May.xlsx')
        MAT=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Material-Rev031-May.xlsx')
        DC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Diecasting-Rev031-May.xlsx')
        FN=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Finishing-Rev031-May.xlsx')
        SB=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\ShotBlasting-Rev031-May.xlsx')
        MC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Machining-Rev031-May.xlsx')
        QC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\QC-Rev031-May.xlsx')

        UnitCostBF=pd.merge(DC,pd.merge(MT,MAT,on='Part_No'),on='Part_No',how='left',suffixes={'',"_ADD"})
        UnitCostBF['BF-Pcs-Cost']=UnitCostBF[['Part_No','MT-Pcs-Cost','Mat-Pcs-Cost','DCPROD-Pcs-Cost']].sum(axis=1)
        # UnitCostBF[['Part_No','MT-Pcs-Cost','Mat-Pcs-Cost','DCPROD-Pcs-Cost','BF-Pcs-Cost']]
        BFCOST=UnitCostBF[['Part_No','BF-Pcs-Cost']]
        BFCOST.set_index('Part_No',inplace=True)
        FNCOST=FN[['Part_No','FN-Pcs-Cost']]
        FNCOST.set_index('Part_No',inplace=True)
        SBCOST=SB[['Part_No','SB-Pcs-Cost']]
        SBCOST.set_index('Part_No',inplace=True)
        MCCOST=MC[['Part_No','MC-Pcs-Cost']]
        MCCOST.set_index('Part_No',inplace=True)
        QCCOST=QC[['Part_No','QC-Pcs-Cost']]
        QCCOST.set_index('Part_No',inplace=True)
        BSCOST=pd.concat([BFCOST,FNCOST.reindex(BFCOST.index)],axis=1)
        BMCOST=pd.concat([BSCOST,SBCOST.reindex(BSCOST.index)],axis=1)
        FG0COST=pd.concat([BMCOST,MCCOST.reindex(BMCOST.index)],axis=1)
        FGCOST=pd.merge(FG0COST,QCCOST,on='Part_No')
        FGCOST=FGCOST.fillna(0)
        FGCOST['FG1-Pcs-Cost']=FGCOST.sum(axis=1)
        FGCOST=pd.merge(FGCOST,db[['Part_No','Prices-Q2-22']],on='Part_No',how='left')
        FGCOST['BL-(Bht)']=FGCOST['Prices-Q2-22']-FGCOST['FG1-Pcs-Cost']
        FGCOST['BL-(%)']=(FGCOST['BL-(Bht)']/FGCOST['Prices-Q2-22'])*100
        FGCOST=FGCOST.groupby('Part_No').mean()
        FGCOST
        st.write('Balance Unit Cost %:',round(FGCOST['BL-(%)'].mean(),2),'%')
        ############## Sales May ###########################
        st.write('Sales: May-2022')
        filtMay=(Sales['วันที่'].str.contains('2022-05'))
        SalesM=Sales.loc[filtMay]
        SalesM.rename(columns={'รหัสสินค้า':'Part_No', 'จำนวน':'Quty', 'มูลค่าสินค้า':'AMT'},inplace=True)
        SalesM.set_index('Part_No',inplace=True)
    # ################### SUM Cost ##################################
        SUMCost=pd.read_excel('May-FG1-Final.xlsx')
        SUMCost.set_index('Part_No',inplace=True)
        SalesM=SalesM.fillna(0)
        SalesM=pd.merge(SalesM,SUMCost['FG1-Sales'],left_index=True,right_index=True,how='left')
        SalesM=SalesM[['Quty','AMT','FG1-Sales']].groupby('Part_No').agg({'Quty':'sum','AMT':'sum','FG1-Sales':'mean'})
        SalesM=SalesM.fillna(0)
        SalesM['Sales-Cost']=SalesM['Quty']*SalesM['FG1-Sales']
        SalesM['Sales-BL']=SalesM['AMT']-SalesM['Sales-Cost']
        SalesM['BL-%']=(SalesM['Sales-BL']/SalesM['AMT'])*100
        SalesM
        st.write('Sales AMT:',round(SalesM['AMT'].sum(),2),'B')
        st.write('Sales Balance:',round(SalesM['Sales-BL'].sum(),2),'B')
        st.write('Sales Balance %:',round(SalesM['BL-%'].mean(),2),'B')
        ####################### End Stock #########################################
        st.write('SUM Stock End Month')
        Stock=pd.read_excel('Stock-End-Month-2022.xlsx',sheet_name='May')
        Stock[['Part_No','TOTAL - BM' ,'TOTAL - FG0','TOTAL - FG1']]
    ####################### Unit Cost Analysis Jun ########################################################
    if Main2=='Jun':
        st.write('**Unit Cost Analysis-Jun**')
        MT=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Melting-Rev031-Jun.xlsx')
        MAT=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Material-Rev031-Jun.xlsx')
        DC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Diecasting-Rev031-Jun.xlsx')
        FN=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Finishing-Rev031-Jun.xlsx')
        SB=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\ShotBlasting-Rev031-Jun.xlsx')
        MC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Machining-Rev031-Jun.xlsx')
        QC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\QC-Rev031-Jun.xlsx')

        UnitCostBF=pd.merge(DC,pd.merge(MT,MAT,on='Part_No'),on='Part_No',how='left',suffixes={'',"_ADD"})
        UnitCostBF['BF-Pcs-Cost']=UnitCostBF[['Part_No','MT-Pcs-Cost','Mat-Pcs-Cost','DCPROD-Pcs-Cost']].sum(axis=1)
        # UnitCostBF[['Part_No','MT-Pcs-Cost','Mat-Pcs-Cost','DCPROD-Pcs-Cost','BF-Pcs-Cost']]
        BFCOST=UnitCostBF[['Part_No','BF-Pcs-Cost']]
        BFCOST.set_index('Part_No',inplace=True)
        FNCOST=FN[['Part_No','FN-Pcs-Cost']]
        FNCOST.set_index('Part_No',inplace=True)
        SBCOST=SB[['Part_No','SB-Pcs-Cost']]
        SBCOST.set_index('Part_No',inplace=True)
        MCCOST=MC[['Part_No','MC-Pcs-Cost']]
        MCCOST.set_index('Part_No',inplace=True)
        QCCOST=QC[['Part_No','QC-Pcs-Cost']]
        QCCOST.set_index('Part_No',inplace=True)
        BSCOST=pd.concat([BFCOST,FNCOST.reindex(BFCOST.index)],axis=1)
        BMCOST=pd.concat([BSCOST,SBCOST.reindex(BSCOST.index)],axis=1)
        FG0COST=pd.concat([BMCOST,MCCOST.reindex(BMCOST.index)],axis=1)
        FGCOST=pd.merge(FG0COST,QCCOST,on='Part_No')
        FGCOST=FGCOST.fillna(0)
        FGCOST['FG1-Pcs-Cost']=FGCOST.sum(axis=1)
        FGCOST=pd.merge(FGCOST,db[['Part_No','Prices-Q2-22']],on='Part_No',how='left')
        FGCOST['BL-(Bht)']=FGCOST['Prices-Q2-22']-FGCOST['FG1-Pcs-Cost']
        FGCOST['BL-(%)']=(FGCOST['BL-(Bht)']/FGCOST['Prices-Q2-22'])*100
        FGCOST=FGCOST.groupby('Part_No').mean()
        FGCOST
        st.write('Balance Unit Cost %:',round(FGCOST['BL-(%)'].mean(),2),'%')
        ############## Sales Jun ###########################
        st.write('Sales: Jun-2022')
        filtJun=(Sales['วันที่'].str.contains('2022-06'))
        SalesM=Sales.loc[filtJun]
        SalesM.rename(columns={'รหัสสินค้า':'Part_No', 'จำนวน':'Quty', 'มูลค่าสินค้า':'AMT'},inplace=True)
        SalesM.set_index('Part_No',inplace=True)
    # ################### SUM Cost ##################################
        SUMCost=pd.read_excel('Jun-FG1-Final.xlsx')
        SUMCost.set_index('Part_No',inplace=True)
        SalesM=SalesM.fillna(0)
        SalesM=pd.merge(SalesM,SUMCost['FG1-Sales'],left_index=True,right_index=True,how='left')
        SalesM=SalesM[['Quty','AMT','FG1-Sales']].groupby('Part_No').agg({'Quty':'sum','AMT':'sum','FG1-Sales':'mean'})
        SalesM=SalesM.fillna(0)
        # SalesM['FG1-Sales1']=SalesM[['FG1-Pcs-Cost','FG1-Sales']].where(SalesM['FG1-Pcs-Cost']==0).sum(axis=1)
        # SalesM['FG1-Sales2']=SalesM[['FG1-Pcs-Cost','FG1-Sales1']].sum(axis=1)

        SalesM['Sales-Cost']=SalesM['Quty']*SalesM['FG1-Sales']
        SalesM['Sales-BL']=SalesM['AMT']-SalesM['Sales-Cost']
        SalesM['BL-%']=(SalesM['Sales-BL']/SalesM['AMT'])*100
        SalesM
        st.write('Sales AMT:',round(SalesM['AMT'].sum(),2),'B')
        st.write('Sales Balance:',round(SalesM['Sales-BL'].sum(),2),'B')
        st.write('Sales Balance %:',round(SalesM['BL-%'].mean(),2),'B')
        ####################### End Stock #########################################
        st.write('SUM Stock End Month')
        Stock=pd.read_excel('Stock-End-Month-2022.xlsx',sheet_name='Jun')
        Stock[['Part_No','TOTAL - BM' ,'TOTAL - FG0','TOTAL - FG1']]
    ####################### Unit Cost Analysis Jul ########################################################

    if Main2=='Jul':
        st.write('**Unit Cost Analysis-Jul**')
        MT=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Melting-Rev031-Jul.xlsx')
        MAT=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Material-Rev031-Jul.xlsx')
        DC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Diecasting-Rev031-Jul.xlsx')
        FN=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Finishing-Rev031-Jul.xlsx')
        SB=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\ShotBlasting-Rev031-Jul.xlsx')
        MC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Machining-Rev031-Jul.xlsx')
        QC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\QC-Rev031-Jul.xlsx')

        UnitCostBF=pd.merge(DC,pd.merge(MT,MAT,on='Part_No'),on='Part_No',how='left',suffixes={'',"_ADD"})
        UnitCostBF['BF-Pcs-Cost']=UnitCostBF[['Part_No','MT-Pcs-Cost','Mat-Pcs-Cost','DCPROD-Pcs-Cost']].sum(axis=1)
        # UnitCostBF[['Part_No','MT-Pcs-Cost','Mat-Pcs-Cost','DCPROD-Pcs-Cost','BF-Pcs-Cost']]
        BFCOST=UnitCostBF[['Part_No','BF-Pcs-Cost']]
        BFCOST.set_index('Part_No',inplace=True)
        FNCOST=FN[['Part_No','FN-Pcs-Cost']]
        FNCOST.set_index('Part_No',inplace=True)
        SBCOST=SB[['Part_No','SB-Pcs-Cost']]
        SBCOST.set_index('Part_No',inplace=True)
        MCCOST=MC[['Part_No','MC-Pcs-Cost']]
        MCCOST.set_index('Part_No',inplace=True)
        QCCOST=QC[['Part_No','QC-Pcs-Cost']]
        QCCOST.set_index('Part_No',inplace=True)
        BSCOST=pd.concat([BFCOST,FNCOST.reindex(BFCOST.index)],axis=1)
        BMCOST=pd.concat([BSCOST,SBCOST.reindex(BSCOST.index)],axis=1)
        FG0COST=pd.concat([BMCOST,MCCOST.reindex(BMCOST.index)],axis=1)
        FGCOST=pd.merge(FG0COST,QCCOST,on='Part_No')
        FGCOST=FGCOST.fillna(0)
        FGCOST['FG1-Pcs-Cost']=FGCOST.sum(axis=1)
        FGCOST=pd.merge(FGCOST,db[['Part_No','Prices-Q3-22']],on='Part_No',how='left')
        FGCOST['BL-(Bht)']=FGCOST['Prices-Q3-22']-FGCOST['FG1-Pcs-Cost']
        FGCOST['BL-(%)']=(FGCOST['BL-(Bht)']/FGCOST['Prices-Q3-22'])*100
        FGCOST=FGCOST.groupby('Part_No').mean()
        FGCOST
        st.write('Balance Unit Cost %:',round(FGCOST['BL-(%)'].mean(),2),'%')
        ############## Sales Jul ###########################
        st.write('Sales: Jul-2022')
        filtJul=(Sales['วันที่'].str.contains('2022-07'))
        SalesM=Sales.loc[filtJul]
        SalesM.rename(columns={'รหัสสินค้า':'Part_No', 'จำนวน':'Quty', 'มูลค่าสินค้า':'AMT'},inplace=True)
        SalesM.set_index('Part_No',inplace=True)
    # ################### SUM Cost ##################################
        SUMCost=pd.read_excel('Jul-FG1-Final.xlsx')
        SUMCost.set_index('Part_No',inplace=True)
        SalesM=SalesM.fillna(0)
        SalesM=pd.merge(SalesM,SUMCost['FG1-Sales'],left_index=True,right_index=True,how='left')
        SalesM=SalesM[['Quty','AMT','FG1-Sales']].groupby('Part_No').agg({'Quty':'sum','AMT':'sum','FG1-Sales':'mean'})
        SalesM=SalesM.fillna(0)
        # SalesM['FG1-Sales1']=SalesM[['FG1-Pcs-Cost','FG1-Sales']].where(SalesM['FG1-Pcs-Cost']==0).sum(axis=1)
        # SalesM['FG1-Sales2']=SalesM[['FG1-Pcs-Cost','FG1-Sales1']].sum(axis=1)

        SalesM['Sales-Cost']=SalesM['Quty']*SalesM['FG1-Sales']
        SalesM['Sales-BL']=SalesM['AMT']-SalesM['Sales-Cost']
        SalesM['BL-%']=(SalesM['Sales-BL']/SalesM['AMT'])*100
        SalesM
        st.write('Sales AMT:',round(SalesM['AMT'].sum(),2),'B')
        st.write('Sales Balance:',round(SalesM['Sales-BL'].sum(),2),'B')
        st.write('Sales Balance %:',round(SalesM['BL-%'].mean(),2),'B')
        ####################### End Stock #########################################
        st.write('SUM Stock End Month')
        Stock=pd.read_excel('Stock-End-Month-2022.xlsx',sheet_name='Jul')
        Stock[['Part_No','TOTAL - BM' ,'TOTAL - FG0','TOTAL - FG1']]
        
    ####################### Unit Cost Analysis Aug ########################################################
    if Main2=='Aug':
        st.write('**Unit Cost Analysis-Aug**')
        MT=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Melting-Rev031-Aug.xlsx')
        MAT=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Material-Rev031-Aug.xlsx')
        DC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Diecasting-Rev031-Aug.xlsx')
        FN=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Finishing-Rev031-Aug.xlsx')
        SB=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\ShotBlasting-Rev031-Aug.xlsx')
        MC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Machining-Rev031-Aug.xlsx')
        QC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\QC-Rev031-Aug.xlsx')

        UnitCostBF=pd.merge(DC,pd.merge(MT,MAT,on='Part_No'),on='Part_No',how='left',suffixes={'',"_ADD"})
        UnitCostBF['BF-Pcs-Cost']=UnitCostBF[['Part_No','MT-Pcs-Cost','Mat-Pcs-Cost','DCPROD-Pcs-Cost']].sum(axis=1)
        # UnitCostBF[['Part_No','MT-Pcs-Cost','Mat-Pcs-Cost','DCPROD-Pcs-Cost','BF-Pcs-Cost']]
        BFCOST=UnitCostBF[['Part_No','BF-Pcs-Cost']]
        BFCOST.set_index('Part_No',inplace=True)
        FNCOST=FN[['Part_No','FN-Pcs-Cost']]
        FNCOST.set_index('Part_No',inplace=True)
        SBCOST=SB[['Part_No','SB-Pcs-Cost']]
        SBCOST.set_index('Part_No',inplace=True)
        MCCOST=MC[['Part_No','MC-Pcs-Cost']]
        MCCOST.set_index('Part_No',inplace=True)
        QCCOST=QC[['Part_No','QC-Pcs-Cost']]
        QCCOST.set_index('Part_No',inplace=True)
        BSCOST=pd.concat([BFCOST,FNCOST.reindex(BFCOST.index)],axis=1)
        BMCOST=pd.concat([BSCOST,SBCOST.reindex(BSCOST.index)],axis=1)
        FG0COST=pd.concat([BMCOST,MCCOST.reindex(BMCOST.index)],axis=1)
        FGCOST=pd.merge(FG0COST,QCCOST,on='Part_No')
        FGCOST=FGCOST.fillna(0)
        FGCOST['FG1-Pcs-Cost']=FGCOST.sum(axis=1)
        FGCOST=pd.merge(FGCOST,db[['Part_No','Prices-Q3-22']],on='Part_No',how='left')
        FGCOST['BL-(Bht)']=FGCOST['Prices-Q3-22']-FGCOST['FG1-Pcs-Cost']
        FGCOST['BL-(%)']=(FGCOST['BL-(Bht)']/FGCOST['Prices-Q3-22'])*100
        FGCOST=FGCOST.groupby('Part_No').mean()
        FGCOST
        st.write('Balance Unit Cost %:',round(FGCOST['BL-(%)'].mean(),2),'%')
        ############## Sales Aug ###########################
        st.write('Sales: Aug-2022')
        filtAug=(Sales['วันที่'].str.contains('2022-08'))
        SalesM=Sales.loc[filtAug]
        SalesM.rename(columns={'รหัสสินค้า':'Part_No', 'จำนวน':'Quty', 'มูลค่าสินค้า':'AMT'},inplace=True)
        SalesM.set_index('Part_No',inplace=True)
    # ################### SUM Cost ##################################
        SUMCost=pd.read_excel('Aug-FG1-Final.xlsx')
        SUMCost.set_index('Part_No',inplace=True)
        SalesM=SalesM.fillna(0)
        SalesM=pd.merge(SalesM,SUMCost['FG1-Sales'],left_index=True,right_index=True,how='left')
        SalesM=SalesM[['Quty','AMT','FG1-Sales']].groupby('Part_No').agg({'Quty':'sum','AMT':'sum','FG1-Sales':'mean'})
        SalesM=SalesM.fillna(0)
        # SalesM['FG1-Sales1']=SalesM[['FG1-Pcs-Cost','FG1-Sales']].where(SalesM['FG1-Pcs-Cost']==0).sum(axis=1)
        # SalesM['FG1-Sales2']=SalesM[['FG1-Pcs-Cost','FG1-Sales1']].sum(axis=1)

        SalesM['Sales-Cost']=SalesM['Quty']*SalesM['FG1-Sales']
        SalesM['Sales-BL']=SalesM['AMT']-SalesM['Sales-Cost']
        SalesM['BL-%']=(SalesM['Sales-BL']/SalesM['AMT'])*100
        SalesM
        st.write('Sales AMT:',round(SalesM['AMT'].sum(),2),'B')
        st.write('Sales Balance:',round(SalesM['Sales-BL'].sum(),2),'B')
        st.write('Sales Balance %:',round(SalesM['BL-%'].mean(),2),'B')
####################### End Stock #########################################
        st.write('SUM Stock End Month')
        Stock=pd.read_excel('Stock-End-Month-2022.xlsx',sheet_name='Aug')
        Stock[['Part_No','TOTAL - BM' ,'TOTAL - FG0','TOTAL - FG1']]
    ####################### Unit Cost Analysis Sep ########################################################
    if Main2=='Sep':
        Stock=pd.read_excel('Stock-End-Month-2022.xlsx',sheet_name='Sep')
        DCproddata=pd.read_excel('DC-Report-Jan-Sep-2022.xlsx',sheet_name='Sep')
        DCproddata.rename(columns={'Good-Pcs':'Good Parts'},inplace=True)
        DCPCS=DCproddata[['Part_No','Good Parts']]
        DCPCS
        st.write('**Unit Cost Analysis-Sep**')
        MT=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Melting-Rev031-Sep.xlsx')
        MAT=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Material-Rev031-Sep.xlsx')
        DC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Diecasting-Rev031-Sep.xlsx')
        FN=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Finishing-Rev031-Sep.xlsx')
        SB=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\ShotBlasting-Rev031-Sep.xlsx')
        MC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\Machining-Rev031-Sep.xlsx')
        QC=pd.read_excel(r'C:\Users\utaie\Desktop\Costing\DATA-Rev031\QC-Rev031-Sep.xlsx')

        UnitCostBF=pd.merge(Stock,pd.merge(DC,pd.merge(MT,MAT,on='Part_No'),on='Part_No',how='left',suffixes={'',"_ADD"}),on='Part_No',how='right')

        UnitCostBF[['Part_No','Good Parts','TOTAL - BM','TOTAL - FG0','TOTAL - FG1']]
        UnitCostBF['BF-Pcs-Cost']=UnitCostBF[['Part_No','MT-Pcs-Cost','Mat-Pcs-Cost','DCPROD-Pcs-Cost']].sum(axis=1)
        # UnitCostBF[['Part_No','MT-Pcs-Cost','Mat-Pcs-Cost','DCPROD-Pcs-Cost','BF-Pcs-Cost']]
        BFCOST=UnitCostBF[['Part_No','BF-Pcs-Cost']]
        BFCOST.set_index('Part_No',inplace=True)
        FNCOST=FN[['Part_No','FN-Pcs-Cost']]
        FNCOST.set_index('Part_No',inplace=True)
        SBCOST=SB[['Part_No','SB-Pcs-Cost']]
        SBCOST.set_index('Part_No',inplace=True)
        MCCOST=MC[['Part_No','MC-Pcs-Cost']]
        MCCOST.set_index('Part_No',inplace=True)
        QCCOST=QC[['Part_No','QC-Pcs-Cost']]
        QCCOST.set_index('Part_No',inplace=True)
        BSCOST=pd.concat([BFCOST,FNCOST.reindex(BFCOST.index)],axis=1)
        BMCOST=pd.concat([BSCOST,SBCOST.reindex(BSCOST.index)],axis=1)
        FG0COST=pd.concat([BMCOST,MCCOST.reindex(BMCOST.index)],axis=1)
        FGCOST=pd.merge(FG0COST,QCCOST,on='Part_No')
        FGCOST=FGCOST.fillna(0)
        FGCOST['FG1-Pcs-Cost']=FGCOST.sum(axis=1)
        FGCOST=pd.merge(FGCOST,db[['Part_No','Prices-Q3-22']],on='Part_No',how='left')
        FGCOST['BL-(Bht)']=FGCOST['Prices-Q3-22']-FGCOST['FG1-Pcs-Cost']
        FGCOST['BL-(%)']=(FGCOST['BL-(Bht)']/FGCOST['Prices-Q3-22'])*100
        # FGCOST.set_index('Part_No',inplace=True)
        FGCOST=FGCOST.groupby('Part_No').mean()
        FGCOST
        st.write('Balance Unit Cost %:',round(FGCOST['BL-(%)'].mean(),2),'%')
        ############## Sales Sep ###########################
        st.write('Sales: Sep-2022')
        filtSep=(Sales['วันที่'].str.contains('2022-09'))
        SalesM=Sales.loc[filtSep]
        SalesM.rename(columns={'รหัสสินค้า':'Part_No', 'จำนวน':'Quty', 'มูลค่าสินค้า':'AMT'},inplace=True)
        SalesM.set_index('Part_No',inplace=True)
    # ################### SUM Cost ##################################
        SUMCost=pd.read_excel('Sep-FG1-Final.xlsx')
        SUMCost.set_index('Part_No',inplace=True)
        SalesM=SalesM.fillna(0)
        SalesM=pd.merge(SalesM,SUMCost['FG1-Sales'],left_index=True,right_index=True,how='left')
        SalesM=SalesM[['Quty','AMT','FG1-Sales']].groupby('Part_No').agg({'Quty':'sum','AMT':'sum','FG1-Sales':'mean'})
        SalesM=SalesM.fillna(0)
        # SalesM['FG1-Sales1']=SalesM[['FG1-Pcs-Cost','FG1-Sales']].where(SalesM['FG1-Pcs-Cost']==0).sum(axis=1)
        # SalesM['FG1-Sales2']=SalesM[['FG1-Pcs-Cost','FG1-Sales1']].sum(axis=1)

        SalesM['Sales-Cost']=SalesM['Quty']*SalesM['FG1-Sales']
        SalesM['Sales-BL']=SalesM['AMT']-SalesM['Sales-Cost']
        SalesM['BL-%']=(SalesM['Sales-BL']/SalesM['AMT'])*100
        SalesM
        st.write('Sales AMT:',round(SalesM['AMT'].sum(),2),'B')
        st.write('Sales Balance:',round(SalesM['Sales-BL'].sum(),2),'B')
        st.write('Sales Balance %:',round(SalesM['BL-%'].mean(),2),'B')
        ################### Ending Stock ####################################
        st.write('SUM Stock End Month')
        
        Stock[['Part_No','TOTAL - BM' ,'TOTAL - FG0','TOTAL - FG1']]
        SalesDF=SalesM['Quty']
        SalesDF
        StockDF=Stock[['Part_No','TOTAL - FG1']]
        FGCOST['FG1-Pcs-Cost']
        StockDF.set_index('Part_No',inplace=True)
        StockDF
        DCR=DCproddata[['Part_No','Good Parts']]
        DCR
        DCR.set_index('Part_No',inplace=True)
        BLDF=pd.merge(DCR,pd.merge(SalesDF,StockDF,left_on='Part_No',right_on='Part_No'),left_on='Part_No',right_on='Part_No')
        BLDF