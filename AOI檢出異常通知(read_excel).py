
# coding: utf-8

# In[35]:


import cx_Oracle
import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import seaborn as sns
import zeep
import base64
from io import BytesIO
import time
import datetime
from pandas.core.frame import DataFrame


# In[36]:


conn = cx_Oracle.connect("ML5DI1", "ML5DI1", "10.1.100.152:1522/L5DH")
cur = conn.cursor() # 獲取操作遊標，也就是開始操作
cur.arraysize = 50 #每次取50行做儲存 (優化)
# execute a query returning the results to the cursor
print("Starting cursor at", time.ctime())


# In[37]:

# 0.1/144=一分鐘，0.017=10秒鐘
#查詢資料庫
sql ="""select O.TEST_TIME,O.op_id,O.ko_id,O.abbr_no,O.lot_id, O.sheet_id,O.slot_no, O.DFT_CODE,O.pox_x, O.pox_y, O.img_file_path, E.OPER_ID,E.EQP_ID,E.UNIT_ID,E.LOGOFF_TIME,O.lm_time,O.AOI_REPAIR_FLAG from	(select A.*,B.slot_no from
		(select test_time,tft_lot_id lot_id, tft_sheet_id sheet_id,CURRENT_DEF_CODE_DESC DFT_CODE, substr(OP_KEY,1,4) op_id, eqp_id ko_id,
		substr(TFT_LOT_ID,1,2) abbr_no, pox_x, pox_y,'http://hyadiwa1/dms/L5DAIDI/ARYAOI_L5D/' | | IMG_FILE_URL_PATH || IMAGE_FILE_NAME IMG_FILE_PATH,to_char(lm_time,'yyyy/mm/dd hh24:mi:ss') lm_time, RETYPE_REPAIR_FLAG AOI_REPAIR_FLAG
			from ARYH.H_AIDI_SECDEFECT 
			where CURRENT_DEF_CODE_DESC not in 'null' 
			and lm_time >= sysdate-0.1/144) A 
	join aryh.H_AIDI_SECSHEET B 
	on A.sheet_id = B.tft_sheet_id
        and A.op_id =substr(B.OP_KEY,1,4)
	and A.test_time = B.test_time) O
 join \
 (\
	SELECT*FROM 
       		 (SELECT B.LOT_ID,B.SHEET_ID_CHIP_ID SHEET_ID,B.OP_ID OPER_ID,B.EQP_ID,B.ROUTE_ID,A.UNIT_ID,
               	 to_char(B.LOGOFF_TIME,'YYYY/MM/DD HH24:MI:SS') AS LOGOFF_TIME,
               	 RANK() OVER (PARTITION BY B.SHEET_ID_CHIP_ID,B.OP_ID ORDER BY B.LOGOFF_TIME DESC) RANK 
                   	 FROM ARYODS.H_SHEET_OPER_ODS B Left Outer Join ARYODS.H_DAX_EQPUNIT_ODS A ON  
                        B.SHEET_ID_CHIP_ID=A.SHEET_ID 
                        AND B.EQP_ID=A.EQP_ID \
                        AND SUBSTR(B.OP_ID,1,4)=SUBSTR(A.OP_ID,1,4) \
                        AND to_char(B.LOGOFF_TIME,'yyyy/mm/dd')<=to_char(A.PROCESS_TIME,'yyyy/mm/dd')  \
               ) C WHERE RANK=1 \
) E\
 on O.sheet_id = E.SHEET_ID\
 and to_char(O.test_time,'yyyy/mm/dd hh24:mi:ss') > E.LOGOFF_TIME\
                   """
cur.execute(sql) # 執行查詢語句 抓一個小時資料更新
print("Finished cursor at", time.ctime())
# where O.DFT_CODE='E-PV-DP Hole'\
# and O.lot_id='PW99HABC90'\


# In[51]:


spec = pd.read_excel('\\\L5di1\\d\\HYBI1\\晨會報告\\Array\\Array Daily Report\\Server run\\停機_Defect_spec_V3.xlsx',dtype={'op_id':np.str},na_values='NULL')
spec = spec.fillna('')
spec


# In[52]:


spec['oper_id(重要站點)'] = spec['oper_id(重要站點)'].str.split(',')
spec['主旨帶入站點'] = spec['主旨帶入站點'].str.split(',')
spec['EVENTID'] = spec['EVENTID'].str.split(',')
spec


# # raw data

# In[40]:


rows = cur.fetchall() # 獲取查詢結果
col_result = cur.description #獲取查詢結果的欄位描述
columns = []
for i in range(len(col_result)):
    columns.append(col_result[i][0])  # 獲取欄位名，以列表形式儲存
df = pd.DataFrame(rows,columns=columns)   #轉dataframe格式
df = df.sort_values(['OPER_ID','SHEET_ID']) #按照SHEET_ID、OPER_ID數字升冪排序
pd.set_option('display.max_colwidth', -1) # 將IMG_FILE_PATH完整秀出
df.loc[:,"IMG_FILE_PATH"]='<a href="'+df[["IMG_FILE_PATH"]]+'" >'+'<img width="200" height="200" src="'+df[["IMG_FILE_PATH"]]+'" />'+'</a>' #將IMG_FILE_PATH轉為HTML格式
df.head()


# # 計算每一批的Defect code的數量

# In[41]:


DFT_Cnt = df.groupby(['LOT_ID','DFT_CODE'])['IMG_FILE_PATH'].nunique() # 使用nunique 去計算IMG_FILE_PATH避免重複計算
DFT_Cnt = pd.DataFrame(DFT_Cnt)
DFT_Cnt = DFT_Cnt.rename(columns={"IMG_FILE_PATH": "DFT_Cnt"})
print(DFT_Cnt)


# # 計算每一批by ABBR_NO、OP_ID的Defect code的數量

# In[42]:


DEFECT_COUNT_BY_ABBR_OP = df.groupby(['LOT_ID','ABBR_NO','DFT_CODE','OP_ID'])['IMG_FILE_PATH'].nunique() # 使用nunique 去計算IMG_FILE_PATH避免重複計算
DEFECT_COUNT_BY_ABBR_OP = pd.DataFrame(DEFECT_COUNT_BY_ABBR_OP)
DEFECT_COUNT_BY_ABBR_OP = DEFECT_COUNT_BY_ABBR_OP.rename(columns={"IMG_FILE_PATH": "DFT_Cnt"})
print(DEFECT_COUNT_BY_ABBR_OP)


# # 計算每一片Defect code的數量

# In[43]:


DEFECT_COUNT_BY_SHEET_ID = df.groupby(['LOT_ID','ABBR_NO','DFT_CODE','OP_ID','SHEET_ID'])['IMG_FILE_PATH'].nunique() # 使用nunique 去計算IMG_FILE_PATH避免重複計算
DEFECT_COUNT_BY_SHEET_ID = pd.DataFrame(DEFECT_COUNT_BY_SHEET_ID)
DEFECT_COUNT_BY_SHEET_ID = DEFECT_COUNT_BY_SHEET_ID.rename(columns={"IMG_FILE_PATH": "DFT_Cnt"})
DEFECT_COUNT_BY_SHEET_ID = DEFECT_COUNT_BY_SHEET_ID.reset_index()
print(type(DEFECT_COUNT_BY_SHEET_ID))
print(DEFECT_COUNT_BY_SHEET_ID)


# # 計算每片顆數 join 原始資料

# In[44]:


df = pd.merge(df,DEFECT_COUNT_BY_SHEET_ID,on=('LOT_ID','ABBR_NO','DFT_CODE','OP_ID','SHEET_ID'),how='left')
df


# # 一批LOT 檢出幾片

# In[45]:


SHEET = df.groupby(['ABBR_NO','LOT_ID','OP_ID','SHEET_ID']).count()
LOT_COUNT = SHEET.groupby(['LOT_ID']).count()
LOT_COUNT = LOT_COUNT.rename(columns={"DFT_CODE": "LOT_COUNT"})
LOT_COUNT = LOT_COUNT[['LOT_COUNT']]
print(LOT_COUNT) 


# # 每一批的Density by Defect code

# In[46]:


Density = DFT_Cnt['DFT_Cnt'] / LOT_COUNT['LOT_COUNT']
Density = round(Density,2)
Density


# # 每一批的Density by Defect code、ABBR_NO、OP_ID

# In[47]:


Density_BY_ABBR_OP = DEFECT_COUNT_BY_ABBR_OP['DFT_Cnt'] / LOT_COUNT['LOT_COUNT']
Density_BY_ABBR_OP = round(Density_BY_ABBR_OP,2)
Density_BY_ABBR_OP

# # 轉XML檔 & 登入FTP上拋函式

# In[38]:


import xml.etree.cElementTree as ET
from ftplib import FTP         #導入ftp套件
ftp = FTP("10.22.10.43")           #設定ftp伺服器地址
ftp.login('anonymous', 'anonymous')      #設定登入賬號和密碼
def xml_ftp_upload(tool_id,EVENT_ID,TRIGGER_TIME,MESSAGE,SITE_ID,ALARM_CODE,MAIL_SUBJECT,MAIL_BODY,file_name):
    root = ET.Element("MSG")

    toolid = ET.SubElement(root, "TOOLID")
    toolid.text = tool_id

    EVENTID = ET.SubElement(root, "EVENTID")
    EVENTID.text = EVENT_ID

    TRIGGERTIME = ET.SubElement(root, "TRIGGERTIME")
    TRIGGERTIME.text = TRIGGER_TIME

    ALARM_MESSAGE = ET.SubElement(root, "ALARM_MESSAGE")
    ALARM_MESSAGE.text = MESSAGE

    SITEID = ET.SubElement(root, "SITEID")
    SITEID.text = SITE_ID

    ALARMCODE = ET.SubElement(root, "ALARMCODE")
    ALARMCODE.text = ALARM_CODE

    MAILSUBJECT = ET.SubElement(root, "MAILSUBJECT")
    MAILSUBJECT.text = MAIL_SUBJECT

    MAILBODY = ET.SubElement(root, "MAILBODY")
    MAILBODY.text = MAIL_BODY
    
    URL = ET.SubElement(root, "URL")

    tree = ET.ElementTree(root)
    tree.write("C:\\Users\\SHANESLU\\Desktop\\python\\Tornado\\"+file_name+'.xml',encoding="UTF-8") #建立XML檔，需注意檔名、路徑
    
    file_remote = file_name+'.xml'
    file_local = "C:\\Users\\SHANESLU\\Desktop\\python\\Tornado\\"+file_name+'.xml' #讀xml檔，需注意檔名、路徑
    fp = open(file_local, 'rb')
    ftp.storbinary('STOR ' + file_remote, fp, 1024) #上拋FTP，暫存空間為1024
    fp.close()

# ##  評斷是否過spec和發送mail的function

# In[48]:


def judge(i,j):
    if spec['mail+hold'][i] > Density_BY_ABBR_OP[j] >= spec['mail'][i]:
        a = df[( df['LOT_ID']==Density_BY_ABBR_OP.index[j][0]) & (df['DFT_CODE']==Density_BY_ABBR_OP.index[j][2])]['SLOT_NO']
        b = []
        for k in range(len(a)):
            if a.iloc[k] not in b:   #抓出哪幾片
                b.append(a.iloc[k])
                continue
        c=b #內文片號
        if len(b)>6:
            b = '['+str(len(b))+' Sheets]'
        print(Density_BY_ABBR_OP.index[j][3]+
                '\nslot '+str(b)+' suffered '+Density_BY_ABBR_OP.index[j][2]+' issue.', 
                '\nLOT_ID: '+Density_BY_ABBR_OP.index[j][0], #LOT_ID
                '\nDEFECT_CODE: '+Density_BY_ABBR_OP.index[j][2], #DEFECT_CODE
                '\n要求確認機台狀況(mail): '+str(spec['mail'][i]),
                '\n單批異常立刻停機(mail+hold): '+str(spec['mail+hold'][i]),
                '\nDefect Qty: '+str(DFT_Cnt['DFT_Cnt'][Density_BY_ABBR_OP.index[j][0]][Density_BY_ABBR_OP.index[j][2]]), #DFT_Cnt
                '\nSheet Qty: '+str(LOT_COUNT['LOT_COUNT'][Density_BY_ABBR_OP.index[j][0]]), #LOT_COUNT
                '\nDensity: '+str(Density_BY_ABBR_OP[j])) #DENSITY
                # 抓出對應的RAW DATA     
        raw_data = df[( df['LOT_ID']==Density_BY_ABBR_OP.index[j][0]) & (df['DFT_CODE']==Density_BY_ABBR_OP.index[j][2])]
        #EQP_ID + UNIT_ID
        raw_data['UNIT_ID'] = raw_data['EQP_ID']+" "+raw_data['UNIT_ID'].fillna('') #需加入fillna('')否則無法合併
        #製程站點取前四碼+後四碼
        raw_data['OPER_ID'] = raw_data['OPER_ID'].str[:4]+raw_data['OPER_ID'].str[7:11]
        #INFORMATION
        key_oper = raw_data
        key_oper = key_oper.drop(['ABBR_NO','OP_ID','KO_ID','POX_X','POX_Y','IMG_FILE_PATH','TEST_TIME'],axis=1) 
        key_oper = key_oper.drop_duplicates()
        key_oper = key_oper.reset_index(drop=True)
        key_oper = key_oper[['LOT_ID','SHEET_ID','SLOT_NO','DFT_CODE','DFT_Cnt','OPER_ID','EQP_ID','UNIT_ID','LOGOFF_TIME']]
        #RAW_DATA
        raw_data = raw_data.drop(['ABBR_NO','OPER_ID','EQP_ID','UNIT_ID','LOGOFF_TIME'],axis=1)
        raw_data = raw_data.drop_duplicates()
        raw_data = raw_data.reset_index(drop=True)
        raw_data = raw_data[['TEST_TIME','OP_ID','KO_ID','LOT_ID','SHEET_ID','DFT_CODE','DFT_Cnt','POX_X','POX_Y','IMG_FILE_PATH','AOI_REPAIR_FLAG']]
           #濾出製程站點
        if spec['oper_id(重要站點)'][i] != '':   
            if type(spec['oper_id(重要站點)'][i]) == str:  #單一站點
                # oper title
                key_oper_title = key_oper[key_oper['OPER_ID']==spec['主旨帶入站點'][i]]
                key_oper_title = key_oper_title.reset_index(drop=True)
                ALL_TOOL = key_oper_title['UNIT_ID'].unique().tolist()
                ALL_TOOL = ",".join(ALL_TOOL)
                # oper 內文
                key_oper = key_oper[key_oper['OPER_ID']==spec['oper_id(重要站點)'][i]]
                key_oper = key_oper.reset_index(drop=True)
                    #修改站點名稱
                key_oper = key_oper.rename(columns={"UNIT_ID":key_oper['OPER_ID'].iloc[0]})
                key_oper = key_oper.drop(['EQP_ID','OPER_ID'],axis=1)
            else:   #多個站點
                # oper title
                if type(spec['主旨帶入站點'][i]) == str:
                    key_oper_title = key_oper[key_oper['OPER_ID']==spec['主旨帶入站點'][i]]
                    key_oper_title = key_oper_title.reset_index(drop=True)
                    ALL_TOOL = key_oper_title['UNIT_ID'].unique().tolist()
                    ALL_TOOL = ",".join(ALL_TOOL)
                else:
                    key_oper_title = key_oper[key_oper['OPER_ID'].isin(spec['主旨帶入站點'][i])]
                    key_oper_title = key_oper_title.reset_index(drop=True)
                    ALL_TOOL = key_oper_title['UNIT_ID'].unique().tolist()
                    ALL_TOOL = ",".join(ALL_TOOL)
                #oper 內文
                key_oper_filter = key_oper[key_oper['OPER_ID'].isin(spec['oper_id(重要站點)'][i])]
                oper_list = key_oper_filter['OPER_ID'].unique().tolist()  #帶出有的站點
                key_oper = key_oper_filter[key_oper_filter['OPER_ID']==oper_list[0]].reset_index(drop=True)       
                # 刪除"第一站"合併機台多餘欄位
                    #修改站點名稱
                key_oper = key_oper.rename(columns={"UNIT_ID":key_oper['OPER_ID'].iloc[0]})
                key_oper = key_oper.drop(['EQP_ID','OPER_ID'],axis=1)
                    #切割製程站點、機台、時間
                for a in range(len(oper_list)-1):
                    key_oper1 = key_oper_filter[key_oper_filter['OPER_ID']==oper_list[a+1]]['OPER_ID']
                    UNIT_ID = key_oper_filter[key_oper_filter['OPER_ID']==oper_list[a+1]]['UNIT_ID']
                    LOGOFF_TIME = key_oper_filter[key_oper_filter['OPER_ID']==oper_list[a+1]]['LOGOFF_TIME']
                    key_oper1 = pd.Series(key_oper1).reset_index(drop=True)
                    UNIT_ID = pd.Series(UNIT_ID).reset_index(drop=True)
                    LOGOFF_TIME = pd.Series(LOGOFF_TIME).reset_index(drop=True)
                    key_oper[key_oper1.iloc[0]]=UNIT_ID
                    key_oper['LOGOFF_TIME'+str(a+1)]=LOGOFF_TIME
                        #key_oper.insert(a+7,'LOGOFF_TIME'+str(a+1), LOGOFF_TIME)
                    #key_oper = pd.concat([key_oper,key_oper1],axis=1)
                    #key_oper['OPER_ID'+str(a+1)]=key_oper1
                    #key_oper['OPER_ID'+str(a+1)]=key_oper1[key_oper1['OPER_ID']==spec['oper_id(重要站點)'][i][a+1]].loc[:,'OPER_ID']
                    #key_oper['EQP_ID'+str(a+1)]=key_oper1[key_oper1['EQP_ID']==spec['oper_id(重要站點)'][i][a+1]].loc[:,'EQP_ID']
        elif spec['oper_id(重要站點)'][i] =='': #全部站點
            # 刪除"第一站"合併機台多餘欄位
            all_oper = key_oper['OPER_ID'].unique().tolist()
            key_oper_filter = key_oper[key_oper['OPER_ID'].isin(all_oper)]
            key_oper = key_oper[key_oper['OPER_ID']==all_oper[0]]
                #濾出所有機台
            #ALL_TOOL = key_oper_filter['UNIT_ID'].unique().tolist()
            #ALL_TOOL = ",".join(ALL_TOOL)
            ALL_TOOL=''
            key_oper = key_oper.rename(columns={"UNIT_ID":key_oper['OPER_ID'].iloc[0]})
            key_oper = key_oper.drop(['EQP_ID','OPER_ID'],axis=1)  
            #切割製程站點、機台、時間
            for a in range(len(all_oper)-1):
                key_oper1 = key_oper_filter[key_oper_filter['OPER_ID']==all_oper[a+1]]['OPER_ID']
                UNIT_ID = key_oper_filter[key_oper_filter['OPER_ID']==all_oper[a+1]]['UNIT_ID']
                LOGOFF_TIME = key_oper_filter[key_oper_filter['OPER_ID']==all_oper[a+1]]['LOGOFF_TIME']
                key_oper1 = pd.Series(key_oper1).reset_index(drop=True)
                UNIT_ID = pd.Series(UNIT_ID).reset_index(drop=True)
                LOGOFF_TIME = pd.Series(LOGOFF_TIME).reset_index(drop=True)
                key_oper[key_oper1.iloc[0]]=UNIT_ID
                #key_oper.insert(a+6,key_oper1.iloc[0], EQP_ID1)
                key_oper['LOGOFF_TIME'+str(a+1)]=LOGOFF_TIME
                #key_oper.insert(a+7,'LOGOFF_TIME'+str(a+1), LOGOFF_TIME)
        INFORMATION = key_oper.drop_duplicates()
            # defect code 照片
        picture = raw_data['IMG_FILE_PATH']
                # 散佈圖
        #plt.style.use('ggplot')  #使用ggplot美化
        plt.figure(figsize=(4,3.5))
        plt.scatter(df[( df['LOT_ID']==Density_BY_ABBR_OP.index[j][0]) & (df['DFT_CODE']==Density_BY_ABBR_OP.index[j][2])].POX_X/1000,
                        df[( df['LOT_ID']==Density_BY_ABBR_OP.index[j][0]) & (df['DFT_CODE']==Density_BY_ABBR_OP.index[j][2])].POX_Y/1000)
        plt.xlim(0, 1300) #固定X axis為1300
        plt.ylim(0, 1100) #固定y axis為1100
                         # figure 儲存為二進位檔案
        plt.grid(b=True, which='major', color='#666666', linestyle='--')                 
        buffer = BytesIO()
        plt.savefig(buffer)  
        plot_data = buffer.getvalue()
        # 圖片資料轉為 HTML 格式
        imb = base64.b64encode(plot_data)  
        ims = imb.decode()
        imd = "data:image/png;base64,"+ims
        scatter = "<h4>散佈圖"+"_"+Density_BY_ABBR_OP.index[j][2]+" & 九宮格照片</h4>"# + """<img src=\'"%s\'> """ % imd   
         # 圖片存檔 轉為Base64
        plt.savefig('C:\\Users\\SHANESLU\\Desktop\\python\\AOI_scatter\\'+'散佈圖'+'_'+Density_BY_ABBR_OP.index[j][2]+'.png')
        plt.pause(0.01) #僅顯示0.01秒
        with open('C:\\Users\\SHANESLU\\Desktop\\python\\AOI_scatter\\'+'散佈圖'+'_'+Density_BY_ABBR_OP.index[j][2]+'.png', "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read())
            encoded_string = encoded_string.decode('utf-8')  #去掉 b'XXXX' 
#ManualSend_39  
        if len(picture)>=9:
                    picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" rowspan="3"><img src=\'"%s\'></td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd,picture.iloc[3],picture.iloc[4],
                                               picture.iloc[5],picture.iloc[6],picture.iloc[7],picture.iloc[8])
                    )
        elif len(picture)==8:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" rowspan="3"><img src=\'"%s\'></td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd,picture.iloc[3],picture.iloc[4],
                                               picture.iloc[5],picture.iloc[6],picture.iloc[7])
                     )
        elif len(picture)==7:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" rowspan="3"><img src=\'"%s\'></td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd,picture.iloc[3],picture.iloc[4],
                                               picture.iloc[5],picture.iloc[6])
                     )
        elif len(picture)==6:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" rowspan="2"><img src=\'"%s\'></td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd,picture.iloc[3],picture.iloc[4],
                                               picture.iloc[5])
                     )
        elif len(picture)==5:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" rowspan="2"><img src=\'"%s\'></td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd,picture.iloc[3],picture.iloc[4]
                                               )
                     )
        elif len(picture)==4:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" rowspan="2"><img src=\'"%s\'></td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd,picture.iloc[3]
                                               )
                     )
        elif len(picture)==3:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" colspan="3"><img src=\'"%s\'></td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd)
                     )
        elif len(picture)==2:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" colspan="2"><img src=\'"%s\'></td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],imd)
                     )
        elif len(picture)==1:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td width="250" colspan="2"><img src=\'"%s\'></td>
                                 </tr>
                                </table>""" % (picture.iloc[0],imd)
                     )
        if Density_BY_ABBR_OP.index[j][2][:1] == 'T':
            print('T')
            ManualSend_39 = {
                        'strMailCode': 'sDM8fpGZVQ8=', #MailCode
                        'strRecipients':spec['function'][i], #收件人
                        'strCopyRecipients':'andyou@auo.com', #副本
                        'strSubject':'AOI'+'_'+Density_BY_ABBR_OP.index[j][3]+'_'+Density_BY_ABBR_OP.index[j][0]+'_'+str(b)+' Suffered '+Density_BY_ABBR_OP.index[j][2]+' : '+str(ALL_TOOL), #標題
                        'strBody':'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][0]+
                        ' </font></b>'+'slot '+'<b><font color="#E00000">'+str(c)+'</font></b>'+
                        ' suffered '+'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][2]+
                        '</font></b>'+' issue.'+'<br>'+'<br>'+ 
                        """<table border="1">
                             <tr>
                             <td bgcolor="#FFB630">要求確認機台狀況(mail)</td>
                             <td bgcolor="#FFB630"><b>%s</b></td>
                             </tr>
                             <tr>
                             <td bgcolor="#DB0000"><font color="#FFFFFF">單批異常立刻停機(mail+hold)</td>
                             <td bgcolor="#DB0000"><b><font color="#FFFFFF">%s</font></b></td>
                             </tr>
                             <tr>
                             <td>Sheet Qty</td>
                             <td>%s</td>
                             </tr>
                             <tr>
                             <td>Defect Qty</td>
                             <td>%s</td>
                             </tr>
                              <tr>
                             <td bgcolor="#00DB00"><font color="#0000D6">Density</td>
                             <td bgcolor="#00DB00"><b><font color="#0000D6">%s</b></td>
                             </tr>
                            </table>""" % (str(spec['mail'][i]),str(spec['mail+hold'][i]),
                                           str(LOT_COUNT['LOT_COUNT'][Density_BY_ABBR_OP.index[j][0]]),
                                           str(DFT_Cnt['DFT_Cnt'][Density_BY_ABBR_OP.index[j][0]][Density_BY_ABBR_OP.index[j][2]]),
                                           str(Density_BY_ABBR_OP[j]))+
                         "<h2>Information</h2>"+str(INFORMATION.to_html())+        
                        scatter+
                            str(picture9)+
                        "<h2>EDA Link</h2>"+
                '<a href="http://autceda/dashboard/sites/L5D/show/L5D_DS_TEST_Array_T1_Defect_Density_Light/PVD_KO_Density_Light?title=Density%20LIGHT&layers=160&layers=321&layers=368&">ML5DT1</a>'+
                "<br>"+'<a href="http://autceda/dashboard/sites/L5D/show/L5D_Array_INT_7D_InlineSummary/7D_InlineSummary?Defect Code=%s">InlineSummary</a>'%(str(Density_BY_ABBR_OP.index[j][2]))+
                        "<h2>Raw data(抓前50筆)</h2>"+str(raw_data[:50].to_html()), #raw data 抓前50筆 '
                        'strFileBase64String':'散佈圖'+'_'+Density_BY_ABBR_OP.index[j][2]+'.png:'+str(encoded_string) #夾檔案 需先將圖片轉base64， 格式要為: 'XXX.png' +   ':' + encoded_string
                        }
            response_39 = client.service.ManualSend_39(**ManualSend_39)
            print('是否成功發送mail: '+str(response_39))
            print()
        elif Density_BY_ABBR_OP.index[j][2][:1] == 'P':
            print('P')
            ManualSend_39 = {
                        'strMailCode': 'sDM8fpGZVQ8=', #MailCode
                        'strRecipients':spec['function'][i], #收件人
                        'strCopyRecipients':'andyou@auo.com', #副本
                        'strSubject':'AOI'+'_'+Density_BY_ABBR_OP.index[j][3]+'_'+Density_BY_ABBR_OP.index[j][0]+'_'+str(b)+' Suffered '+Density_BY_ABBR_OP.index[j][2]+' : '+str(ALL_TOOL), #標題
                        'strBody':'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][0]+
                        ' </font></b>'+'slot '+'<b><font color="#E00000">'+str(c)+'</font></b>'+
                        ' suffered '+'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][2]+
                        '</font></b>'+' issue.'+'<br>'+'<br>'+ 
                        """<table border="1">
                             <tr>
                             <td bgcolor="#FFB630">要求確認機台狀況(mail)</td>
                             <td bgcolor="#FFB630"><b>%s</b></td>
                             </tr>
                             <tr>
                             <td bgcolor="#DB0000"><font color="#FFFFFF">單批異常立刻停機(mail+hold)</td>
                             <td bgcolor="#DB0000"><b><font color="#FFFFFF">%s</font></b></td>
                             </tr>
                             <tr>
                             <td>Sheet Qty</td>
                             <td>%s</td>
                             </tr>
                             <tr>
                             <td>Defect Qty</td>
                             <td>%s</td>
                             </tr>
                              <tr>
                             <td bgcolor="#00DB00"><font color="#0000D6">Density</td>
                             <td bgcolor="#00DB00"><b><font color="#0000D6">%s</b></td>
                             </tr>
                            </table>""" % (str(spec['mail'][i]),str(spec['mail+hold'][i]),
                                           str(LOT_COUNT['LOT_COUNT'][Density_BY_ABBR_OP.index[j][0]]),
                                           str(DFT_Cnt['DFT_Cnt'][Density_BY_ABBR_OP.index[j][0]][Density_BY_ABBR_OP.index[j][2]]),
                                           str(Density_BY_ABBR_OP[j]))+
                         "<h2>Information</h2>"+str(INFORMATION.to_html())+        
                        scatter+
                            str(picture9)+
                        "<h2>EDA Link</h2>"+
                        '<a href="http://autceda/dashboard/sites/L5D/show/L5D_Array_INT_7D_InlineSummary/7D_InlineSummary?Defect Code=%s">InlineSummary</a>'%(str(Density_BY_ABBR_OP.index[j][2]))+
                        "<h2>Raw data(抓前50筆)</h2>"+str(raw_data[:50].to_html()), #raw data 抓前50筆 '
                        'strFileBase64String':'散佈圖'+'_'+Density_BY_ABBR_OP.index[j][2]+'.png:'+str(encoded_string) #夾檔案 需先將圖片轉base64， 格式要為: 'XXX.png' +   ':' + encoded_string
                        }
            response_39 = client.service.ManualSend_39(**ManualSend_39)
            print('是否成功發送mail: '+str(response_39))
            print()
        elif Density_BY_ABBR_OP.index[j][2][:1] == 'E':
            print('E')
            ManualSend_39 = {
                        'strMailCode': 'sDM8fpGZVQ8=', #MailCode
                        'strRecipients':spec['function'][i], #收件人
                        'strCopyRecipients':'andyou@auo.com', #副本
                        'strSubject':'AOI'+'_'+Density_BY_ABBR_OP.index[j][3]+'_'+Density_BY_ABBR_OP.index[j][0]+'_'+str(b)+' Suffered '+Density_BY_ABBR_OP.index[j][2]+' : '+str(ALL_TOOL), #標題
                        'strBody':'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][0]+
                        ' </font></b>'+'slot '+'<b><font color="#E00000">'+str(c)+'</font></b>'+
                        ' suffered '+'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][2]+
                        '</font></b>'+' issue.'+'<br>'+'<br>'+ 
                        """<table border="1">
                             <tr>
                             <td bgcolor="#FFB630">要求確認機台狀況(mail)</td>
                             <td bgcolor="#FFB630"><b>%s</b></td>
                             </tr>
                             <tr>
                             <td bgcolor="#DB0000"><font color="#FFFFFF">單批異常立刻停機(mail+hold)</td>
                             <td bgcolor="#DB0000"><b><font color="#FFFFFF">%s</font></b></td>
                             </tr>
                             <tr>
                             <td>Sheet Qty</td>
                             <td>%s</td>
                             </tr>
                             <tr>
                             <td>Defect Qty</td>
                             <td>%s</td>
                             </tr>
                              <tr>
                             <td bgcolor="#00DB00"><font color="#0000D6">Density</td>
                             <td bgcolor="#00DB00"><b><font color="#0000D6">%s</b></td>
                             </tr>
                            </table>""" % (str(spec['mail'][i]),str(spec['mail+hold'][i]),
                                           str(LOT_COUNT['LOT_COUNT'][Density_BY_ABBR_OP.index[j][0]]),
                                           str(DFT_Cnt['DFT_Cnt'][Density_BY_ABBR_OP.index[j][0]][Density_BY_ABBR_OP.index[j][2]]),
                                           str(Density_BY_ABBR_OP[j]))+
                 "<h2>Information</h2>"+str(INFORMATION.to_html())+        
                scatter+
                            str(picture9)+
                        "<h2>EDA Link</h2>"+
                '<a href="http://autceda/dashboard/sites/L5D/show/L5D_DS_TEST_Array_E1_Defect_Density_Light/DRY_KO_Density_Light?title=Density_Light&layers=160&layers=321&layers=486&">ML5DE1</a>'+
                "<br>"+'<a href="http://autceda/dashboard/sites/L5D/show/L5D_DS_TEST_Array_E2_Defect_Density_Light/WET_KO_Density_Light?title=Defect_Light&layers=160&layers=321&layers=540&">ML5DE2</a>'+
                "<br>"+'<a href="http://autceda/dashboard/sites/L5D/show/L5D_Array_INT_7D_InlineSummary/7D_InlineSummary?Defect Code=%s">InlineSummary</a>'%(str(Density_BY_ABBR_OP.index[j][2]))+
                        "<h2>Raw data(抓前50筆)</h2>"+str(raw_data[:50].to_html()), #raw data 抓前50筆 '
                        'strFileBase64String':'散佈圖'+'_'+Density_BY_ABBR_OP.index[j][2]+'.png:'+str(encoded_string) #夾檔案 需先將圖片轉base64， 格式要為: 'XXX.png' +   ':' + encoded_string
                        }
            response_39 = client.service.ManualSend_39(**ManualSend_39)
            print('是否成功發送mail: '+str(response_39))
            print()
        elif Density_BY_ABBR_OP.index[j][2][:1] == 'I':
            print('I')
            ManualSend_39 = {
                        'strMailCode': 'sDM8fpGZVQ8=', #MailCode
                        'strRecipients':spec['function'][i], #收件人
                        'strCopyRecipients':'andyou@auo.com', #副本
                        'strSubject':'AOI'+'_'+Density_BY_ABBR_OP.index[j][3]+'_'+Density_BY_ABBR_OP.index[j][0]+'_'+str(b)+' Suffered '+Density_BY_ABBR_OP.index[j][2]+' : '+str(ALL_TOOL), #標題
                        'strBody':'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][0]+
                        ' </font></b>'+'slot '+'<b><font color="#E00000">'+str(c)+'</font></b>'+
                        ' suffered '+'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][2]+
                        '</font></b>'+' issue.'+'<br>'+'<br>'+ 
                        """<table border="1">
                             <tr>
                             <td bgcolor="#FFB630">要求確認機台狀況(mail)</td>
                             <td bgcolor="#FFB630"><b>%s</b></td>
                             </tr>
                             <tr>
                             <td bgcolor="#DB0000"><font color="#FFFFFF">單批異常立刻停機(mail+hold)</td>
                             <td bgcolor="#DB0000"><b><font color="#FFFFFF">%s</font></b></td>
                             </tr>
                             <tr>
                             <td>Sheet Qty</td>
                             <td>%s</td>
                             </tr>
                             <tr>
                             <td>Defect Qty</td>
                             <td>%s</td>
                             </tr>
                              <tr>
                             <td bgcolor="#00DB00"><font color="#0000D6">Density</td>
                             <td bgcolor="#00DB00"><b><font color="#0000D6">%s</b></td>
                             </tr>
                            </table>""" % (str(spec['mail'][i]),str(spec['mail+hold'][i]),
                                           str(LOT_COUNT['LOT_COUNT'][Density_BY_ABBR_OP.index[j][0]]),
                                           str(DFT_Cnt['DFT_Cnt'][Density_BY_ABBR_OP.index[j][0]][Density_BY_ABBR_OP.index[j][2]]),
                                           str(Density_BY_ABBR_OP[j]))+
                         "<h2>Information</h2>"+str(INFORMATION.to_html())+        
                        scatter+
                            str(picture9)+
                        "<h2>EDA Link</h2>"+
                '<a href="http://autceda/dashboard/sites/L5D/show/L5D_DS_TEST_Array_T1_Defect_Density_Light/PVD_KO_Density_Light?title=Density%20LIGHT&layers=160&layers=321&layers=368&">ML5DT1</a>'+
                "<br>"+
                '<a href="http://autceda/dashboard/sites/L5D/show/L5D_DS_TEST_Array_E1_Defect_Density_Light/DRY_KO_Density_Light?title=Density_Light&layers=160&layers=321&layers=486&">ML5DE1</a>'+
                "<br>"+'<a href="http://autceda/dashboard/sites/L5D/show/L5D_DS_TEST_Array_E2_Defect_Density_Light/WET_KO_Density_Light?title=Defect_Light&layers=160&layers=321&layers=540&">ML5DE2</a>'+
                "<br>"+'<a href="http://autceda/dashboard/sites/L5D/show/L5D_Array_INT_7D_InlineSummary/7D_InlineSummary?Defect Code=%s">InlineSummary</a>'%(str(Density_BY_ABBR_OP.index[j][2]))+
                        "<h2>Raw data(抓前50筆)</h2>"+str(raw_data[:50].to_html()), #raw data 抓前50筆 '
                        'strFileBase64String':'散佈圖'+'_'+Density_BY_ABBR_OP.index[j][2]+'.png:'+str(encoded_string) #夾檔案 需先將圖片轉base64， 格式要為: 'XXX.png' +   ':' + encoded_string
                        }
            response_39 = client.service.ManualSend_39(**ManualSend_39)
            print('是否成功發送mail: '+str(response_39))
            print()
    elif Density_BY_ABBR_OP[j] >= spec['mail+hold'][i]:
        a = df[( df['LOT_ID']==Density_BY_ABBR_OP.index[j][0]) & (df['DFT_CODE']==Density_BY_ABBR_OP.index[j][2])]['SLOT_NO']
        b = []
        for k in range(len(a)):
            if a.iloc[k] not in b:   #抓出哪幾片
                b.append(a.iloc[k])
                continue
        c=b #內文片號        
        if len(b)>6:
            b = '['+str(len(b))+' Sheets]'
        print(Density_BY_ABBR_OP.index[j][3]+
                '\nslot '+str(b)+' suffered '+Density_BY_ABBR_OP.index[j][2]+' issue.', 
                '\nLOT_ID: '+Density_BY_ABBR_OP.index[j][0], #LOT_ID
                '\nDEFECT_CODE: '+Density_BY_ABBR_OP.index[j][2], #DEFECT_CODE
                '\n要求確認機台狀況(mail): '+str(spec['mail'][i]),
                '\n單批異常立刻停機(mail+hold): '+str(spec['mail+hold'][i]),
                '\nDefect Qty: '+str(DFT_Cnt['DFT_Cnt'][Density_BY_ABBR_OP.index[j][0]][Density_BY_ABBR_OP.index[j][2]]), #DFT_Cnt
                '\nSheet Qty: '+str(LOT_COUNT['LOT_COUNT'][Density_BY_ABBR_OP.index[j][0]]), #LOT_COUNT
                '\nDensity: '+str(Density_BY_ABBR_OP[j])) #DENSITY
                # 抓出對應的RAW DATA      
        raw_data = df[( df['LOT_ID']==Density_BY_ABBR_OP.index[j][0]) & (df['DFT_CODE']==Density_BY_ABBR_OP.index[j][2])]
        #EQP_ID + UNIT_ID 
        raw_data['UNIT_ID'] = raw_data['EQP_ID']+" "+raw_data['UNIT_ID'].fillna('') #需加入fillna('')否則無法合併
        #製程站點取前四碼+後四碼
        raw_data['OPER_ID'] = raw_data['OPER_ID'].str[:4]+raw_data['OPER_ID'].str[7:11]
        #INFORMATION
        key_oper = raw_data
        key_oper = key_oper.drop(['ABBR_NO','OP_ID','KO_ID','POX_X','POX_Y','IMG_FILE_PATH','TEST_TIME'],axis=1) 
        key_oper = key_oper.drop_duplicates()
        key_oper = key_oper.reset_index(drop=True)
        key_oper = key_oper[['LOT_ID','SHEET_ID','SLOT_NO','DFT_CODE','DFT_Cnt','OPER_ID','EQP_ID','UNIT_ID','LOGOFF_TIME']]
        #RAW_DATA
        raw_data = raw_data.drop(['ABBR_NO','OPER_ID','EQP_ID','UNIT_ID','LOGOFF_TIME'],axis=1)
        raw_data = raw_data.drop_duplicates()
        raw_data = raw_data.reset_index(drop=True)
        raw_data = raw_data[['TEST_TIME','OP_ID','KO_ID','LOT_ID','SHEET_ID','DFT_CODE','DFT_Cnt','POX_X','POX_Y','IMG_FILE_PATH','AOI_REPAIR_FLAG']]
           #濾出製程站點
        if spec['oper_id(重要站點)'][i] != '':   
            if type(spec['oper_id(重要站點)'][i]) == str:  #單一站點
                # oper title
                key_oper_title = key_oper[key_oper['OPER_ID']==spec['主旨帶入站點'][i]]
                key_oper_title = key_oper_title.reset_index(drop=True)
                ALL_TOOL = key_oper_title['UNIT_ID'].unique().tolist()
                ALL_TOOL = ",".join(ALL_TOOL)
                # oper 內文
                key_oper = key_oper[key_oper['OPER_ID']==spec['oper_id(重要站點)'][i]]
                key_oper = key_oper.reset_index(drop=True)
                    #修改站點名稱
                key_oper = key_oper.rename(columns={"UNIT_ID":key_oper['OPER_ID'].iloc[0]})
                key_oper = key_oper.drop(['EQP_ID','OPER_ID'],axis=1)
            else:   #多個站點
                # oper title
                if type(spec['主旨帶入站點'][i]) == str:
                    key_oper_title = key_oper[key_oper['OPER_ID']==spec['主旨帶入站點'][i]]
                    key_oper_title = key_oper_title.reset_index(drop=True)
                    ALL_TOOL = key_oper_title['UNIT_ID'].unique().tolist()
                    ALL_TOOL = ",".join(ALL_TOOL)
                else:
                    key_oper_title = key_oper[key_oper['OPER_ID'].isin(spec['主旨帶入站點'][i])]
                    key_oper_title = key_oper_title.reset_index(drop=True)
                    ALL_TOOL = key_oper_title['UNIT_ID'].unique().tolist()
                    ALL_TOOL = ",".join(ALL_TOOL)
                #oper 內文
                key_oper_filter = key_oper[key_oper['OPER_ID'].isin(spec['oper_id(重要站點)'][i])]
                oper_list = key_oper_filter['OPER_ID'].unique().tolist()  #帶出有的站點
                key_oper = key_oper_filter[key_oper_filter['OPER_ID']==oper_list[0]].reset_index(drop=True)       
                # 刪除"第一站"合併機台多餘欄位
                    #修改站點名稱
                key_oper = key_oper.rename(columns={"UNIT_ID":key_oper['OPER_ID'].iloc[0]})
                key_oper = key_oper.drop(['EQP_ID','OPER_ID'],axis=1)
                    #切割製程站點、機台、時間
                for a in range(len(oper_list)-1):
                    key_oper1 = key_oper_filter[key_oper_filter['OPER_ID']==oper_list[a+1]]['OPER_ID']
                    UNIT_ID = key_oper_filter[key_oper_filter['OPER_ID']==oper_list[a+1]]['UNIT_ID']
                    LOGOFF_TIME = key_oper_filter[key_oper_filter['OPER_ID']==oper_list[a+1]]['LOGOFF_TIME']
                    key_oper1 = pd.Series(key_oper1).reset_index(drop=True)
                    UNIT_ID = pd.Series(UNIT_ID).reset_index(drop=True)
                    LOGOFF_TIME = pd.Series(LOGOFF_TIME).reset_index(drop=True)
                    key_oper[key_oper1.iloc[0]]=UNIT_ID
                    key_oper['LOGOFF_TIME'+str(a+1)]=LOGOFF_TIME
                        #key_oper.insert(a+7,'LOGOFF_TIME'+str(a+1), LOGOFF_TIME)
                    #key_oper = pd.concat([key_oper,key_oper1],axis=1)
                    #key_oper['OPER_ID'+str(a+1)]=key_oper1
                    #key_oper['OPER_ID'+str(a+1)]=key_oper1[key_oper1['OPER_ID']==spec['oper_id(重要站點)'][i][a+1]].loc[:,'OPER_ID']
                    #key_oper['EQP_ID'+str(a+1)]=key_oper1[key_oper1['EQP_ID']==spec['oper_id(重要站點)'][i][a+1]].loc[:,'EQP_ID']
        elif spec['oper_id(重要站點)'][i] =='': #全部站點
                # oper title
            if type(spec['主旨帶入站點'][i]) == str:
                key_oper_title = key_oper[key_oper['OPER_ID'].str[:4]==spec['主旨帶入站點'][i]]
                key_oper_title = key_oper_title.reset_index(drop=True)
                #ALL_TOOL = key_oper_title['UNIT_ID'].unique().tolist()
                #ALL_TOOL = ",".join(ALL_TOOL)
                ALL_TOOL=''
            else:
                key_oper_title = key_oper[key_oper['OPER_ID'].str[:4].isin(spec['主旨帶入站點'][i])]
                key_oper_title = key_oper_title.reset_index(drop=True)
                #ALL_TOOL = key_oper_title['UNIT_ID'].unique().tolist()
                #ALL_TOOL = ",".join(ALL_TOOL)
                ALL_TOOL=''
            # 刪除"第一站"合併機台多餘欄位
            all_oper = key_oper['OPER_ID'].unique().tolist()
            key_oper_filter = key_oper[key_oper['OPER_ID'].isin(all_oper)]
            key_oper = key_oper[key_oper['OPER_ID']==all_oper[0]]
                #濾出所有機台
            key_oper = key_oper.rename(columns={"UNIT_ID":key_oper['OPER_ID'].iloc[0]})
            key_oper = key_oper.drop(['EQP_ID','OPER_ID'],axis=1)  
            #切割製程站點、機台、時間
            for a in range(len(all_oper)-1):
                key_oper1 = key_oper_filter[key_oper_filter['OPER_ID']==all_oper[a+1]]['OPER_ID']
                UNIT_ID = key_oper_filter[key_oper_filter['OPER_ID']==all_oper[a+1]]['UNIT_ID']
                LOGOFF_TIME = key_oper_filter[key_oper_filter['OPER_ID']==all_oper[a+1]]['LOGOFF_TIME']
                key_oper1 = pd.Series(key_oper1).reset_index(drop=True)
                UNIT_ID = pd.Series(UNIT_ID).reset_index(drop=True)
                LOGOFF_TIME = pd.Series(LOGOFF_TIME).reset_index(drop=True)
                key_oper[key_oper1.iloc[0]]=UNIT_ID
                #key_oper.insert(a+6,key_oper1.iloc[0], EQP_ID1)
                key_oper['LOGOFF_TIME'+str(a+1)]=LOGOFF_TIME
                #key_oper.insert(a+7,'LOGOFF_TIME'+str(a+1), LOGOFF_TIME)
        #抓出對應的資訊
        INFORMATION = key_oper.drop_duplicates()
            # defect code 照片
        picture = raw_data['IMG_FILE_PATH']
                # 散佈圖
        #plt.style.use('ggplot')  #使用ggplot美化
        plt.figure(figsize=(4,3.5))
        plt.scatter(df[( df['LOT_ID']==Density_BY_ABBR_OP.index[j][0]) & (df['DFT_CODE']==Density_BY_ABBR_OP.index[j][2])].POX_X/1000,
                        df[( df['LOT_ID']==Density_BY_ABBR_OP.index[j][0]) & (df['DFT_CODE']==Density_BY_ABBR_OP.index[j][2])].POX_Y/1000)
        plt.xlim(0, 1300) #固定X axis為1300
        plt.ylim(0, 1100) #固定y axis為1100
                         # figure 儲存為二進位檔案
        plt.grid(b=True, which='major', color='#666666', linestyle='--')
        buffer = BytesIO()
        plt.savefig(buffer)  
        plot_data = buffer.getvalue()

        # 圖片資料轉為 HTML 格式
        imb = base64.b64encode(plot_data)  
        ims = imb.decode()
        imd = "data:image/png;base64,"+ims
        scatter = "<h4>散佈圖"+"_"+Density_BY_ABBR_OP.index[j][2]+" & 九宮格照片</h4>" #+ """<img src=\'"%s\'> """ % imd   
         # 圖片存檔 轉為Base64
        plt.savefig('C:\\Users\\SHANESLU\\Desktop\\python\\AOI_scatter\\'+'散佈圖'+'_'+Density_BY_ABBR_OP.index[j][2]+'.png')
        plt.pause(0.01) #僅顯示0.01秒
        with open('C:\\Users\\SHANESLU\\Desktop\\python\\AOI_scatter\\'+'散佈圖'+'_'+Density_BY_ABBR_OP.index[j][2]+'.png', "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read())
            encoded_string = encoded_string.decode('utf-8')  #去掉 b'XXXX' 
#ManualSend_39  
        if len(picture)>=9:
                    picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" rowspan="3"><img src=\'"%s\'></td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd,picture.iloc[3],picture.iloc[4],
                                               picture.iloc[5],picture.iloc[6],picture.iloc[7],picture.iloc[8])
                    )
        elif len(picture)==8:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" rowspan="3"><img src=\'"%s\'></td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd,picture.iloc[3],picture.iloc[4],
                                               picture.iloc[5],picture.iloc[6],picture.iloc[7])
                     )
        elif len(picture)==7:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" rowspan="3"><img src=\'"%s\'></td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd,picture.iloc[3],picture.iloc[4],
                                               picture.iloc[5],picture.iloc[6])
                     )
        elif len(picture)==6:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" rowspan="2"><img src=\'"%s\'></td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd,picture.iloc[3],picture.iloc[4],
                                               picture.iloc[5])
                     )
        elif len(picture)==5:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" rowspan="2"><img src=\'"%s\'></td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd,picture.iloc[3],picture.iloc[4]
                                               )
                     )
        elif len(picture)==4:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" rowspan="2"><img src=\'"%s\'></td>
                                 </tr>
                                 <tr>
                                 <td>%s</td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd,picture.iloc[3]
                                               )
                     )
        elif len(picture)==3:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" colspan="3"><img src=\'"%s\'></td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],picture.iloc[2],imd)
                     )
        elif len(picture)==2:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td>%s</td>
                                 <td width="250" colspan="2"><img src=\'"%s\'></td>
                                 </tr>
                                </table>""" % (picture.iloc[0],picture.iloc[1],imd)
                     )
        elif len(picture)==1:
                     picture9=(
                     """<table border="1">
                                 <tr>
                                 <td>%s</td>
                                 <td width="250" colspan="2"><img src=\'"%s\'></td>
                                 </tr>
                                </table>""" % (picture.iloc[0],imd)
                     )
        if Density_BY_ABBR_OP.index[j][2][:1] == 'T':
            print('T & MFG')
            ManualSend_39 = {
                        'strMailCode': 'sDM8fpGZVQ8=', #MailCode
                        'strRecipients':spec['function'][i], #收件人
                        'strCopyRecipients':'andyou@auo.com', #副本
                        'strSubject':'AOI'+'_'+Density_BY_ABBR_OP.index[j][3]+'_'+Density_BY_ABBR_OP.index[j][0]+'_'+str(b)+' Suffered '+Density_BY_ABBR_OP.index[j][2]+' : '+str(ALL_TOOL)+' (單批異常立刻停機)', #標題
                        'strBody':'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][0]+
                        ' </font></b>'+'slot '+'<b><font color="#E00000">'+str(c)+'</font></b>'+
                        ' suffered '+'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][2]+
                        '</font></b>'+' issue.'+'<br>'+'<br>'+ 
                        """<table border="1">
                             <tr>
                             <td bgcolor="#FFB630">要求確認機台狀況(mail)</td>
                             <td bgcolor="#FFB630"><b>%s</b></td>
                             </tr>
                             <tr>
                             <td bgcolor="#DB0000"><font color="#FFFFFF">單批異常立刻停機(mail+hold)</td>
                             <td bgcolor="#DB0000"><b><font color="#FFFFFF">%s</font></b></td>
                             </tr>
                             <tr>
                             <td>Sheet Qty</td>
                             <td>%s</td>
                             </tr>
                             <tr>
                             <td>Defect Qty</td>
                             <td>%s</td>
                             </tr>
                              <tr>
                             <td bgcolor="#00DB00"><font color="#0000D6">Density</td>
                             <td bgcolor="#00DB00"><b><font color="#0000D6">%s</b></td>
                             </tr>
                            </table>""" % (str(spec['mail'][i]),str(spec['mail+hold'][i]),
                                           str(LOT_COUNT['LOT_COUNT'][Density_BY_ABBR_OP.index[j][0]]),
                                           str(DFT_Cnt['DFT_Cnt'][Density_BY_ABBR_OP.index[j][0]][Density_BY_ABBR_OP.index[j][2]]),
                                           str(Density_BY_ABBR_OP[j]))+
                         "<h2>Information</h2>"+str(INFORMATION.to_html())+        
                        scatter+
                            str(picture9)+
                        "<h2>EDA Link</h2>"+
                '<a href="http://autceda/dashboard/sites/L5D/show/L5D_DS_TEST_Array_T1_Defect_Density_Light/PVD_KO_Density_Light?title=Density%20LIGHT&layers=160&layers=321&layers=368&">ML5DT1</a>'+
                "<br>"+'<a href="http://autceda/dashboard/sites/L5D/show/L5D_Array_INT_7D_InlineSummary/7D_InlineSummary?Defect Code=%s">InlineSummary</a>'%(str(Density_BY_ABBR_OP.index[j][2]))+
                        "<h2>Raw data(抓前50筆)</h2>"+str(raw_data[:50].to_html()), #raw data 抓前50筆 '
                        'strFileBase64String':'散佈圖'+'_'+Density_BY_ABBR_OP.index[j][2]+'.png:'+str(encoded_string) #夾檔案 需先將圖片轉base64， 格式要為: 'XXX.png' +   ':' + encoded_string
                        }
            response_39 = client.service.ManualSend_39(**ManualSend_39)
            now = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            for eventid in spec['EVENTID'][i]:  #電話語音發送
                xml_ftp_upload('站點'+Density_BY_ABBR_OP.index[j][3]+Density_BY_ABBR_OP.index[j][2]+'機台'+ALL_TOOL+'超規需停機',eventid,now,'AIDI','B','0','AIDI停機通知','機台'+ALL_TOOL+'超規需停機',Density_BY_ABBR_OP.index[j][3]+Density_BY_ABBR_OP.index[j][2]+eventid[7:13])
            print('是否成功發送mail: '+str(response_39))
            print()
        elif Density_BY_ABBR_OP.index[j][2][:1] == 'P':
            print('P & MFG')
            ManualSend_39 = {
                        'strMailCode': 'sDM8fpGZVQ8=', #MailCode
                        'strRecipients':spec['function'][i], #收件人
                        'strCopyRecipients':'andyou@auo.com', #副本
                        'strSubject':'AOI'+'_'+Density_BY_ABBR_OP.index[j][3]+'_'+Density_BY_ABBR_OP.index[j][0]+'_'+str(b)+' Suffered '+Density_BY_ABBR_OP.index[j][2]+' : '+str(ALL_TOOL)+' (單批異常立刻停機)', #標題
                        'strBody':'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][0]+
                        ' </font></b>'+'slot '+'<b><font color="#E00000">'+str(c)+'</font></b>'+
                        ' suffered '+'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][2]+
                        '</font></b>'+' issue.'+'<br>'+'<br>'+ 
                        """<table border="1">
                             <tr>
                             <td bgcolor="#FFB630">要求確認機台狀況(mail)</td>
                             <td bgcolor="#FFB630"><b>%s</b></td>
                             </tr>
                             <tr>
                             <td bgcolor="#DB0000"><font color="#FFFFFF">單批異常立刻停機(mail+hold)</td>
                             <td bgcolor="#DB0000"><b><font color="#FFFFFF">%s</font></b></td>
                             </tr>
                             <tr>
                             <td>Sheet Qty</td>
                             <td>%s</td>
                             </tr>
                             <tr>
                             <td>Defect Qty</td>
                             <td>%s</td>
                             </tr>
                              <tr>
                             <td bgcolor="#00DB00"><font color="#0000D6">Density</td>
                             <td bgcolor="#00DB00"><b><font color="#0000D6">%s</b></td>
                             </tr>
                            </table>""" % (str(spec['mail'][i]),str(spec['mail+hold'][i]),
                                           str(LOT_COUNT['LOT_COUNT'][Density_BY_ABBR_OP.index[j][0]]),
                                           str(DFT_Cnt['DFT_Cnt'][Density_BY_ABBR_OP.index[j][0]][Density_BY_ABBR_OP.index[j][2]]),
                                           str(Density_BY_ABBR_OP[j]))+
                         "<h2>Information</h2>"+str(INFORMATION.to_html())+        
                        scatter+
                            str(picture9)+
                        "<h2>EDA Link</h2>"+
                        '<a href="http://autceda/dashboard/sites/L5D/show/L5D_Array_INT_7D_InlineSummary/7D_InlineSummary?Defect Code=%s">InlineSummary</a>'%(str(Density_BY_ABBR_OP.index[j][2]))+
                        "<h2>Raw data(抓前50筆)</h2>"+str(raw_data[:50].to_html()), #raw data 抓前50筆 '
                        'strFileBase64String':'散佈圖'+'_'+Density_BY_ABBR_OP.index[j][2]+'.png:'+str(encoded_string) #夾檔案 需先將圖片轉base64， 格式要為: 'XXX.png' +   ':' + encoded_string
                        }
            response_39 = client.service.ManualSend_39(**ManualSend_39)
            now = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            for eventid in spec['EVENTID'][i]:  #電話語音發送
                xml_ftp_upload('站點'+Density_BY_ABBR_OP.index[j][3]+Density_BY_ABBR_OP.index[j][2]+'機台'+ALL_TOOL+'超規需停機',eventid,now,'AIDI','B','0','AIDI停機通知','機台'+ALL_TOOL+'超規需停機',Density_BY_ABBR_OP.index[j][3]+Density_BY_ABBR_OP.index[j][2]+eventid[7:13])
            print('是否成功發送mail: '+str(response_39))
            print()
        elif Density_BY_ABBR_OP.index[j][2][:1] == 'E':
            print('E & MFG')
            ManualSend_39 = {
                        'strMailCode': 'sDM8fpGZVQ8=', #MailCode
                        'strRecipients':spec['function'][i], #收件人
                        'strCopyRecipients':'andyou@auo.com', #副本
                        'strSubject':'AOI'+'_'+Density_BY_ABBR_OP.index[j][3]+'_'+Density_BY_ABBR_OP.index[j][0]+'_'+str(b)+' Suffered '+Density_BY_ABBR_OP.index[j][2]+' : '+str(ALL_TOOL)+' (單批異常立刻停機)', #標題
                        'strBody':'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][0]+
                        ' </font></b>'+'slot '+'<b><font color="#E00000">'+str(c)+'</font></b>'+
                        ' suffered '+'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][2]+
                        '</font></b>'+' issue.'+'<br>'+'<br>'+ 
                        """<table border="1">
                             <tr>
                             <td bgcolor="#FFB630">要求確認機台狀況(mail)</td>
                             <td bgcolor="#FFB630"><b>%s</b></td>
                             </tr>
                             <tr>
                             <td bgcolor="#DB0000"><font color="#FFFFFF">單批異常立刻停機(mail+hold)</td>
                             <td bgcolor="#DB0000"><b><font color="#FFFFFF">%s</font></b></td>
                             </tr>
                             <tr>
                             <td>Sheet Qty</td>
                             <td>%s</td>
                             </tr>
                             <tr>
                             <td>Defect Qty</td>
                             <td>%s</td>
                             </tr>
                              <tr>
                             <td bgcolor="#00DB00"><font color="#0000D6">Density</td>
                             <td bgcolor="#00DB00"><b><font color="#0000D6">%s</b></td>
                             </tr>
                            </table>""" % (str(spec['mail'][i]),str(spec['mail+hold'][i]),
                                           str(LOT_COUNT['LOT_COUNT'][Density_BY_ABBR_OP.index[j][0]]),
                                           str(DFT_Cnt['DFT_Cnt'][Density_BY_ABBR_OP.index[j][0]][Density_BY_ABBR_OP.index[j][2]]),
                                           str(Density_BY_ABBR_OP[j]))+
                         "<h2>Information</h2>"+str(INFORMATION.to_html())+        
                        scatter+
                            str(picture9)+
                        "<h2>EDA Link</h2>"+
                '<a href="http://autceda/dashboard/sites/L5D/show/L5D_DS_TEST_Array_E1_Defect_Density_Light/DRY_KO_Density_Light?title=Density_Light&layers=160&layers=321&layers=486&">ML5DE1</a>'+
                "<br>"+'<a href="http://autceda/dashboard/sites/L5D/show/L5D_DS_TEST_Array_E2_Defect_Density_Light/WET_KO_Density_Light?title=Defect_Light&layers=160&layers=321&layers=540&">ML5DE2</a>'+
                "<br>"+'<a href="http://autceda/dashboard/sites/L5D/show/L5D_Array_INT_7D_InlineSummary/7D_InlineSummary?Defect Code=%s">InlineSummary</a>'%(str(Density_BY_ABBR_OP.index[j][2]))+
                        "<h2>Raw data(抓前50筆)</h2>"+str(raw_data[:50].to_html()), #raw data 抓前50筆 '
                        'strFileBase64String':'散佈圖'+'_'+Density_BY_ABBR_OP.index[j][2]+'.png:'+str(encoded_string) #夾檔案 需先將圖片轉base64， 格式要為: 'XXX.png' +   ':' + encoded_string
                        }
            response_39 = client.service.ManualSend_39(**ManualSend_39)
            now = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            for eventid in spec['EVENTID'][i]:  #電話語音發送
                xml_ftp_upload('站點'+Density_BY_ABBR_OP.index[j][3]+Density_BY_ABBR_OP.index[j][2]+'機台'+ALL_TOOL+'超規需停機',eventid,now,'AIDI','B','0','AIDI停機通知','機台'+ALL_TOOL+'超規需停機',Density_BY_ABBR_OP.index[j][3]+Density_BY_ABBR_OP.index[j][2]+eventid[7:13])
            print('是否成功發送mail: '+str(response_39))
            print()
        elif Density_BY_ABBR_OP.index[j][2][:1] == 'I':
            print('I & MFG')
            ManualSend_39 = {
                        'strMailCode': 'sDM8fpGZVQ8=', #MailCode
                        'strRecipients':spec['function'][i], #收件人
                        'strCopyRecipients':'andyou@auo.com', #副本
                        'strSubject':'AOI'+'_'+Density_BY_ABBR_OP.index[j][3]+'_'+Density_BY_ABBR_OP.index[j][0]+'_'+str(b)+' Suffered '+Density_BY_ABBR_OP.index[j][2]+' : '+str(ALL_TOOL)+' (單批異常立刻停機)', #標題
                        'strBody':'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][0]+
                        ' </font></b>'+'slot '+'<b><font color="#E00000">'+str(c)+'</font></b>'+
                        ' suffered '+'<b><font color="#E00000">'+Density_BY_ABBR_OP.index[j][2]+
                        '</font></b>'+' issue.'+'<br>'+'<br>'+ 
                        """<table border="1">
                             <tr>
                             <td bgcolor="#FFB630">要求確認機台狀況(mail)</td>
                             <td bgcolor="#FFB630"><b>%s</b></td>
                             </tr>
                             <tr>
                             <td bgcolor="#DB0000"><font color="#FFFFFF">單批異常立刻停機(mail+hold)</td>
                             <td bgcolor="#DB0000"><b><font color="#FFFFFF">%s</font></b></td>
                             </tr>
                             <tr>
                             <td>Sheet Qty</td>
                             <td>%s</td>
                             </tr>
                             <tr>
                             <td>Defect Qty</td>
                             <td>%s</td>
                             </tr>
                              <tr>
                             <td bgcolor="#00DB00"><font color="#0000D6">Density</td>
                             <td bgcolor="#00DB00"><b><font color="#0000D6">%s</b></td>
                             </tr>
                            </table>""" % (str(spec['mail'][i]),str(spec['mail+hold'][i]),
                                           str(LOT_COUNT['LOT_COUNT'][Density_BY_ABBR_OP.index[j][0]]),
                                           str(DFT_Cnt['DFT_Cnt'][Density_BY_ABBR_OP.index[j][0]][Density_BY_ABBR_OP.index[j][2]]),
                                           str(Density_BY_ABBR_OP[j]))+
                         "<h2>Information</h2>"+str(INFORMATION.to_html())+        
                        scatter+
                            str(picture9)+
                        "<h2>EDA Link</h2>"+
                '<a href="http://autceda/dashboard/sites/L5D/show/L5D_DS_TEST_Array_T1_Defect_Density_Light/PVD_KO_Density_Light?title=Density%20LIGHT&layers=160&layers=321&layers=368&">ML5DT1</a>'+
                "<br>"+
                '<a href="http://autceda/dashboard/sites/L5D/show/L5D_DS_TEST_Array_E1_Defect_Density_Light/DRY_KO_Density_Light?title=Density_Light&layers=160&layers=321&layers=486&">ML5DE1</a>'+
                "<br>"+'<a href="http://autceda/dashboard/sites/L5D/show/L5D_DS_TEST_Array_E2_Defect_Density_Light/WET_KO_Density_Light?title=Defect_Light&layers=160&layers=321&layers=540&">ML5DE2</a>'+
                "<br>"+'<a href="http://autceda/dashboard/sites/L5D/show/L5D_Array_INT_7D_InlineSummary/7D_InlineSummary?Defect Code=%s">InlineSummary</a>'%(str(Density_BY_ABBR_OP.index[j][2]))+
                        "<h2>Raw data(抓前50筆)</h2>"+str(raw_data[:50].to_html()), #raw data 抓前50筆 '
                        'strFileBase64String':'散佈圖'+'_'+Density_BY_ABBR_OP.index[j][2]+'.png:'+str(encoded_string) #夾檔案 需先將圖片轉base64， 格式要為: 'XXX.png' +   ':' + encoded_string
                        }
            response_39 = client.service.ManualSend_39(**ManualSend_39)
            now = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            for eventid in spec['EVENTID'][i]:  #電話語音發送
                xml_ftp_upload('站點'+Density_BY_ABBR_OP.index[j][3]+Density_BY_ABBR_OP.index[j][2]+'機台'+ALL_TOOL+'超規需停機',eventid,now,'AIDI','B','0','AIDI停機通知','機台'+ALL_TOOL+'超規需停機',Density_BY_ABBR_OP.index[j][3]+Density_BY_ABBR_OP.index[j][2]+eventid[7:13])
            print('是否成功發送mail: '+str(response_39))
            print()
    return


# # Density by Defect code，卡spec，過標才抓出來

# In[53]:


client = zeep.Client("http://ids.cdn.corpnet.auo.com/IDS_WS/Mail.asmx?wsdl") #連結IDS

for j in range(len(Density_BY_ABBR_OP)):
    for i in range(len(spec['Defect_Code'])):  
        if (Density_BY_ABBR_OP.index[j][2] == spec['Defect_Code'][i]) &  (spec['op_id'][i]=="nan" ) & (Density_BY_ABBR_OP.index[j][1] == spec['abbr_no'][i]): #沒有對到站點，有對到model(針對model類型)
                    print("no op_id & but model",Density_BY_ABBR_OP.index[j],spec['Defect_Code'][i],spec['op_id'][i],spec['oper_id(重要站點)'][i],spec['abbr_no'][i])
                    print(j,i,Density_BY_ABBR_OP.index[j][2],spec['Defect_Code'][i])
                    judge(i,j)
                    break
        elif (Density_BY_ABBR_OP.index[j][2] == spec['Defect_Code'][i]) & (Density_BY_ABBR_OP.index[j][3] == spec['op_id'][i]) & (spec['abbr_no'][i]=="" ): #有對到站點，沒對到model(針對站點類型)
                    print("op_id but no model",Density_BY_ABBR_OP.index[j],spec['Defect_Code'][i],spec['op_id'][i],spec['oper_id(重要站點)'][i],spec['abbr_no'][i])
                    print(j,i,Density_BY_ABBR_OP.index[j][2],spec['Defect_Code'][i])
                    judge(i,j)
                    break
        elif (Density_BY_ABBR_OP.index[j][2] == spec['Defect_Code'][i]) & (Density_BY_ABBR_OP.index[j][3] == spec['op_id'][i]) & (Density_BY_ABBR_OP.index[j][1] == spec['abbr_no'][i]): #有對到站點，有對到model(針對model+站點類型)
                    print("op_id & model",Density_BY_ABBR_OP.index[j],spec['Defect_Code'][i],spec['op_id'][i],spec['oper_id(重要站點)'][i],spec['abbr_no'][i])
                    print(j,i,Density_BY_ABBR_OP.index[j][2],spec['Defect_Code'][i])
                    judge(i,j)
                    break
        elif (Density_BY_ABBR_OP.index[j][2] == spec['Defect_Code'][i]) & (spec['op_id'][i]=="nan") & (spec['abbr_no'][i]==""):  #沒有對到站點和model(一般類型)
                    print("nothing",Density_BY_ABBR_OP.index[j],spec['Defect_Code'][i],spec['op_id'][i],spec['oper_id(重要站點)'][i])
                    print(j,i,Density_BY_ABBR_OP.index[j][2],spec['Defect_Code'][i])
                    judge(i,j)
                    break
