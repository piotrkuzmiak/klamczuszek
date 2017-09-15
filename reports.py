
# coding: utf-8

import sys
from teradata_connection import KHD
from parsers import Orange, Ddi, Netia, Ip
import pandas as pd
from decorator_for_timing import estimate_time
from excel_report import Excel_report
import numpy as np
#%%
#data przekazana do skryptu
if len(sys.argv) < 2:
    data_range = None
else:
    data_range = sys.argv[1]
#lista z instancjami parserow
parsers_list = []

#iteracja po plikach
import glob

supplier_list = [{'orange':{'path':'D:/CRM/klamczuszek/bilingi','extension':'xlsx'}},\
        {'ddi':{'path':'D:/CRM/klamczuszek/Biling_DDI', 'extension':'txt' }}, \
        {'netia':{'path':'D:/CRM/klamczuszek/Netia','extension':'csv'}}, \
        {'ip':{'path':'D:/CRM/klamczuszek/IP','extension':'txt'}}, \
        {'polkomtel':{'path':'D:/CRM/klamczuszek/polkomtel','extension':'xlsx'}},
        {'mapa':{'path':'D:/CRM/klamczuszek/mapa', 'extension':'xlsx'}}]
for sup_dict in supplier_list:
    for k, v in sup_dict.items():
#        print(k)
        all_files = glob.glob(v['path'] + '/*.'+v['extension'])
        if 'orange' in k:
#            print(all_files)
            [parsers_list.append(Orange(file_name=x)) for x in all_files]
        if 'ip' in k:
#            print(all_files)
            [parsers_list.append(Ip(file_name=x)) for x in all_files]
        if 'ddi' in k:
#            print(all_files)
            [parsers_list.append(Ddi(file_name=x)) for x in all_files]
        if 'netia' in k:
            [parsers_list.append(Netia(file_name=x)) for x in all_files]
        if 'mapa' in k:
            df_mapa = pd.read_excel(all_files[0], dtype={'Numer telefonu':np.str})
            df_mapa.rename(inplace=True, columns={'Numer osobowy':'SKP', 'Numer telefonu':'MOBILE_NO_DOR'})
            df_mapa['SKP'] = df_mapa.SKP.astype(dtype='int')
            df_mapa['MOBILE_NO_DOR'] = df_mapa.MOBILE_NO_DOR.str.strip()
#Wspolny zbior dla wszystkich
#TODO:dodac przetwarzanie rownolegle w Pool
def clear_phone(number):
    """
        Funkcja dodaje '48' do numerow telefonow komorkowych.
        Dziala dla str jesli maja ciag dlugosci 9 lub 10 znakow.
        
        Parameters
        ----------
        number:str
            numer telefonu komorkowego w postaci ciagu znakow
            
        Returns
        -------
        number:str
            zmieniony numer telefonu
        Notes:
        -----
        None
        """
    if len(number)==10:
        if number.startswith('0'):
            number=number.replace('0', '48', 1)
    else:
        if len(number)==9:
            number='48'+str(number)
    return number
print('Zaczytuje bilingi...')            
@estimate_time
def all_connections():
#    oczyszczanie danych z bilingow
    df_all_connections=pd.concat([x.df_connections() for x in parsers_list])
    df_all_connections.drop_duplicates(subset=['NUMER_A','NUMER_B'],inplace=True)
    df_all_connections['NUMER_A'] = df_all_connections.NUMER_A.str.strip()
    df_all_connections['NUMER_A'] = df_all_connections.NUMER_A.apply(lambda x: clear_phone(str(x)))
    df_all_connections['NUMER_B'] = df_all_connections.NUMER_B.str.strip()
    df_all_connections['NUMER_B'] = df_all_connections.NUMER_B.apply(lambda x: clear_phone(str(x)))
    df_all_connections = df_all_connections[df_all_connections.NUMER_B.str.len()>7]
    return df_all_connections

df_all_connections = all_connections()
#wyciagniecie danych o doradcy z khd
sql_doradcy = "SELECT DISTINCT \
  CICNTID AS CONTACT_ID\
, cicif AS CIF\
, CICNTDTE AS CONTACT_DATE\
, UD.SKP\
, UD.PHONE_NO  \
, UD.MOBILE_NO  \
,ouh.BRANCH_NAME AS oddzial\
,ouh.REGION_NAME AS region\
,ouh.MACROREGION_NAME AS makroregion\
    FROM MISV01P.acup61001 acup\
        JOIN misv01p.USERS_DETAIL ud\
        ON ud.skp = acup.ciskprlze\
        JOIN ddhdv01p.ORGANISATION_UNIT_HIERARCHY ouh\
        ON ouh.WH_ORG_UNIT_ID=ud.ICBS_NO\
    WHERE CIMEDCOM = 'P'\
    AND CICNTDTE >= '2017-06-01'\
    AND ( UD.PHONE_NO IS NOT NULL OR UD.MOBILE_NO IS NOT NULL)\
"
#kryterium daty na khd
if data_range == None:
    data_for_khd = '2017-05-31'
else:
    data_for_khd = data_range
    
sql_kontakty="SELECT DISTINCT \
/* acup.CICNTID AS CONTACT_ID*/ \
ouh.BRANCH_NAME AS oddzial\
,ouh.REGION_NAME AS region\
,ouh.MACROREGION_NAME AS makroregion\
,acup.ciskprlze as SKP \
, acup.cicif as CIF \
/*, acup.CICNTDTE AS CONTACT_DATE*/ \
, offer_type_nm as NAZWA_KAMPANII \
, LS.OFFER_TYPE_CD as offer_type_cd \
, csd.MOB_PHONE_NO as MOBILE_NO_CIF \
, csd.COMP_HOME_PH_NO as PHONE_NO_CIF \
FROM MISV01P.acup61001 acup \
JOIN misv01p.crm_lead cl \
on  cl.contact_id = acup.cicntid \
JOIN MISV01P.LEAD_STATUS LS \
ON LS.LEAD_KEY = CL.LEAD_ID \
JOIN MISV01P.OFFER_STATUS OS \
ON LS.CIF_KEY = OS.CIF_KEY \
AND LS.OFFER_TYPE_CD = OS.OFFER_TYPE_CD \
AND LS.VALID_TO <= OS.VALID_TO \
AND LS.VALID_FROM >= OS.VALID_FROM \
AND OS.OFFER_TYPE_CD NOT LIKE '%TEST%' \
AND (OS.OFFER_GRP_CD NOT LIKE '%ZW%' OR OS.OFFER_GRP_CD IS NULL) \
AND OS.CAMP_EXT_EXEC_MODE = 'P' \
AND OS.OFFER_LOCK_FLG = 'N' \
JOIN MISV01P.B_DICT_OFFER_TYPES OT \
ON OT.OFFER_TYPE_CD = LS.OFFER_TYPE_CD \
AND \
(\
LS.VALID_FROM BETWEEN OT.START_DT AND OT.END_DT \
)\
JOIN misv01p.USERS_DETAIL ud \
ON ud.skp = acup.ciskprlze \
JOIN ddhdv01p.ORGANISATION_UNIT_HIERARCHY ouh \
ON ouh.WH_ORG_UNIT_ID=ud.ICBS_NO \
LEFT JOIN \
( \
SELECT DISTINCT \
CELL_PACKAGE_SK \
,MARKETING_CELL_NM \
,CASE WHEN \
UPPER(MARKETING_CELL_NM) LIKE '%MODEL%' \
AND UPPER(MARKETING_CELL_NM) NOT LIKE '%PROBKA%' \
AND UPPER(MARKETING_CELL_NM) NOT LIKE '%PRBKA%' \
THEN 1 ELSE 0 END AS MODEL_FLG \
,CASE WHEN MODEL_FLG = 1 \
THEN SUBSTR( \
MARKETING_CELL_NM \
,1 \
,COALESCE(NULLIFZERO(POSITION(' ' IN MARKETING_CELL_NM))-1,LENGTH(MARKETING_CELL_NM)) \
) \
END AS MODEL_NM \
FROM \
MISV01P.CI_CELL_PACKAGE \
WHERE \
MODEL_FLG = 1 \
)CCP \
ON CCP.CELL_PACKAGE_SK = LS.CELL_PACKAGE_SK \
join ddhdv08p.customer_detail2 cd2 \
ON         CD2.SRCE_CUST_NO = ls.CIF_KEY \
AND       substr(CD2.SRCE_CUST_NO,1,1) not in ('O','B')             \
join ddhdv08p.CUSTOMER_SUB_DETAIL2_BZWBK csd \
on csd.wh_cust_no = cd2.wh_cust_no \
WHERE \
LS.CAMP_EXT_EXEC_MODE = 'P' \
AND  LS.OFFER_TYPE_CD NOT LIKE '%TEST%' \
AND  ( LS.OFFER_GRP_CD NOT LIKE '%ZW%' OR LS.OFFER_GRP_CD IS NULL) \
AND  ( LS.COMM_EXT_BIZ_CODE NOT LIKE '%ZamrBazy%' OR LS.COMM_EXT_BIZ_CODE IS NULL ) \
AND  ( LS.LOCKED_REASON_CD NOT IN ('24', '25', '26', '27', '28') OR ls.LOCKED_REASON_CD IS NULL) \
AND LS.SLOT_CD IN('O','P') \
AND LAST_DAY(ACUP.CICNTDTE) = '"+ data_for_khd +"' \
AND ACUP.CIMEDCOM = 'P' \
AND OFFER_SEGMENT = 'IND' \
"
@estimate_time
def ask_khd():
    print('odpytuje KHD...')
    khd = KHD(pass_string='d:/programowanie/python/jupyter_notebooks/teraAuth.txt',system_name='KHD_LIVE')
    return khd.query_to_df(query_str=sql_kontakty)
df_doradcy_khd = ask_khd()
#oczyszczanie danych
df_doradcy_khd.SKP = df_doradcy_khd.SKP.astype(dtype='str', copy=False)
df_doradcy_khd.SKP = df_doradcy_khd.SKP.str.strip()
df_doradcy_khd.SKP = df_doradcy_khd.SKP.astype(dtype='int', copy=False)
df_doradcy_khd.PHONE_NO_CIF = df_doradcy_khd.PHONE_NO_CIF.str.replace('-','')
df_doradcy_khd.PHONE_NO_CIF = df_doradcy_khd.PHONE_NO_CIF.str.replace(' ','')
df_doradcy_khd.PHONE_NO_CIF = df_doradcy_khd.PHONE_NO_CIF.str.replace('(','')
df_doradcy_khd.PHONE_NO_CIF = df_doradcy_khd.PHONE_NO_CIF.str.replace(')','')
df_doradcy_khd.PHONE_NO_CIF = df_doradcy_khd.PHONE_NO_CIF.str.strip()
df_doradcy_khd.MOBILE_NO_CIF = df_doradcy_khd.MOBILE_NO_CIF.str.strip()
df_doradcy_khd.PHONE_NO_CIF = df_doradcy_khd.PHONE_NO_CIF.apply(lambda x : clear_phone(str(x)))
df_doradcy_khd.MOBILE_NO_CIF = df_doradcy_khd.MOBILE_NO_CIF.apply(lambda x : clear_phone(str(x)))
df_mapa.MOBILE_NO_DOR = df_mapa.MOBILE_NO_DOR.str.strip()
df_mapa.MOBILE_NO_DOR = df_mapa.MOBILE_NO_DOR.apply(lambda x : clear_phone(str(x)))

#symulowanie obecnosci numerow w bilingach:
df_all_connections.set_value(0, 'NUMER_A', '48669930522')
df_all_connections.set_value(19, 'NUMER_A', '48669930843')
df_all_connections.set_value(0, 'NUMER_B', '48605646092')
df_all_connections.replace(to_replace='48517474040',value='48669930522', inplace=True)
#kryterium daty ze skryptu z ograniczeniem na wczytane bilingi
df_all_connections=df_all_connections[df_all_connections.DATA_POCZ.str.startswith(data_for_khd)]
#laczenie inner mapa z bilingiem
#TODO:dolaczyc dane ze stacjonarnych

#tylko te SKP, ktore sa w df_mapa:
#konieczne podpiece stacjonarnych jeszcze
df_skp_in_mob = df_mapa.merge\
(right=df_all_connections, right_on='NUMER_A', left_on='MOBILE_NO_DOR', how='inner')[['SKP','NUMER_A', 'NUMER_B']]
df_skp_in_phones = df_skp_in_mob.copy()
df_skp_in_phones.NUMER_B = ''
df_skp_in_phones.set_value(2,'SKP', 101364)
df_skp_in_phones.set_value(2,'NUMER_B', '48717959116')
df_skp_in_phones.set_value(4,'SKP', 101364)
df_skp_in_phones.set_value(4,'NUMER_B', '48713284122')
df_skp_in_phones.set_value(1,'SKP', 101367)
df_skp_in_phones.set_value(1,'NUMER_B', '48756418635')
df_skp_in_phones.set_value(1,'SKP', 101367)
df_skp_in_phones.set_value(1,'NUMER_B', '48413761468')
#blad logiczny ponizej
#join musi byc robiony po mobile lub stacjonarny dla tego co wyszlo z khd i linijki powyzej:
#https://stackoverflow.com/questions/43925603/python-pandas-merge-with-or-logic
df_merged_doradcy_biling = df_doradcy_khd.merge(right=df_skp_in_mob,\
                            left_on=['SKP','MOBILE_NO_CIF'],right_on=['SKP','NUMER_B'], how='left')\
                            [['makroregion','region','oddzial','SKP','CIF',\
                            'NAZWA_KAMPANII','offer_type_cd','MOBILE_NO_CIF','PHONE_NO_CIF',\
                            'NUMER_B']]
df_merged_doradcy_biling = df_merged_doradcy_biling.merge(right=df_skp_in_phones,\
                            left_on=['SKP','PHONE_NO_CIF'],right_on=['SKP','NUMER_B'], how='left', suffixes=('_mobile','_wired'))
                    
df_to_raport = df_merged_doradcy_biling[(df_merged_doradcy_biling.NUMER_B_mobile.notnull() \
                | df_merged_doradcy_biling.NUMER_B_wired.notnull())][['SKP','NAZWA_KAMPANII','MOBILE_NO_CIF','PHONE_NO_CIF','NUMER_B_mobile','NUMER_B_wired']]                            
#podsumowania:
df_merged_doradcy_biling['khd_info'] = df_merged_doradcy_biling[['MOBILE_NO_CIF','PHONE_NO_CIF']].count(axis=1)                            
df_merged_doradcy_biling['billings'] = df_merged_doradcy_biling[['NUMER_B_mobile','NUMER_B_wired']].count(axis=1)
#zamiana na bool dla pozniejszego podsumowania tylko wartosci True
df_merged_doradcy_biling.loc[:,['khd_info','billings']]=df_merged_doradcy_biling[['khd_info','billings']].astype(np.int)
df_merged_doradcy_biling.loc[:,['khd_info','billings']]=df_merged_doradcy_biling[['khd_info','billings']].astype(np.int)
df_lier = df_merged_doradcy_biling[['makroregion','region','oddzial','SKP','NAZWA_KAMPANII','offer_type_cd','khd_info','billings']]
#usuwanie zbednych spacji dla wybranych kolumn
df_lier.loc[:,['makroregion','region','oddzial']] = df_lier.loc[:,['makroregion','region','oddzial']].applymap(lambda x : x.strip())
df_lier['% oznaczonych kontaktów'] = df_lier.apply(lambda x : x['billings']/x['khd_info'], axis=1)

#budowa tabeli przestawnej:

#piv_tab.reindex_axis(labels=['makroregion','region','oddzial','SKP','NAZWA_KAMPANII','offer_type_cd','khd_info','billings','% oznaczonych kontaktów'],axis=1, copy=False)

#                                   .groupby(by=['makroregion','region','oddzial','SKP','NAZWA_KAMPANII','offer_type_cd']).sum()
#                                   .reset_index(level=[0,1])
#df_merged_doradcy_biling['billings'] = df_merged_doradcy_biling['billings'].replace(0,np.nan)
#df_raport=df_merged_doradcy_biling[['SKP','NAZWA_KAMPANII','khd_info', 'billings']].groupby(by=['SKP','NAZWA_KAMPANII']).sum()
#df_merged_doradcy_biling['%oznaczonych'] = sum(df_merged_doradcy_biling.billings/df_merged_doradcy_biling.khd_info)
#print('jeszcze chwila, teraz przygotowuje raport w Excel...')

#%%
df_from_hdf = pd.read_hdf('df_lier.hdf')
excel = Excel_report(dataframe=df_from_hdf, groupby=['makroregion','region','oddzial','SKP','NAZWA_KAMPANII','offer_type_cd'])
#pv=excel.make_pivot(index=['makroregion','skp', 'kampania'], values=['rejestr','biling'])
#excel.unload('D:/CRM/wymiana/klamczuszek.xlsx','Raport', pivot_table=pv)
excel.unload('D:/CRM/wymiana/klamczuszek.xlsx')
#%%
import pandas as pd
import numpy as np
df = pd.DataFrame([['zachod',129175,'sprzedaz',True, True],\
                   ['wschod',118158,'sprzedaz', True, False],\
                   ['zachod',129175,'sprzedaz',True,False],\
                   ['zachod',129175,'konto',True,True],\
                   ['zachod',129175,'sprzedaz',True,True],\
                   ['zachod',129175,'sprzedaz',True,True],
                   ['poludnie',130115,'konto',True,True]], columns=['makroregion','skp','kampania', 'rejestr', 'biling'])
