#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import re
import win32com.client
import docx
import pandas as pd
import nltk
import numpy as np

pd.set_option('display.max_rows',None)
pd.set_option('display.max_colwidth',None)


# In[8]:


#Goals: 변수 및 Paragraph 선택

def level_pre_processing(data):
    """
    주어진 데이터를 필터링하고 관련 열만 선택하여 전처리합니다.

    매개변수:
    data (DataFrame): 전처리할 입력 데이터.

    반환값:
    DataFrame: 관련 열만 포함하는 전처리된 데이터.
    """
    # 'Type'이 'Paragraph'인 행을 필터링
    df_paragrpah=data[data['Type']=='Paragraph']
    
    # 관련 열 선택
    columns=['ID','Text','Style','Numbering','Numbering Type','Level','SECTION','Parent Index','Start','End','File ID']
    df_paragrpah=df_paragrpah[columns]
    
    # 결측치 인덱스 선택
    null_index=df_paragrpah[df_paragrpah[['Numbering']].isnull().any(axis=1)].index

    df_paragrpah.loc[null_index,'Numbering']='No Type'
    df_paragrpah.loc[null_index,'Numbering Type']='No Type'
    
    return df_paragrpah


# In[9]:


# Goals: 검증 데이터 프레임 

def level_update_numbering(data,index,numbeing_type,numbering):
    """
    주어진 데이터 프레임에서 특정 인덱스의 'Numbering' 값을 업데이트합니다.

    Args:
    data (DataFrame): 업데이트할 데이터 프레임.
    index (int or list): 'Numbering' 값을 업데이트할 행의 인덱스.
    numbering (any): 새로운 'Numbering' 값.

    Returns:
    DataFrame: 업데이트된 데이터 프레임.
    """
    data.at[index,'Numbering Type']=numbeing_type
    data.at[index,'Numbering']=numbering
        
    return data


# In[10]:


def verification_level_numberingtype(data,level_bowl):
    """
    데이터 프레임에서 'Level'과 'Numbering Type'의 일관성을 검증합니다.

    Args:
    data (DataFrame): 검증할 데이터 프레임.
    level_bowl (list): 검증할 레벨들을 포함하는 리스트.

    Returns:
    DataFrame: 검증 결과가 업데이트된 데이터 프레임.
    """
    # 제외할 타입 정의
    exclude_types = ['NOTE', 'Bullet','No Type','NOTE:n']
    point_five_exclude_types=['NOTE','NOTE:n']
    section_type=['CoverPage','RevisionSummary','TOC']
    
     # 초기 검증 컨텐츠 및 결과 설정
    data['Verification Content']='Normal'
    data['Verification Result']=True
    data['Numbering Type Change']=data['Numbering Type']
    
    # 각 레벨에 대한 검증 수행
    for level in level_bowl:
        level_dataset=data[data['Level']==level]
        
        # 정수 레벨과 소수 레벨에 대한 처리
        if  level == int(level):
            filtered_data=level_dataset[~level_dataset['Numbering Type'].isin(exclude_types)]
            mode_series=filtered_data['Numbering Type'].mode()
        else:
            filtered_data=level_dataset[~level_dataset['Numbering Type'].isin(point_five_exclude_types)]
            mode_series=filtered_data['Numbering Type'].mode()
            
         # 모드 값에 따른 검증 결과 업데이트
        if not mode_series.empty:
            level_mode = mode_series.iloc[0]

            for index,row in filtered_data.iterrows():
                if level == int(level):
                    if row['Numbering Type'] not in exclude_types and row['Numbering Type']!= level_mode and row['SECTION'] not in section_type:
                        data.loc[index,'Verification Content']="Level and numbering type don't match."
                        data.loc[index,'Verification Result']=False
                        data.loc[index,'Numbering Type Change']=level_mode
                else:
                    if row['Numbering Type'] not in point_five_exclude_types and row['Numbering Type']not in level_mode and row['SECTION'] not in section_type:
                        data.loc[index,'Verification Content']="Level and numbering type don't match."
                        data.loc[index,'Verification Result']=False   
                        data.loc[index,'Numbering Type Change']=[level_mode+','+'NOTE'+','+'NOTE:n']
        else:
            continue
    # 검증 타입 설정       
    data['Verification Type']='Verifying the Level And Numbering'
    
    return data 


# In[11]:


def check_level_and_type(data):
    """
    'Level'과 'Numbering Type'의 일관성을 검증하고, 검증 결과가 False인 행들을 선택합니다.

    Args:
    df_paragraph (DataFrame): 검증할 데이터 프레임.

    Returns:
    DataFrame: 검증 결과가 False인 행들을 포함하는 데이터 프레임.
    
    """
    # 검증할 레벨 범위 설정
    level_bowl=[1,1.5,2,2.5,3,3.5,4,4.5,5,5.5,6]
    # 레벨과 타입의 일관성 검증 수행
    combined_df=verification_level_numberingtype(data,level_bowl)
    # 검증 결과가 False인 행들 선택
    verification_dataset=combined_df[combined_df['Verification Result']==False]
    
    verification_dataset=verification_dataset[['ID', 'Text', 'Style', 'Numbering', 'Numbering Type', 'Level',
       'SECTION', 'Parent Index', 'Start', 'End',
       'Verification Content', 'Verification Result', 'Numbering Type Change',
       'Verification Type','File ID']]
    
    return combined_df,verification_dataset


# In[12]:


# Golas: Format verification paragraph full table

def level_sendout_format_verfification_paragraph(data,ord_dataset):
    """
    주어진 데이터 프레임과 주문 검증 데이터 프레임을 병합하여 포맷 검증 단락 전체 테이블을 생성하고, 
    이를 CSV 파일로 저장합니다.

    Args:
    data (DataFrame): 원본 데이터 프레임.
    ord_dataset (DataFrame): 레벨 검증 결과가 포함된 데이터 프레임.

    Returns:
    DataFrame: 레벨 검증 결과가 병합된 데이터 프레임.
    """
    # 레벨 검증 결과 데이터 프레임의 열 선택 및 이름 변경
    ord_dataset=ord_dataset[['ID','Verification Result']]
    ord_dataset.columns=['ID','Level Verification Result']
    
    # 원본 데이터 프레임과 레벨 검증 결과 데이터 프레임 병합
    paragraph_full_table=pd.merge(data,ord_dataset,on='ID',how='left')
    
    # 레벨 검증 결과가 없는 경우 True로 채우기
    paragraph_full_table['Level Verification Result']=paragraph_full_table['Level Verification Result'].fillna(True)
    
    # 필요한 열만 선택
    paragraph_full_table=paragraph_full_table[['Unique Id', 'ID', 'Type', 'Text', 'Style', 'Numbering',
       'Numbering Type', 'Level', 'SECTION', 'Parent Index', 'pos_tags',
       'parse_tree', 'ner', 'Indentation', 'Paragraph Type', 'ActionVerb',
       'TargetObject', 'Page', 'Start', 'End', 'Order Verification Result','Level Verification Result','File ID']]
    
    return paragraph_full_table


# In[13]:


def save_dataframe(data,docx_file,name):
    """
    지정된 경로에 데이터 프레임을 CSV 파일로 저장합니다.

    Args:
    docx_file (str): 원본 문서 파일의 경로.
    dataframes_dir (str): 데이터 프레임을 저장할 디렉토리의 경로.
    name (str): 저장할 파일의 이름에 추가할 문자열.

    Note:
    이 함수는 'df'라는 이름의 데이터 프레임 변수를 전역으로 가정합니다.
    """
    file_name = docx_file.split('\\')[-1]
    path = f"C:\\Users\\pc021\Desktop\\{name}\\{file_name}_{name}.csv"
    data.to_csv(path, encoding='utf-8-sig',index=False)

