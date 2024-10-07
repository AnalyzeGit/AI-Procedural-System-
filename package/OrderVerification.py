#!/usr/bin/env python
# coding: utf-8

# In[3]:


import pandas as pd
import re
import win32com.client
import docx
import pandas as pd
import nltk
import numpy as np

pd.set_option('display.max_rows',None)
pd.set_option('display.max_colwidth',None)


# <span style=' border:0.5px solid black; padding:5px; border-radius:5px;'>docx file 로드 </span> 

# In[53]:


# Goals: 검증 데이터 프레임에서 사용할 변수 및 Paragraph 선택

def pre_processing(data):
    """
    주어진 데이터를 필터링하고 관련 열만 선택하여 전처리합니다.

    매개변수:
    data (DataFrame): 전처리할 입력 데이터.

    반환값:
    DataFrame: 관련 열만 포함하는 전처리된 데이터.
    """
    # 'Type'이 'Paragraph'인 행을 필터링
    df_paragraph = data[data['Type'] == 'Paragraph']

    # 관련 열 선택
    columns = ['ID', 'Text', 'Style', 'Numbering', 'Numbering Type', 
               'Level', 'SECTION', 'Parent Index', 'Start', 'End', 'File ID']
    df_paragraph = df_paragraph[columns]

    return df_paragraph


# In[39]:


# Goals: 검증 데이터 프레임() 
def update_numbering(data, index, numbering):
    """
    주어진 데이터 프레임에서 특정 인덱스의 'Numbering' 값을 업데이트합니다.

    Args:
    data (DataFrame): 업데이트할 데이터 프레임.
    index (int or list): 'Numbering' 값을 업데이트할 행의 인덱스.
    numbering (any): 새로운 'Numbering' 값.

    Returns:
    DataFrame: 업데이트된 데이터 프레임.
    """
    data.at[index, 'Numbering'] = numbering
    return data


# <span style=' border:0.5px solid black; padding:5px; border-radius:5px;'>알고리즘 개발 </span> 

# In[32]:


def extract_part(numbering, numbering_type):
    
    """
    주어진 번호 매기기(numbering) 값에서 특정 부분을 추출합니다.

    Args:
    numbering (str): 분석할 번호 매기기 문자열.
    numbering_type (str): 번호 매기기의 타입 (예: 'Bullet', 'NOTE').

    Returns:
    str or NaN: 추출된 부분 또는 NaN (적합한 부분이 없는 경우).
    """
        
    # numbering이 NaN인 경우
    if pd.isna(numbering):
        return np.nan
    
    # numbering_type이 'Bullet' 또는 'NOTE'로 시작하는 경우
    if (numbering_type == 'Bullet') or (numbering_type.startswith('NOTE')):
        return np.nan

    # '.'을 기준으로 문자열 분리 
    parts = numbering.split('.')

    # 분리된 결과가 있는 경우
    if parts:
        pattern = r'\b[a-zA-Z]\.(?!\d)|\b\d\.\s|\b\d\)\s?'
        matche=re.findall(pattern, numbering)
        # numbering이 오직 숫자 또는 오직 영어로만 구성되어 있거나, 영어와 숫자를 모두 포함하는 경우
        if matche:
            return parts[0]  # 마지막 부분 반환
        else:
            return parts[-1]  # 그 외의 경우 첫 번째 부분 반환 
    else:
        return numbering  # 분리할 수 없는 경우 원본 문자열 반환


# In[34]:


def check_sequence_and_start_order(data):
    
    """
    주어진 번호 매기기(numbering Sequence) 값을 이용하여 형식검증을 진행합니다.

    Args:
    data (DataFrame): 업데이트할 데이터 프레임.

    Returns:
    DataFrame: 업데이트된 데이터 프레임.
    """
        
        
    # 'None' 값이 있는 행 제거
    data = data.dropna(subset=['Numbering Sequence'])
    data['Verification Content']='Normal'
    data['Verification Result']=True
    
    # 데이터가 비어 있는 경우를 확인
    if data.empty:
        return True, "None Type pass", []

    def is_numeric_sequence_ordered(sequence):
        wrong_indices = []
        previous_value = None
        sequence['Numbering Sequence']=sequence['Numbering Sequence'].astype(int)
        
        for index, row in sequence.iterrows():
            if previous_value is not None and row['Numbering Sequence'] != previous_value + 1:
                data.loc[index,'Verification Content']="Numeric sequence is not continuously increasing."
                data.loc[index,'Verification Result']=False
                
            previous_value = row['Numbering Sequence']
                

    def is_alphabetic_sequence_ordered(sequence):
        wrong_indices = []
        previous_value = None
        for index, row in sequence.iterrows():
            if previous_value is not None and ord(row['Numbering Sequence']) != ord(previous_value) + 1:
                data.loc[index,'Verification Content']= "Alphabetic sequence is not in correct order. "
                data.loc[index,'Verification Result']=False
            previous_value = row['Numbering Sequence']

    data['Numbering Sequence'] = data['Numbering Sequence'].astype(str)

    numeric_part = data['Numbering Sequence'].str.extract('(\d+)')[0].dropna().astype(int)
    alphabetic_part = data['Numbering Sequence'].str.extract('([A-Za-z]+)')[0].dropna()

    if not numeric_part.empty:
        # 'Numbering Sequence' 열에서 숫자가 아닌 모든 문자를 제거합니다.
        data['Numbering Sequence'] = data['Numbering Sequence'].str.replace(r'\D', '', regex=True)
        is_numeric_sequence_ordered(data)

    if not alphabetic_part.empty:
        is_alphabetic_sequence_ordered(data)
    
    data['Validation Type']='Verifying the Numbering Order'
    
    return data


# In[35]:


# Goals: 영어 시퀀스 숫자로 처리

def convert_alpha_to_number(data):
    """
    데이터 프레임의 'Numbering Sequence' 열에서 영어 알파벳을 숫자로 변환합니다.

    Args:
    data (DataFrame): 변환을 수행할 데이터 프레임.

    Returns:
    DataFrame: 변환된 데이터 프레임.
    """
        
    for index,row in data.iterrows():
        sequence=row['Numbering Sequence']
         # 'Numbering Sequence'가 문자열이고 알파벳인 경우
        if isinstance(sequence, str) and sequence.isalpha():  # 알파벳인 경우에만 변환
             # 알파벳을 대문자로 변환하고 ASCII 값에서 64를 빼서 숫자로 변환
            number = ord(sequence.upper()) - 64
            data.loc[index,'Numbering Sequence']=number
        else:  # 알파벳이 아닌 경우, 원래의 문자(또는 숫자)를 그대로 추가
            pass
    return data


# In[47]:


def organize_dataframes(data):
    """
    데이터 프레임의 열을 재구성하고 열 이름을 변경합니다.

    Args:
    data (DataFrame): 재구성할 데이터 프레임.

    Returns:
    DataFrame: 열이 재구성되고 이름이 변경된 데이터 프레임.
    """
        
    # 열 이름 변경
    data=data[['ID','Text','Style','Numbering','Numbering Type','Level','SECTION','Parent Index','Start','End','Numbering Sequence','Validation Type','Verification Result','Verification Content','File ID']]
    
    # 열 이름 변경
    data.columns=['Paragraph Id','Paragraph','Style','Numbering','Numbering Type','Level','SECTION','Parent Index','Start','End','Numbering Sequence','Verification Type','Verification Result','Verification Content','File Id']
    
    return data

def sequence_verification_algorithm(data):
    """
    주어진 데이터에서 번호 매기기 시퀀스를 검증하고, 데이터를 조직화하는 알고리즘을 수행합니다.

    Args:
    data (DataFrame): 검증 및 조직화할 데이터 프레임.

    Returns:
    DataFrame: 검증된 번호 매기기 시퀀스와 조직화된 데이터를 포함한 데이터 프레임.
    """
        
    # 데이터프레임을 담을 리스트
    dataframes = []
    
   # 부모 인덱스별로 데이터프레임 추출 및 검증
    parent_indices=list(data['Parent Index'].unique())
    parent_indices.remove('highest level')
    
    # 같은 부모의 데이터프레임 추출 후 순서 검증
    for unique_parent in parent_indices:
         # 특정 부모 인덱스를 가진 데이터프레임 추출
        child_dataset=data[data['Parent Index']==unique_parent]
        # 번호 매기기 시퀀스 추출
        child_dataset['Numbering Sequence'] =child_dataset.apply(lambda x: extract_part(x['Numbering'], x['Numbering Type']), axis=1)
        # 순서 검증
        child_dataset=check_sequence_and_start_order(child_dataset)
       
        # 검증된 데이터프레임을 리스트에 추가
        if isinstance(child_dataset, pd.DataFrame):
            dataframes.append(child_dataset)
        else:
            print(f"반복 {unique_parent}에서 child_dataset은 DataFrame이 아닙니다: {type(child_dataset)}")
    
    # 모든 데이터프레임을 하나로 합치기
    combined_df=pd.concat(dataframes, ignore_index=True)
    # 알파벳 번호를 숫자로 변환
    combined_df=convert_alpha_to_number(combined_df)
    # 데이터프레임 재구성
    combined_df=organize_dataframes(combined_df)
    
    return combined_df


# In[50]:


# Goals: Order Verification 추출

def sendout_order_verification(data):
    """
    'Verification Result'가 False인 데이터를 기반으로 주문 검증 데이터를 추출하고,
    이를 CSV 파일로 저장합니다.

    Args:
    data (DataFrame): 검증할 데이터 프레임.

    Returns:
    DataFrame: 주문 검증 데이터를 포함한 데이터 프레임.
    """
    # 'Verification Result'가 False인 데이터 추출    
    order_verification_false=data[data['Verification Result']==False]
    # 해당 데이터의 부모 인덱스 추출
    parents=list(order_verification_false['Parent Index'].unique())
    # 해당 부모 인덱스를 가진 모든 데이터 추출
    order_verification=data[data['Parent Index'].isin(parents)]
    
    return order_verification


# In[40]:


# Golas: Format verification paragraph full tableObjecObjec
def sendout_format_verfification_paragraph(data,ord_dataset):
    """
    주어진 데이터에 순서 검증 결과를 병합하고 포맷 검증 단락 전체 테이블을 CSV 파일로 저장합니다.

    Args:
    data (DataFrame): 원본 데이터 프레임.
    ord_dataset (DataFrame): 주문 검증 결과가 포함된 데이터 프레임.
    name (str): 저장될 파일의 이름을 지정하는 문자열.

    Returns:
    DataFrame: 주문 검증 결과가 병합된 데이터 프레임.
    """
     # 주문 검증 결과 데이터 프레임의 열 선택 및 이름 변경
    ord_dataset=ord_dataset[['Paragraph Id','Verification Result']]
    ord_dataset.columns=['ID','Order Verification Result']
    
    # 원본 데이터 프레임과 주문 검증 결과 데이터 프레임 병합
    paragraph_full_table=pd.merge(data,ord_dataset,on='ID',how='left')
    # 주문 검증 결과가 없는 경우 'No type verification'으로 채우기
    paragraph_full_table['Order Verification Result']=paragraph_full_table['Order Verification Result'].fillna('True')
    # 필요한 열만 선택
    paragraph_full_table=paragraph_full_table[['Unique Id', 'ID', 'Type', 'Text', 'Style', 'Numbering',
       'Numbering Type', 'Level', 'SECTION', 'Parent Index', 'pos_tags',
       'parse_tree', 'ner', 'Indentation', 'Paragraph Type', 'ActionVerb',
       'TargetObject', 'Page', 'Start', 'End', 'Order Verification Result','File ID']]
    
    return paragraph_full_table


# In[ ]:


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

