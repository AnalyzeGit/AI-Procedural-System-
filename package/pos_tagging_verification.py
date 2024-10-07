#!/usr/bin/env python
# coding: utf-8

# In[60]:


import pandas as pd
import re
import win32com.client
import docx
import pandas as pd
import nltk
import numpy as np

pd.set_option('display.max_rows',2400)
pd.set_option('display.max_colwidth',None)


# In[4]:


# 중복된 POS 태그가 있는 Token을 가진 행만 필터링하는 함수
def find_tokens_with_duplicate_pos(df):
    return df[df.groupby('Token')['POS'].transform('nunique') > 1]['Token'].unique()

# 중복된 POS 태그를 가진 Token에 대해 해당하는 행만 필터링하는 함수
def filter_rows_with_duplicate_pos(df):
    tokens_with_duplicate_pos = find_tokens_with_duplicate_pos(df)
    return df[df['Token'].isin(tokens_with_duplicate_pos)]

# 중복된 POS 태그가 있는 Token을 가진 대표 행만 선택하는 함수
def select_representative_rows(df):
    # Token과 POS에 대해 그룹화하고, 각 그룹의 첫 번째 행을 선택
    representative_rows = df.drop_duplicates(subset=['Token', 'POS'])
    return representative_rows

def extract_pos_tag_dataset(data):
    """
    데이터에서 'Text' 열이 '.'이 아닌 행들을 필터링하고, 
    'Token'과 'Paragraph Id'에 따라 정렬한 후, 검증 결과를 False로 설정합니다.

    Args:
    data (DataFrame): 처리할 데이터 프레임.

    Returns:
    DataFrame: 필터링되고 정렬된 데이터 프레임.
    """
    duplicate_pos_dataset=filter_rows_with_duplicate_pos(data)
    duplicate_pos_dataset=select_representative_rows(duplicate_pos_dataset)
    
    # 'Text' 열이 '.'이 아닌 행들을 필터링
    duplicate_pos_dataset=duplicate_pos_dataset[duplicate_pos_dataset['Text']!='.']
    
    # 'Token'과 'Paragraph Id'에 따라 데이터 정렬
    duplicate_pos_dataset=duplicate_pos_dataset.sort_values(by=['Token','Paragraph Id'])
    
    # 검증 결과 열 추가 및 False로 설정
    duplicate_pos_dataset['Verification Result']=False
    
    # 변수 재 정렬
    duplicate_pos_dataset=duplicate_pos_dataset[['Nlp Id','Paragraph Id','Lemma','POS','Dependency','Dependency Head','NER','Kind','Letter Case','Text','Verification Result','File ID']]
    
    return duplicate_pos_dataset


# In[7]:


# Golas: Format verification paragraph full table

def pos_sendout_format_verfification_paragraph(data,pos_dataset):
    """
    주어진 데이터 프레임과 형태소 태그(Pos Tag) 검증 데이터 프레임을 병합하여 
    포맷 검증 단락 전체 테이블을 생성하고, 이를 CSV 파일로 저장합니다.

    Args:
    data (DataFrame): 원본 데이터 프레임.
    ord_dataset (DataFrame): 형태소 태그 검증 결과가 포함된 데이터 프레임.

    Returns:
    DataFrame: 형태소 태그 검증 결과가 병합된 데이터 프레임.
    """
    # 형태소 태그 검증 결과 데이터 프레임의 열 선택 및 이름 변경
    pos_dataset=pos_dataset[['Paragraph Id','Verification Result']]
    pos_dataset.columns=['ID','Pos Tag Verification Result']
    
    # 원본 데이터 프레임과 형태소 태그 검증 결과 데이터 프레임 병합
    paragraph_full_table=pd.merge(data,pos_dataset,on='ID',how='left')
    # 형태소 태그 검증 결과가 없는 경우 True로 채우기
    paragraph_full_table['Pos Tag Verification Result']=paragraph_full_table['Pos Tag Verification Result'].fillna(True)
    
    # 필요한 열만 선택
    paragraph_full_table=paragraph_full_table[['Unique Id', 'ID', 'Type', 'Text', 'Style', 'Numbering',
       'Numbering Type', 'Level', 'SECTION', 'Parent Index', 'pos_tags',
       'parse_tree', 'ner', 'Indentation', 'Paragraph Type', 'ActionVerb',
       'TargetObject', 'Page', 'Start', 'End', 'Order Verification Result','Level Verification Result','Pos Tag Verification Result','File ID']]
    
    return paragraph_full_table

