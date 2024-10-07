#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from docx import Document
from difflib import SequenceMatcher
from locationApply import *
import re
import win32com.client
import docx
import pandas as pd
import Levenshtein
import glob


# In[ ]:


def preprocess_table_result(table_result):
    """
    테이블 결과 데이터를 전처리합니다. 'Numbering' 열을 추가하고, 'Type'을 'Table'로 설정합니다.
    필요한 열만 선택하고, 'Style' 열의 null 값을 'Normal'로 채웁니다.
    
    매개변수:
    - table_result: 전처리할 테이블 결과 데이터를 포함하는 DataFrame입니다.
    
    반환값:
    - DataFrame: 전처리된 테이블 결과 데이터를 포함하는 DataFrame을 반환합니다.
    """
    table_result['Numbering'] = 'None'
    table_result['Type'] = 'Table'
    table_result = table_result[['ID', 'Text', 'Style', 'Numbering', 'Type', 'Page', 'Start', 'End']]
    table_result['Style'] = table_result['Style'].fillna('Normal')
    return table_result

def preprocess_location(location_df):
    """
    위치 데이터를 전처리합니다. 'Type'을 'Paragraph'로 설정하고, 필요한 열만 선택한 후, 열 이름을 변경합니다.
    
    매개변수:
    - location_df: 전처리할 위치 데이터를 포함하는 DataFrame입니다.
    
    반환값:
    - DataFrame: 전처리된 위치 데이터를 포함하는 DataFrame을 반환합니다.
    """
    location_df['Type'] = 'Paragraph'
    location_df = location_df[['Paragraph Id', 'Paragraph', 'Paragraph Style Client', 'Paragraph Numbering Text', 'Type', 'Total Level',
                               'Numbering Type Code','SECTION','Parent Index','Paragraph Page', 'Start', 'End']]
    location_df.columns = ['ID', 'Text', 'Style', 'Numbering', 'Type', 'Level','Numbering Type','SECTION',
                           'Parent Index','Page', 'Start', 'End']
    return location_df

def remove_empty_lines(df):
    """
    DataFrame에서 빈 줄('\n')을 포함하는 행을 제거합니다.
    
    매개변수:
    - df: 처리할 DataFrame 객체입니다.
    
    반환값:
    - DataFrame: 빈 줄이 제거된 DataFrame을 반환합니다.
    """
    drop_index = df[df['Text'] == '\n'].index
    df = df.drop(index=drop_index).reset_index()
    return df

def put_in_order(df):
    """
    DataFrame의 열 순서를 재정렬합니다.
    
    매개변수:
    - df: 열 순서를 재정렬할 DataFrame 객체입니다.
    
    반환값:
    - DataFrame: 열 순서가 재정렬된 DataFrame을 반환합니다.
    """
    df=df[['ID','Type','Text','Style','Numbering','Numbering Type','Level','SECTION','Parent Index','Indentation','Page','Start','End']]
    return df

def build_location_dataset(docx_file, Location):
    """
    주어진 Word 문서와 위치 데이터를 사용하여 최종 위치 DataFrame을 생성합니다.
    이 과정에는 테이블 텍스트 추출, 테이블 위치 추출, 문장 매핑, 전처리, 정렬 및 빈 줄 제거가 포함됩니다.
    
    매개변수:
    - docx_file: 처리할 Word 문서의 경로입니다.
    - Location: 위치 데이터를 포함하는 DataFrame 객체입니다.
    
    반환값:
    - DataFrame: 최종 위치 데이터를 포함하는 DataFrame을 반환합니다.
    """
    # 주어진 docx 파일에서 테이블의 텍스트를 추출
    doc_table = extract_table_text(docx_file)
    
    # 테이블의 위치(예: 페이지 번호, 테이블의 위치 등)에 초점
    loc_table = extract_table_location(docx_file)
    
    # 추출된 텍스트와 위치 데이터를 매핑
    table_result = table_map_sentences_levenshtein(doc_table, loc_table)
    
    # 매핑된 결과를 전처리
    table_result = preprocess_table_result(table_result)
        
    # 위치 데이터를 전처리     
    location_df = preprocess_location(Location)
    
    # 위치 데이터와 테이블 결과를 병합
    concat_df = pd.concat([location_df, table_result])
    
    #병합된 데이터를 'Start' 컬럼 기준으로 정렬
    sorted_location = concat_df.sort_values(by=['Start'], ascending=True).reset_index(drop=True)
    
    # 정렬된 데이터에서 빈 줄을 제거
    final_location_df = remove_empty_lines(sorted_location)
    
    # 최종 위치 데이터 프레임에 들여쓰기를 추가
    create_indentation(final_location_df)
    
    # 데이터 프레임의 내용을 특정 순서로 재배치
    final_location_df=put_in_order(final_location_df)
    
    # 데이터 프레임의 ID를 재설정
    final_location_df=reset_id(final_location_df)
        
    return final_location_df

