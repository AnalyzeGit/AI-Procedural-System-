#!/usr/bin/env python
# coding: utf-8

# In[2]:


from docx import Document
from ExtractParagraph import *
from ExtractRun import *
from DocClientMapping import map_sentences_levenshtein
import re
import win32com.client
import docx
import pandas as pd


# In[ ]:


def process_doc_paragraph(docx_file):
    """
    이 함수는 주어진 Word 문서(.docx 파일)에서 단락을 추출하고, 단락의 정렬 및 null 값들을 정제합니다.
    
    매개변수:
    - docx_file: 분석할 Word 문서의 경로 또는 파일 객체입니다.
    
    반환값:
    - DataFrame: 정제된 단락 데이터를 포함하는 pandas DataFrame을 반환합니다.
    """
    paragraph = extract_paragraph(docx_file)
    paragraph = refine_alignment(paragraph)
    return refine_null_values(paragraph)

def process_client_paragraph(docx_file):
    """
    이 함수는 클라이언트의 Word 문서(.docx_file)에서 단락을 추출하고, 원하지 않는 행을 제거하며 특정 문자를 대체합니다.
    
    매개변수:
    - docx_file: 분석할 Word 문서의 경로 또는 파일 객체입니다.
    
    반환값:
    - DataFrame: 처리된 단락 데이터를 포함하는 pandas DataFrame을 반환합니다.
    """
    paragraph = extract_client_paragraph(docx_file)
    paragraph = remove_unwanted_rows(paragraph)
    paragraph['Paragraph'] = paragraph['Paragraph'].str.replace('', '-')
    return paragraph

def filter_and_rename(paragraph_data):
    """
    이 함수는 주어진 DataFrame에서 빈 단락을 제거하고, 컬럼명을 변경합니다.
    
    매개변수:
    - paragraph_data: 처리할 단락 데이터를 포함하는 pandas DataFrame 객체입니다.
    
    반환값:
    - DataFrame: 정제된 단락 데이터를 포함하는 pandas DataFrame을 반환합니다.
    """
    drop_idx = paragraph_data[paragraph_data['Paragraph'] == '\n'].index
    cleaned_data = paragraph_data.drop(drop_idx)
    return cleaned_data.rename({'Paragraph Numbering': 'Paragraph Numbering Text'}, axis=1)

def get_num_of_run(run_data):
    """
    이 함수는 각 단락에 포함된 run(문단 내의 텍스트 블록)의 수를 계산합니다.
    
    매개변수:
    - run_data: run 데이터를 포함하는 pandas DataFrame 객체입니다.
    
    반환값:
    - DataFrame: 각 단락의 run 수를 포함하는 pandas DataFrame을 반환합니다.
    """
    num_of_run = run_data.groupby('Paragraph Id').count()[['Run']].reset_index()
    num_of_run.columns = ['Paragraph Id', 'Num of Run']
    return num_of_run

def replace_bullet_point(df, column_name):
    """
    이 함수는 DataFrame 내의 특정 컬럼에서 불릿 포인트 기호를 'Bullet Point' 텍스트로 대체합니다.
    
    매개변수:
    - df: 처리할 DataFrame 객체입니다.
    - column_name: 불릿 포인트를 대체할 컬럼명입니다.
    
    반환값:
    - 반환 값은 없으며, DataFrame 내부의 데이터가 직접 변경됩니다.
    """
    rectangle_indices = df[df[column_name] == ''].index
    df.loc[rectangle_indices, column_name] = '125A0'
    
    stick_indices = df[df[column_name] == '-'].index
    df.loc[stick_indices, column_name] = 'B23B0'


# In[ ]:


def create_final_paragraph(docx_file):
    """
    이 함수는 Word 문서(.docx 파일)를 처리하여 최종 단락 데이터, 실행 데이터 및 위치 데이터를 생성합니다.
    여러 처리 단계를 통합하며, 진행 상황을 표시하는 진행률 표시줄을 포함합니다.
    
    매개변수:
    - docx_file: 처리할 Word 문서의 경로 또는 파일 객체입니다.
    
    반환값:
    - Tuple(DataFrame, DataFrame, DataFrame): 최종 단락 데이터, 실행 데이터, 위치 데이터를 포함하는 pandas DataFrame들을 반환합니다.
    """
        
  
    doc_paragraph = process_doc_paragraph(docx_file)
    client_paragraph = process_client_paragraph(docx_file)

    mapped_paragraph = map_sentences_levenshtein(doc_paragraph, client_paragraph)
    filtered_paragraph = filter_and_rename(mapped_paragraph)
   
    
    # Select necessary columns
    paragraph_data = filtered_paragraph[[
            'Paragraph Id', 'Paragraph', 'Paragraph Numbering Text', 
            'Paragraph Style', 'Paragraph Alignment',  
            'Paragraph Font Name', 'Paragraph Font Size', 'Numbering Type']]
    
    location_data = filtered_paragraph[[
            'Paragraph Id', 'Paragraph', 'Paragraph Style Client','Numbering Type', 'Paragraph Level','Left Indent',
            'Paragraph Numbering Text', 'Paragraph Page', 'Start', 'End']]

    run_data = extract_run(docx_file)
    refined_run_data = extract_run_case(run_data)
    
    num_of_run = get_num_of_run(refined_run_data)
    final_paragraph_data = pd.merge(paragraph_data, num_of_run, on='Paragraph Id', how='left')

    return final_paragraph_data, refined_run_data, location_data


# In[ ]:


def save_dataframe(data,file_name,name):
    """
    지정된 경로에 데이터 프레임을 CSV 파일로 저장합니다.

    Args:
    docx_file (str): 원본 문서 파일의 경로.
    dataframes_dir (str): 데이터 프레임을 저장할 디렉토리의 경로.
    name (str): 저장할 파일의 이름에 추가할 문자열.

    Note:
    이 함수는 'df'라는 이름의 데이터 프레임 변수를 전역으로 가정합니다.
    """
    path = f"{file_name}\\{name}.csv"
    data.to_csv(path, encoding='utf-8-sig',index=False)

