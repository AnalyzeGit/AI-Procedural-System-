#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from docx import Document
import re
import win32com.client
import docx
import pandas as pd


# In[ ]:


def extract_paragraph(docx_file):
    """
    이 함수는 주어진 Word 문서(.docx 파일)에서 각 단락의 텍스트와 해당 속성들을 추출하여 DataFrame으로 반환합니다.
    추출되는 속성에는 단락의 텍스트, 스타일, 정렬, 왼쪽 들여쓰기, 글꼴 이름 및 크기 등이 포함됩니다.
    
    매개변수:
    - docx_file: 분석할 Word 문서의 경로 또는 파일 객체입니다.
    
    반환값:
    - 추출된 단락과 속성 정보를 포함하는 pandas DataFrame을 반환합니다. 각 단락은 고유한 ID로 식별됩니다.
    
    함수 사용 예시:
    df = extract_paragraph('example.docx')
    
    """
    paragraph_properties_list=[]
    paragraph_id=0
    doc = docx.Document(docx_file)

    for paragraph in doc.paragraphs:
        paragraph_text = paragraph.text
        if paragraph_text != '':
            paragraph_properties_list.append({
                'Paragraph Id': paragraph_id,
                'Paragraph': paragraph_text,
                'Paragraph Style':paragraph.style.name,
                'Paragraph Alignment':paragraph.alignment,
                'Paragraph Font Name':paragraph.style.font.name,
                'Paragraph Font Size':paragraph.style.font.size
            })
            paragraph_id+=1
    
    return pd.DataFrame(paragraph_properties_list) 


# In[ ]:


def refine_alignment(data):
    """
    이 함수는 주어진 DataFrame 내의 'Paragraph Alignment' 컬럼의 숫자 값을 해당하는 텍스트 값으로 변환합니다.
    변환은 LEFT, CENTER, RIGHT, JUSTIFY, DISTRIBUTE에 해당하는 각각의 숫자 값을 대응하는 텍스트로 매핑하여 수행됩니다.
    
    매개변수:
    - data (DataFrame): 'Paragraph Alignment' 값을 포함하는 pandas DataFrame 객체입니다.
    
    반환값:
    - DataFrame: 'Paragraph Alignment' 컬럼의 숫자 값이 해당하는 텍스트 값으로 변환된 DataFrame을 반환합니다.
    
    """
    data['Paragraph Alignment'] = data['Paragraph Alignment'].astype(str)
    
    # Mapping from numerical values to alignment names
    alignment_mapping = {
        0: 'LEFT',
        1.0: 'CENTER',
        2.0: 'RIGHT',
        3.0: 'JUSTIFY',
        4.0: 'DISTRIBUTE'
    }
    
    # Refine the alignment values in the DataFrame
    for value, alignment in alignment_mapping.items():
        value_index = data[data['Paragraph Alignment'] == value].index
        data.loc[value_index, 'Paragraph Alignment'] = str(alignment)
        
    return data


# In[ ]:


def refine_null_values(data):
    """
    이 함수는 주어진 DataFrame 내의 특정 컬럼들에서 null 값들을 기본값으로 치환합니다.
    'Paragraph Alignment', 'Paragraph Left Indent', 'Paragraph Font Size' 등의 컬럼에 대해 설정된 기본값으로 null 값을 대체합니다.
    
    매개변수:
    - data (DataFrame): null 값을 치환할 pandas DataFrame 객체입니다. 이 DataFrame은 특정 컬럼들을 포함해야 합니다.
    
    반환값:
    - DataFrame: 지정된 컬럼의 null 값이 기본값으로 치환된 DataFrame을 반환합니다.
    
    """
    
    # Mapping from column names to default values
    default_values = {
        'Paragraph Alignment': 'LEFT',
        'Paragraph Font Size': 'None'
    }
    
    # Refine null values in the DataFrame
    for col, default_value in default_values.items():
        data[col] = data[col].fillna(default_value)
        
    return data


# In[ ]:


def extract_client_paragraph(document_path):
    
    """
    Word 문서에서 문단을 추출한 후 DataFrame으로 반환
    
    매개변수:
    - document_path(str): Word 문서의 파일 경로
        
    반환값:
    - DataFrame: 단락 정보를 포함하는 DataFrame.
    """
    win32com.client.gencache.EnsureDispatch('Word.Application')
    word_app = win32com.client.Dispatch('Word.Application')
    word_app.Visible = False  # Word 창을 표시하지 않음

    doc = word_app.Documents.Open(document_path)
    paragraph_properties_list = []

    for paragraph in doc.Paragraphs:
        paragraph_text = paragraph.Range.Text.strip()
        list_type=paragraph.Range.ListFormat.ListType
        style_name = paragraph.Style.NameLocal
        page_number = paragraph.Range.Information(win32com.client.constants.wdActiveEndPageNumber)
        paragraph_properties_list.append({
            'Paragraph': paragraph_text,
            'Page Number':page_number,   
            'Numbering Type':list_type,
            'Numbering':paragraph.Range.ListFormat.ListString,
            'Paragraph Level':paragraph.Range.ListFormat.ListLevelNumber,
            'Left Indent': paragraph.Format.LeftIndent,
            'Start':paragraph.Range.Start,
            'End':paragraph.Range.End,
            'Paragraph Style Client':style_name})
        
    # Close the document and quit Word application
    doc.Close()
    word_app.Quit()
    return pd.DataFrame(paragraph_properties_list)


# In[ ]:


def remove_unwanted_rows(data):
    """데이터프레임에서 특정 문자로 시작하거나 비어 있는 행을 제거합니다.
    
    매개변수:
        data (pd.데이터프레임): 정리할 데이터 프레임입니다.
        
    반환값:
        pd.DataFrame: 원치 않는 행이 제거된 새로운 DataFrame입니다.
    """
    
    # Patterns to remove
    patterns_to_remove = ['', ' ', '', '/']
    # extend list
    indices_to_remove = []
    
    #startswith remove
    indices_to_remove.extend(data[data['Paragraph'].str.startswith('*')].index.tolist())
    
    # Identify indices of rows to remove
    for pattern in patterns_to_remove:
        indices_to_remove.extend(data[data['Paragraph']==pattern].index.tolist())
    
    # Remove rows
    data.drop(indices_to_remove, inplace=True)
    
    return data

