#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from docx import Document
import re
import win32com.client
import docx
import pandas as pd


# In[ ]:


def get_run_style(run):
    """ 문단 구조에서 Run 단어의 속성을 추출하는 기능
    
    매개변수:
    - Run: 문장에 속한 Run
    
    반환 값:
    - Run에 속하는 속성 
    
    """
    style = {
        'Run Bold': run.bold,
        'Run Italic': run.italic,
        'Run Underline': run.underline,
        'Run Style':run.style.name,
        'Run Size':run.font.size,}
    return style
    
def extract_run(docx_file):
    """문단 구조에서 Run 단어와 속성을 추출하여 데이터 프레임 변환하는 기능
    
    매개변수:
    - docx_file: 절차서 경로
    
    반환 값:
    - Run 구조 DataFrames
    
    """
    doc = docx.Document(docx_file)

    run_property_list=[]
    index = 0
    run_index=0

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if run.text!='':
                run_style = get_run_style(run)
                run_text = run.text

                run_property_list.append({
            'Run Id':run_index,
            'Paragraph Id': index,
            'Run': run_text,
                **run_style
                    })
            run_index+=1
        index += 1
        
    run_df = pd.DataFrame(run_property_list)
    extract_run_case(run_df)
    return run_df


# In[ ]:


def classify_text_character(text):
    
    """ 주어진 텍스트에서 개별 문자를 분류하는 기능
    
    매개변수:
    - text: 분류될 Run 텍스트
    
    반환 값:
    - Run Case를 포함한 Run DataFrames    
    """
    
    english_pattern = re.compile(r'^[A-Za-z]+$')
    special_characters_pattern = re.compile(r'[!@#$%^&*(),.?":{}|<>\-]')
    number_pattern = re.compile(r'^\d+(\.\d+)?$')
    alphanumeric_pattern = re.compile(r'^[a-zA-Z0-9]+$')
    space_pattern = re.compile(r'^\s+$')
    
    if re.match(special_characters_pattern, text):
        return 'Special Chars & Punct'
    elif text.islower():
        return 'Lower'
    elif text.isupper():
        return 'Upper'
    elif re.fullmatch(number_pattern, text):
        return 'Number'
    elif text == '\n':
        return 'newline'
    elif text == '\t':
        return 'tab'
    elif re.fullmatch(space_pattern, text):
        return 'blank space'
    else:
        return 'contain variety things'

def extract_run_case(run):
    
    """지정된 데이터 프레임에서 런 추출 및 분류 기능"""
    
    run['Run Case'] = [classify_text_character(text) for text in run['Run']]

    # Assuming 'Run' dataframe has NaNs, replacing them with 'None'
    run.fillna('None', inplace=True)

    selected_columns = [
        'Run Id','Paragraph Id', 'Run', 'Run Bold', 'Run Italic', 'Run Underline',
        'Run Style', 'Run Size', 'Run Case'
    ]
    
    return run[selected_columns]

