#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import pyodbc
from textAnalysis import process_text, combined_extraction

# Goals: location.csv 업로드 후 Paragraph 이용 정보 추출(ActionStep..etc) 

def build_paragraph_full_table(data):

    # Clean and process DataFrame(정보 추출 처리)
    data['Text'] = data['Text'].str.replace('\x0b', '')
    data[['pos_tags', 'parse_tree', 'ner']] = data['Text'].apply(process_text).apply(pd.Series)
    data[['Paragraph Type','ActionVerb', 'TargetObject']] = data['Text'].apply(combined_extraction).apply(pd.Series)

    #reset_index 
    data.reset_index(inplace=True)

    # columns 정리
    data=data[['Unique Id','ID', 'Type', 'Text', 'Style', 'Numbering', 'Numbering Type',
       'Level', 'SECTION', 'Parent Index', 'pos_tags', 'parse_tree', 'ner',
       'Indentation','Paragraph Type', 'ActionVerb', 'TargetObject','Page','Start','End','File ID']]

    return data

