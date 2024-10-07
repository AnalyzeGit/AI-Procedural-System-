#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from docx import Document
from difflib import SequenceMatcher
import re
import win32com.client
import docx
import pandas as pd
import string
import Levenshtein
import glob


# In[ ]:


def calculate_similarity(str1, str2):
    """ 문장 유사도 계산 기능
    
    매개변수:
    - str1: 기준 문장 
    - str2: 비교할 문장
    
    반환값:
    - 문서 유사도
    """
    return 1 - Levenshtein.distance(str1, str2) / max(len(str1), len(str2))


def map_sentences_levenshtein(Paragraph, Paragraph_test):
    """ 기준 문장과 비교 문장 유사도를 계산한 문장 매핑 기능 
    
    매개변수:
    - Paragraph: 기준 문장
    - Paragraph_test: 비교 문장
    
    반환 값:
    - 높은 유사도의 문장이 매핑된 DataFrames
    """
    data = {
        'Paragraph Id':[],
        'Paragraph': [],
        'Paragraph Test': [],
        'Paragraph Style':[],
        'Paragraph Style Client':[],
        'Paragraph Alignment':[],
        'Paragraph Font Name':[],
        'Paragraph Font Size':[],
        'Paragraph Numbering': [],
        'Paragraph Level':[],
        'Left Indent':[],
        'Numbering Type': [],
        'Start':[],
        'End':[],
        'Paragraph Page': [],
        'score': [],
        'Test Paragraph Index': []   # 인덱스 값을 저장하기 위한 새로운 컬럼
    }
     
    comparative_idx=[]
    page=0
    
    for idx, val in Paragraph.iterrows():
        Doc_Paragraph = val['Paragraph']
        comparative_Paragraph=[]
        similarity_scores = {}
        idx_scores = {}  # 각 문장의 인덱스 값을 저장하기 위한 딕셔너리
        
        for test_idx, test_val in Paragraph_test.iterrows():
            # 매칭된 파라그래프, 중복된 파라그래프 조건
            if (test_idx not in comparative_idx) & (test_val['Paragraph'] not in comparative_Paragraph): 
                Client_Paragraph = test_val['Paragraph']
                similarity_scores[Client_Paragraph] = calculate_similarity(Doc_Paragraph, Client_Paragraph)
                idx_scores[Client_Paragraph] = test_idx   # 인덱스 값을 저장
                comparative_Paragraph.append(Client_Paragraph)
            else:
                pass
        # 동점 상황에서 마지막 키를 선택, 딕셔너리가 존재할 때
        if len(similarity_scores)>0:
            best_match = max(similarity_scores, key=similarity_scores.get)
            best_score = similarity_scores[best_match]  # 가장 유사한 문장의 인덱스를 가져옴
            best_match_idx = idx_scores[best_match]
            comparative_idx.append(best_match_idx)
            data['Paragraph Id'].append(Paragraph.loc[idx,'Paragraph Id'])
            data['Paragraph'].append(Doc_Paragraph)
            data['Paragraph Test'].append(best_match)
            data['score'].append(best_score)
            data['Test Paragraph Index'].append(best_match_idx)
            data['Paragraph Style'].append(Paragraph.loc[idx,'Paragraph Style'])
            data['Paragraph Style Client'].append(Paragraph_test.loc[best_match_idx,'Paragraph Style Client'])
            data['Paragraph Alignment'].append(Paragraph.loc[idx,'Paragraph Alignment'])
            data['Paragraph Font Name'].append(Paragraph.loc[idx,'Paragraph Font Name'])
            data['Paragraph Font Size'].append(Paragraph.loc[idx,'Paragraph Font Size'])
            data['Paragraph Numbering'].append(Paragraph_test.loc[best_match_idx,'Numbering'])
            data['Numbering Type'].append(Paragraph_test.loc[best_match_idx,'Numbering Type'])
            data['Start'].append(Paragraph_test.loc[best_match_idx,'Start'])
            data['End'].append(Paragraph_test.loc[best_match_idx,'End'])
            data['Paragraph Page'].append(Paragraph_test.loc[best_match_idx,'Page Number'])
            data['Paragraph Level'].append(Paragraph_test.loc[best_match_idx,'Paragraph Level'])
            data['Left Indent'].append(Paragraph_test.loc[best_match_idx,'Left Indent'])
            page+=Paragraph_test.loc[best_match_idx,'Page Number']
            
        #딕셔너리가 존재하지 않을때
        else:
            data['Paragraph Id'].append(Paragraph.loc[idx,'Paragraph Id'])
            data['Paragraph'].append(Doc_Paragraph)
            data['Paragraph Test'].append('None')
            data['score'].append('None')
            data['Test Paragraph Index'].append('None')
            data['Paragraph Style'].append(Paragraph.loc[idx,'Paragraph Style'])
            data['Paragraph Style Client'].append('None')
            data['Paragraph Alignment'].append(Paragraph.loc[idx,'Paragraph Alignment'])
            data['Paragraph Font Name'].append(Paragraph.loc[idx,'Paragraph Font Name'])
            data['Paragraph Font Size'].append(Paragraph.loc[idx,'Paragraph Font Size']) 
            data['Paragraph Numbering'].append('None')
            data['Numbering Type'].append('None')
            data['Paragraph Page'].append('None')
            data['Paragraph Numbering'].append('None')
            data['Numbering Type'].append('None')
            data['Paragraph Page'].append('None')
            data['Paragraph Level'].append('None')
            data['Left Indent'].append('None')
        # 여기에 다른 데이터 (Numbering, Page 등)를 추가하는 코드를 넣을 수 있습니다.
        #Numbering Type	,Paragraph Style,Page Number
    return pd.DataFrame(data)

