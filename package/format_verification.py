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
from OrderVerification import *
from LevelVerification import *
from PosTaggingVerification import *
from nlp import *
from saveDataframe import save_dataframe
from loadDatabase import load_database
from setFileID import set_file_id
from uploadProcedure import upload_procedure


# In[2]:


def make_nlp(docx_file,file_id):
    # nlp 실행
    nlp = natural_language_processing (docx_file)
    
    # nlp 전처리
    pre_nlp=extract_nlp_case(nlp)
    
    # nlp 인덱스 재정렬
    pre_nlp=nlp_index_cleanup(pre_nlp)
    
    # nlp file_id 추가
    pre_nlp=set_file_id(pre_nlp,file_id)
    
    return pre_nlp


# In[3]:


# Goals: 형식 검증 자동화 함수 구현

def type_verification_automation_algorithm(data,nlp):
    
    # Order 형식 검증을 위한 데이터 정제 
    refinind_dataset=pre_processing(data)
    
    # 형식 검증 
    sequential_validated_dataset=sequence_verification_algorithm(refinind_dataset)
    
    # 순서가 맞지 않는 데이터 프레임 추출 
    order_false_dataset=sendout_order_verification(sequential_validated_dataset)
    
    # 다음 검증을 위한 데이터 프레임  
    paragraph_full_table=sendout_format_verfification_paragraph(data,order_false_dataset)
    
    # 레벨과 타입이 맞지 않는 데이터 프레임 추출 
    level_validated_dataset,level_false_dataset=check_level_and_type(paragraph_full_table)
    
    # 일관성 검증 후 다음 검증을 위한 데이터 프레임
    paragraph_full_table=level_sendout_format_verfification_paragraph(level_validated_dataset,level_false_dataset)
    
    # nlp 중복 검증 데이터프레임
    nlp_false_dataset=extract_pos_tag_dataset(nlp)
    paragraph_full_table=pos_sendout_format_verfification_paragraph(paragraph_full_table,nlp_false_dataset)
    
    return paragraph_full_table, order_false_dataset,level_false_dataset,nlp_false_dataset


# In[21]:


def apply_format_algorithm(docx_file,paragraph_full_table,file_path,file_id):
    
    # nlp
    nlp_verification_dataset=make_nlp(docx_file,file_id)
    paragraph_full_table,order_false_dataset,level_false_dataset,nlp_false_dataset=type_verification_automation_algorithm(paragraph_full_table,nlp_verification_dataset)
    
    # 데이터 내보내기  
    save_dataframe(paragraph_full_table,file_path,'format_verification')
    save_dataframe(order_false_dataset,file_path,'rank')
    save_dataframe(level_false_dataset,file_path,'level')
    save_dataframe(nlp_false_dataset,file_path,'nlp_postagg')
    
    # 데이터 로드
    upload_format_dataset=upload_procedure(file_path,'format_verification')
    upload_order_dataset=upload_procedure(file_path,'rank')
    upload_level_dataset=upload_procedure(file_path,'level')
    upload_nlp_dataset=upload_procedure(file_path,'nlp_postagg')
    
    
     # 데이터 적재  
    load_database(upload_format_dataset,'format verification test')
    load_database(upload_order_dataset,'rank test')
    load_database(upload_level_dataset,'consistency test')
    load_database(upload_nlp_dataset,'nlp_test')

