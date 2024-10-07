#!/usr/bin/env python
# coding: utf-8

# In[3]:


from docx import Document
from DocClientMapping import map_sentences_levenshtein
from ParagraphRunLocation import *
from ExtractFinalLocation import build_final_location_structure
from buildParagraphFullTable import build_paragraph_full_table
from setFileID import set_file_id
from saveDataframe import save_dataframe
from loadDatabase import load_database
from uploadProcedure import upload_procedure
import re
import win32com.client
import docx
import pandas as pd
import numpy as np 


# In[11]:


docx_file=r"C:\Users\pc021\Desktop\AI 절차서 검증 시스템\DOCX 파일\docx_files\OP-66. R20.docx"


# In[12]:


def apply_algorithm(docx_file,file_path,file_name,file_id):
    
    # paragraph, run, location 데이터 추출
    paragraph,run,location=create_final_paragraph(docx_file)
    
    # location  버전 B 업데이트
    builded_location=build_final_location_structure(location,docx_file)
    
    # File ID 설정
    builded_location=set_file_id(builded_location,file_id)
    
    # location  버전 C 업데이트
    paragraph_full_table=build_paragraph_full_table(builded_location)
    
    # 데이터 내보내기  
    save_dataframe(paragraph_full_table,file_path,file_name)
    
    # 데이터 로드
    upload_dataset=upload_procedure(file_path,file_name)
    
     # 데이터 적재  
    load_database(upload_dataset,'test_db')
