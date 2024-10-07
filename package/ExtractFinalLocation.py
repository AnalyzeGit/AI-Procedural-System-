#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from docx import Document
from difflib import SequenceMatcher
from AppliedLocation import build_location_dataset
from locationApply import *
import re
import win32com.client
import docx
import pandas as pd
import Levenshtein
import glob


# In[ ]:


# Goals: Location 함수 실행

def build_final_location_structure(location,docx_file):
    """
    주어진 위치와 문서 파일을 기반으로 최종 위치 데이터를 생성하고 DataFrame으로 반환합니다.
    
    매개변수:
    - param location: 데이터를 생성할 위치의 경로 또는 식별자
    - param docx_file: 분석할 문서 파일의 경로
    
    반환 값:
    -  위치 데이터를 포함하는 pandas DataFrame
    """
    # 데이터 처리 단계를 순차적으로 수행
    
    # *. 레벨 생성
    create_level(location)
    
    # 2.단락 업데이트
    updated_paragraphs_after_last_heading = update_paragraphs_before_and_after_last_heading(location)
    
    # 3, 섹션 할당
    location_data=assign_sections_to_paragraphs(updated_paragraphs_after_last_heading)
    
    # 3.5 섹션 업데이트
    location_data=update_sections(location_data)
    
    # 4.부모 할당
    location_data=get_parent_create(location_data)
    
    # 5. Stick Numbering 할당
    location_data=extract_stick_add_numering(location_data)
    
    # 6. NCW 할당
    location_data=add_ncw_to_dataframe(location_data)
    
    # 7. NCW 업데이트
    location_data=update_ncw_add_numbering_dataframe(location_data)
    
    # 8. Bullet Unicode 할당 
    location_data=add_bullet_unicode_dataframe(location_data)
    
    # 9. Numbering Type 할당
    location_data=add_numbering_type_dataframe(location_data)
    
    # 10. location 정제 
    location_data=refining_datasets(location_data)
    
    # 11.ㅣocation  최종 정리 
    location_df = build_location_dataset(docx_file, location_data)
    
    # 12. table section update
    location_df=update_table_sections(location_df)
    
    return location_df

