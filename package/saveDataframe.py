#!/usr/bin/env python
# coding: utf-8

# In[26]:


import pandas as pd


# In[10]:


def save_dataframe(data,file_name,name):   
    
    """
    지정된 경로에 데이터 프레임을 CSV 파일로 저장합니다.
    
    data (str): 원본 문서 
    file_name (str): 데이터 프레임을 저장할 디렉토리의 경로.
    name (str): 저장할 파일의 이름에 추가할 문자열.
    
    """
    
    path = f"{file_name}\\{name}.csv"
    data.to_csv(path, encoding='utf-8-sig',index=False)

