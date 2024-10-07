#!/usr/bin/env python
# coding: utf-8

# In[34]:


import pandas as pd


# In[35]:


def upload_procedure(file_name,name):
    
    # 파일 경로 설정
    path=f"{file_name}\\{name}.csv"
    
    # 파일 가져오기
    dataset=pd.read_csv(path)
    
    return dataset

