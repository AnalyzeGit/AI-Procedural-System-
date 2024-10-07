#!/usr/bin/env python
# coding: utf-8

# In[2]:


from docx import Document
from difflib import SequenceMatcher
import re
import win32com.client
import docx
import pandas as pd
import Levenshtein
import glob


# In[ ]:


# Goals: Total Level 함수 구현

def determine_Note_level(value, memory_level):
    # 상수를 정의하여 'magic numbers'를 제거하고 코드의 가독성을 높입니다.
    # Note Level 조건 추출
    
    LEVEL_1_UPPER_BOUND = 50
    LEVEL_2_UPPER_BOUND = 88
    LEVEL_3_UPPER_BOUND = 139
    LEVEL_4_UPPER_BOUND = 170
    LEVEL_5_UPPER_BOUND = 180
    LEVEL_6_UPPER_BOUND = 200

    """주어진 값에 따라 적절한 레벨을 반환합니다."""
    if value < 0:
        return None  # 값이 음수인 경우 None을 반환합니다.
    if value < LEVEL_1_UPPER_BOUND:
        return 1.5
    elif value < LEVEL_2_UPPER_BOUND:
        return 2.5
    elif value < LEVEL_3_UPPER_BOUND:
        if memory_level == 1:
            return 1.5
        elif memory_level == 2:
            return 2.5
        elif memory_level == 3:
            return 3.5
        elif memory_level == 4:
            return 4.5
        else:
            return 5.5
    elif value < LEVEL_4_UPPER_BOUND:
        return 4.5
    elif value < LEVEL_5_UPPER_BOUND:
        return 5.5
    elif value < LEVEL_6_UPPER_BOUND:
        return 6.5
    else:
        return None  # 주어진 범위를 벗어나는 값에 대해서는 None을 반환합니다
    
def create_level(df):
    """DataFrame에 새 열을 추가하여 각 제목의 수준을 지정합니다."""
    
    # Initialize variables
    df['Total Level'] = 'None'
    memory_level = 0
    
    #  반복문을 통해 Level 생성
    for idx, val in df.iterrows():
        # 조건변수 변수 대입
        indent=val['Left Indent']
        style = val['Paragraph Style Client']
        paragraph=val['Paragraph']
        numbering=val['Paragraph Numbering Text']
        numbering_type=val['Numbering Type']
        
        # 부모 레벨 담기   
        if style not in  ['머리글','표준','목록 단락','목차 1']:
            if (numbering ==''):
                df.loc[idx, 'Total Level'] = memory_level + 1
            else:
                memory_level=val['Paragraph Level']
                df.loc[idx, 'Total Level']=val['Paragraph Level']
            
        # 제목이 아닌 문단이 들어왔을 때  
        elif style in  ['표준','목록 단락','목차 1']:
             # Note 문단 처리 
            if paragraph.startswith('NOTE'): 
                note_level = determine_Note_level(indent,memory_level)
                df.loc[idx,'Total Level']=note_level
             # Note 아닌 문단 처리   
            else:
                if (numbering =='') | (numbering_type==2):
                    df.loc[idx, 'Total Level'] = memory_level + 1
                elif (paragraph.startswith('-\t')) | ( numbering =='—') :
                    df.loc[idx, 'Total Level'] = memory_level + 2
                else:
                    df.loc[idx, 'Total Level'] = memory_level + 0.5
        # 모든 조건에 속하지 않을 때             
        else:
             df.loc[idx,'Total Level']=0


# In[ ]:


def update_paragraphs_before_and_after_last_heading(location):
    """첫 번째 '헤딩 1' 스타일 전에 문단의 'total_level'을 업데이트하고
        팬더 데이터 프레임에서 마지막 '헤딩 1' 스타일 단락부터 1 단락까지
        
    매개변수:
    - location: pandas DataFrame에는 'Franct Style Client'와 'Total Level' 열이 있는 DataFrames
    
    반환 값:
    - Updated DataFrames.
    
    """
    first_heading_index = None
    last_heading_index = None
    
    # 첫 '제목 1' 스타일 location total_level 변경
    for num, rows in location.iterrows():
        if rows['Paragraph Style Client'] == '제목 1':
            if first_heading_index is None:
                first_heading_index = num  # 마지막 '제목 1' 스타일의 인덱스 갱신
                location.loc[:num,'Total Level'] = 1
            last_heading_index = num
            

    # 마지막 '제목 1' 스타일 이후 total_level 변경
    if last_heading_index is not None:
         location.loc[last_heading_index + 1:, 'Total Level'] = 1

    return location


# In[ ]:


def assign_sections_to_paragraphs(df_paragraphs):
    """
    이 함수는 DataFrame 내의 단락에 섹션 번호를 할당합니다. '제목 1' 스타일의 단락을 섹션의 시작으로 간주하고,
    이후 단락들에 동일한 섹션 번호를 할당합니다. 각 '제목 1' 단락이 나타날 때마다 새로운 섹션 번호로 업데이트됩니다.

    매개변수:
    - df_paragraphs: 섹션 번호를 할당할 단락이 포함된 DataFrame입니다. 'Paragraph Style'과 'Numbering' 컬럼이 필요합니다.

    반환값:
    - DataFrame: 각 단락에 섹션 번호가 할당된 DataFrame을 반환합니다.
    """
    section = 0
    df_paragraphs['SECTION'] = 0

    for idx, val in df_paragraphs.iterrows():
        if val['Paragraph Style Client'] == '제목 1':
            section = val['Paragraph Numbering Text']
        df_paragraphs.loc[idx, 'SECTION'] = section
    
    return df_paragraphs


# In[ ]:


def update_sections(df):
    COVER_PAGE = 1
    TOC_SECTION = 'TOC'
    REVISION_SUMMARY_SECTION = 'RevisionSummary'
    """ 페이지 및 스타일 기준에 따라 DataFrame 섹션을 업데이트합니다 """
    update_section(df, df['Paragraph Page'] == COVER_PAGE, 'SECTION', 'CoverPage')

    topic_pages = df[df['Paragraph Style Client'].str.startswith('목차')]['Paragraph Page'].unique()
    for page in topic_pages:
        update_section(df, df['Paragraph Page'] == page, 'SECTION', TOC_SECTION)

    summary_pages = df[df['Paragraph'].str.startswith('REVISION SUMMARY SHEET')]['Paragraph Page'].unique()
    for page in summary_pages:
        update_section(df, df['Paragraph Page'] == page, 'SECTION', REVISION_SUMMARY_SECTION)

    return df

def update_section(df, condition, column, value):
    """ 
    조건에 따라 DataFrame의 특정 섹션을 업데이트합니다. 
    
    매개변수
    - df:업데이트할 DataFrame입니다.
    - condition: DataFrame이 업데이트 될 조건입니다. 이는 일반적으로 불리언 시리즈
    - column: DataFrame에서 업데이트될 열의 이름
    - value: 주어진 조건을 만족하는 행에 할당될 값
    
    반환 값
    - 명시적인 반환 값은 없습니다. 함수는 인자로 전달된 DataFrame에 직접적으로 작동
    """
    df.loc[condition, column] = value


# In[ ]:


def update_table_sections(df):
    """
   Type이 'Table'인 행에 대한 Section(섹션) 열을 업데이트합니다, 앞의 Non-Table 행의 섹션 값을 기준으로 합니다.
    """
    current_section = None
    for idx, row in df.iterrows():
        if row['Type'] != 'Table':
            current_section = row['SECTION']
        else:
            df.at[idx, 'SECTION'] = current_section

    return df


# In[ ]:


# Goals: 부모 추정 함수 구현(최종 업데이트)_2023.12.11

# level 기준 인덱스 업데이트 함수  
def update_levels(levels, level, num):
    if level <= 6:
        level_key = num_to_ordinal(level) + '_level'
        levels[level_key] = num
    else:
        levels['above_level'] = num

    return levels

# level에 따른 딕셔너리 이름 변경
def num_to_ordinal(number):
    ordinals = ['first', 'second', 'third', 'fourth', 'fifth','sixth']
    return ordinals[number - 1] if 0 < number <= len(ordinals) else 'above_level'

import pandas as pd

def get_parent_index(level, levels):
    if level == 1:
        return 'highest level'
    elif 2 <= level <= 6:
        # 이전 레벨의 값을 확인
        prev_value = levels[num_to_ordinal(level - 1) + '_level']
        
        # 이전 레벨의 값이 NaN이면 더 이전 레벨의 값을 찾음
        if pd.isna(prev_value):
            # 더 이전 레벨로 되돌아가기
            for prev_level in range(level - 2, 0, -1):
                prev_value = levels[num_to_ordinal(prev_level) + '_level']
                if not pd.isna(prev_value):
                    return prev_value
            # 모든 이전 레벨이 NaN인 경우
            return 'N/A'
        else:
            return prev_value
    else:
        return 'N/A'

    
# 0.5 전용 부모 레벨 업데이트 함수
def get_parent_index_else(level, levels):
    if 1 <= level <= 6:
        # 현재 레벨의 값을 확인
        current_value = levels[num_to_ordinal(level) + '_level']
        
        # 현재 레벨의 값이 NaN이면 이전 레벨의 값을 반환
        if pd.isna(current_value):
            # 이전 레벨로 되돌아가기
            for prev_level in range(level - 1, 0, -1):
                prev_value = levels[num_to_ordinal(prev_level) + '_level']
                if not pd.isna(prev_value):
                    return prev_value
            # 모든 이전 레벨이 NaN인 경우
            return 'N/A'
        else:
            return current_value
    else:
        return 'N/A'


def get_parent_create(df):
    # 업데이트 딕셔너리 구현
    #levels={}
    # 예시 키 리스트
    keys = ['first_level', 'second_level', 'third_level', 'fourth_level', 'fifth_level', 'sixth_level']

    # 모든 키에 대해 None 값을 가지는 딕셔너리 생성
    levels = {key: None for key in keys}
    
    # Parent Index 생성
    df['Parent Index']='None'
    for num, rows in df.iterrows():
        #num=rows['ID']
        level = rows['Total Level']
        #num= rows['Paragraph Id']
        if isinstance(level, int):
            levels = update_levels(levels, level, num)
            df.loc[num, 'Parent Index'] = get_parent_index(level, levels)
        else:
            if level is not None:
                # level이 None이 아닌 경우에만 연산을 수행합니다.
                parent_level_key = int(level - 0.5)
                df.loc[num, 'Parent Index'] = get_parent_index_else(parent_level_key, levels)
            else:
                # level이 None인 경우의 처리 로직을 여기에 추가하세요.
                # 예를 들어, 기본값을 설정하거나 다른 처리를 할 수 있습니다.
                pass  # 혹은 적절한 처리를 추가하세요
    return df


# In[ ]:


# Goals: '-' Numbering 추출 함수

def extract_stick_add_numering(df):
    """
    DataFrame의 각 행을 검사하여 'Paragraph' 열이 하이픈('-')으로 시작하는 경우, 
    해당 행에 'Numbering' 열을 추가하고 이를 하이픈('-')으로 표시합니다.
    
    이 함수는 주로 문서에서 리스트 항목을 식별하고 이에 대한 표시를 추가하는 데 사용됩니다.
    
    매개변수
    -param df: 처리할 DataFrame. 'Paragraph' 열을 포함해야 합니다.
    
    반환 값:
    -return: 수정된 DataFrame
    """
        
    for num,rows in df.iterrows():
        if rows['Paragraph'].startswith('-'):
            df.loc[num,'Paragraph Numbering Text']='—'
    return df


# In[ ]:


# Goals: Function to extract "NOTE" patterns from the beginning of a string

def extract_note_pattern_at_start(paragraph):
    # Regular expression pattern to match "NOTE", "NOTE :" and "NOTE 숫자:" at the start of the string
    note_pattern = r'^NOTE( :|\s\d+:)?'

    # Search for the pattern at the start of the text
    match = re.search(note_pattern, paragraph)
    if match:
        return match.group()  # Return the matched string
    else:
        return 0  # If no pattern is found, return None

# Apply the function to each row in the DataFrame column

def add_ncw_to_dataframe(df):
    """
    extract_note_pattern_at_start 알고리즘 적용 후 NCW 필드 생성
    
    매개변수:
    - param df: 처리할 DataFrame. 'Paragraph' 열을 포함해야 합니다.
    
    반환 값:
    -return: 수정된 DataFrame
    """ 
    
    df['NCW'] = df['Paragraph'].apply(extract_note_pattern_at_start)
    return df


# In[ ]:


# Goals: NCW Type Numbering 필드에 업데이트

def update_ncw_add_numbering_dataframe(df):
    """
    생성된 NCW 필드를 활용하여 Numbering Update
    
    매개변수:
    - param df: 처리할 DataFrame. 'Paragraph' 열을 포함해야 합니다.
    
    반환 값:
    - return: 수정된 DataFrame
    """ 
    note_bowl=[]

    for num,rows in df.iterrows():
        if rows['NCW']!=0:
            df.loc[num,'Paragraph Numbering Text']=rows['NCW']
    return df


# In[ ]:


# Goals: Numbering Type 업데이트

def match_and_replace(var):
    
    # 입력값이 문자열이 아닌 경우 문자열로 변환
    if not isinstance(var, str):
        var = str(var)    
    # 공백 제거
    var_no_space = var.replace(" ", "")
    
    if re.match(r'^[A-Z]\.$', var_no_space):
        return 'A.'
    elif re.match(r'^[A-Z](\.\d+)+$', var_no_space):
        # Count the number of '.number' patterns
        count = var_no_space.count('.')
        return 'A'+'.n'*count
    elif re.match(r'^[a-z]\.$',var_no_space):
        return 'a.'
    elif re.match('^[a-z](\.\d+)+$',var_no_space):
        count = var_no_space.count('.')
        return 'a'+'.n'*count
    elif re.match(r'^\d+$', var_no_space):
        return 'n'
    elif re.match(r'^\d+\)$', var_no_space):
        return 'n)'
    elif re.match(r'^\d+\.$', var_no_space):
        return 'n.'
    elif re.match(r'^\d+\.\d+(\.\d+)*$', var_no_space):
        count = var_no_space.count('.')
        return 'n'+'.n'*count
    elif re.match(r'^NOTE\d+:',var_no_space):
        return 'NOTE:n'
    elif re.match(r'^NOTE$',var_no_space):
        return 'NOTE'
    elif (var_no_space=='125A0') | (var_no_space=='B23B0'):
        return 'Bullet'
    else:
        return var
    

def add_numbering_type_dataframe(df):
    """
   match_and_replace 함수를 활용하여 Numbering Type Code 생성
    
    매개변수:
    - param df: 처리할 DataFrame. 'Paragraph' 열을 포함해야 합니다.
    
    반환 값:
    - return: 수정된 DataFrame
    """ 
    
    df['Numbering Type Code'] = [match_and_replace(i) for i in df['Paragraph Numbering Text']]
    return df


# In[ ]:


# Goals: Bullet Unicode Update

def replace_bullet_point_unicode(numbering):
    if isinstance(numbering, str) and len(numbering) == 1:  # 문자열이면서 길이가 1인 경우
        # 아스키 코드 저장
        asciicode= ord(numbering)
        if asciicode==61623:
            return '125A0'
        elif asciicode==8212:
            return 'B23B0'
    else:
        if numbering is None:
            return 'No Numbering'
        else:
            return numbering  # 입력이 단일 문자가 아님
    
# 예제 변수들
def add_bullet_unicode_dataframe(df):
    """
   replace_bullet_point_unicode 함수를 활용하여 Bullet Unicode 생성
    
    매개변수:
    - param df: 처리할 DataFrame. 'Paragraph' 열을 포함해야 합니다.
    
    반환 값:
    - return: 수정된 DataFrame
    """ 
        
    numbering = df['Paragraph Numbering Text']
    df['Paragraph Numbering Text']=[replace_bullet_point_unicode(bullet) for bullet in df['Paragraph Numbering Text']]
    return df


# In[ ]:


def delete_missing_values(df):
    """
    주어진 DataFrame에서 결측치(null 또는 NA)를 포함하는 모든 행을 제거합니다.
    
    매개변수:
    - df: 결측치를 제거할 DataFrame 객체입니다.
    
    반환값:
    - DataFrame: 결측치가 제거된 DataFrame을 반환합니다.
    """
    df=df.dropna()
    return df

def delete_meaningless_paragraph(df):
    """
    DataFrame에서 특정 시작 패턴을 가진 의미없는 단락을 제거합니다.
    이 함수는 '(EXPA', 'Continued next page', 'cont'로 시작하는 단락을 제거합니다.
    
    매개변수:
    - df: 의미없는 단락을 제거할 DataFrame 객체입니다.
    
    반환값:
    - DataFrame: 의미없는 단락이 제거된 DataFrame을 반환합니다.
    """
    delete_list=[]
    
    for idx,val in df.iterrows():
        paragraph=val['Paragraph']
        
        if (paragraph.startswith('(EXPA')) | (paragraph.startswith('Continued next page')) | (paragraph.startswith('cont')):
            delete_list.append(idx)
    df.drop(delete_list,inplace=True)
    
    return df

def refining_datasets(df):
    """
    주어진 DataFrame을 정제하는 과정을 수행합니다. 이 과정에는 결측치 제거 및 의미없는 단락 제거가 포함됩니다.
    
    매개변수:
    - df: 정제할 DataFrame 객체입니다.
    
    반환값:
    - DataFrame: 정제된 DataFrame을 반환합니다.
    """
    delete_missing_dataset=delete_missing_values(df)
    refining_dataset=delete_meaningless_paragraph(delete_missing_dataset)
    
    return refining_dataset


# In[ ]:


def extract_table_text(document_path):
    """
    이 함수는 주어진 Word 문서(.docx 파일)에서 모든 테이블의 텍스트를 추출합니다.
    각 테이블은 별도의 텍스트 블록으로 추출되며, 테이블의 각 셀은 탭으로 구분되고 행은 새 줄로 구분됩니다.
    
    매개변수:
    - document_path: 분석할 Word 문서의 경로입니다.
    
    반환값:
    - DataFrame: 추출된 테이블의 텍스트를 포함하는 pandas DataFrame을 반환합니다. 
      각 테이블은 고유한 'Table Id'로 식별되며, 'Text' 컬럼에 해당 테이블의 텍스트가 저장됩니다.
    
    함수 사용 예시:
    df = extract_table_text('example.docx')
    """
    doc = Document(document_path)
    table_contents = []
    table_id=0

    for table in doc.tables:
        table_text = ""
        try:
            for row in table.rows:
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    table_text += cell_text + "\t"  # Add tab separator between cell texts
                table_text += "\n"  # Add newline after each row
        except Exception as e:
            print(f"An error occurred: {e}")
                
        table_contents.append({
            'Table Id': table_id,
            'Text': table_text})

        table_id+=1
    return pd.DataFrame(table_contents)


# In[ ]:


def extract_table_location(document_path):
    """
    이 함수는 주어진 Word 문서(.docx 파일)에서 모든 테이블의 텍스트와 위치 정보를 추출합니다.
    추출된 정보에는 테이블의 텍스트, 스타일, 유형, 페이지 번호, 시작 및 끝 위치가 포함됩니다.
    
    매개변수:
    - document_path: 분석할 Word 문서의 경로입니다.
    
    반환값:
    - DataFrame: 추출된 테이블의 텍스트와 위치 정보를 포함하는 pandas DataFrame을 반환합니다.
      각 테이블은 'Text', 'Style', 'type', 'Page Number', 'start', 'end' 컬럼으로 정보가 저장됩니다.
    
    """
    
    word_app = win32com.client.gencache.EnsureDispatch('Word.Application')
    word_app.Visible = False
    doc = word_app.Documents.Open(document_path)

    positions = []

    for table in doc.Tables:
        table_text = ""
        try:
            for row in table.Rows:
                for cell in row.Cells:
                    cell_text = cell.Range.Text.strip()
                    table_text += cell_text + "\t"  # Add tab separator between cell texts
                table_text += "\n"  # Add newline after each row
        except Exception as e:
            print(f"An error occurred: {e}")
        
        positions.append({
            'Text': table_text,
            'Style':'None',
            'type': 'table',
            'Page Number': table.Range.Information(win32com.client.constants.wdActiveEndPageNumber),
            'start': table.Range.Start,
            'end': table.Range.End
        })

    # ... (remaining code)

    positions = sorted(positions, key=lambda x: x['start'])           
    return pd.DataFrame(positions)


# In[ ]:


def table_calculate_similarity(str1, str2):
    """
    두 문자열 간의 유사도를 계산하여 반환합니다. 이 함수는 레벤슈타인 거리(Levenshtein distance)를 사용하여 
    두 문자열 간의 차이를 측정하고, 이를 유사도 점수로 변환합니다.
    
    매개변수:
    - str1 (str): 비교할 첫 번째 문자열입니다.
    - str2 (str): 비교할 두 번째 문자열입니다.
    
    반환값:
    - float: 두 문자열 간의 유사도 점수를 반환합니다. 값이 클수록 문자열이 더 유사합니다.
    """
    return 1 - Levenshtein.distance(str1, str2) / max(len(str1), len(str2))


def table_map_sentences_levenshtein(Paragraph, Paragraph_test):
    """
    두 DataFrame 간의 텍스트 유사도를 계산하여 매핑합니다. 각 단락의 텍스트를 다른 DataFrame의 단락과 비교하여 
    가장 유사한 단락을 찾고, 그 정보를 새로운 DataFrame으로 반환합니다.
    
    매개변수:
    - Paragraph (DataFrame): 유사도를 계산할 첫 번째 DataFrame입니다.
    - Paragraph_test (DataFrame): 유사도를 계산할 두 번째 DataFrame입니다.
    
    반환값:
    - DataFrame: 매핑된 단락의 정보를 포함하는 DataFrame을 반환합니다. 각 행은 첫 번째 DataFrame의 단락과 
      두 번째 DataFrame에서 찾은 가장 유사한 단락의 정보를 포함합니다.
    """
    data = {
        'ID': [],
        'Text': [],
        'Loc Text': [],
        'Page': [],
        'score': [],
        'Start': [] ,
        'End':[],
        'Style':[]}
    
     # 인덱스 값을 저장하기 위한 새로운 컬럼
    comparative_idx=[]
    page=0
    
    for idx, val in Paragraph.iterrows():
        Doc_Paragraph = val['Text']
        comparative_Paragraph=[]
        similarity_scores = {}
        idx_scores = {}  # 각 문장의 인덱스 값을 저장하기 위한 딕셔너리
        
        for test_idx, test_val in Paragraph_test.iterrows():
            # 매칭된 파라그래프, 중복된 파라그래프 조건
            if (test_idx not in comparative_idx) & (test_val['Text'] not in comparative_Paragraph): 
                Client_Paragraph = test_val['Text']
                similarity_scores[Client_Paragraph] = table_calculate_similarity(Doc_Paragraph, Client_Paragraph)
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
        
            data['Text'].append(Doc_Paragraph)
            data['Loc Text'].append(best_match)
            data['score'].append(best_score)
            data['ID'].append(idx)
            data['Start'].append(Paragraph_test.loc[best_match_idx,'start'])
            data['End'].append(Paragraph_test.loc[best_match_idx,'end'])
            data['Page'].append(Paragraph_test.loc[best_match_idx,'Page Number'])
            data['Style'].append(Paragraph_test.loc[best_match_idx,'Style'])
            page+=Paragraph_test.loc[best_match_idx,'Page Number']
        #딕셔너리가 존재하지 않을때
        else:
            data['Text'].append(Doc_Paragraph)
            data['Loc Text'].append('None')
            data['score'].append('None')
            data['ID'].append('None')
            data['Start'].append('None')
            data['End'].append('None')
            data['Paragraph Page'].append('None')
            data['Style'].append('None')
        # 여기에 다른 데이터 (Numbering, Page 등)를 추가하는 코드를 넣을 수 있습니다.
        #Numbering Type	,Paragraph
    return pd.DataFrame(data)


# In[ ]:


def determine_text_type(row):
    """
    주어진 DataFrame 행의 내용과 속성을 기반으로 텍스트 유형을 결정합니다. 
    'CAUTION\n'으로 시작하고 유형이 'Table'인 텍스트는 'Caution Warning'으로,
    'NOTE:'로 시작하는 텍스트는 'NOTE Warning'으로 분류합니다. 
    위 조건에 해당하지 않는 경우 'Normal'로 분류합니다.
    
    매개변수:
    - row: DataFrame의 각 행을 나타내는 데이터입니다.
    
    반환값:
    - str: 결정된 텍스트 유형을 문자열로 반환합니다.
    """
    note_pattern = r'^NOTE( :|\s\d+:)?'
    
    if row['Text'].startswith('CAUTION\n') and row['Type'] == 'Table':
        return 'Caution Warning'
    elif re.search(note_pattern, row['Text']): 
        return 'NOTE Warning'
    else:
        return 'Normal'

def create_indentation(location_df):
    """
    주어진 DataFrame에 새로운 열을 추가하여 각 텍스트 요소의 유형을 지정합니다.
    'Indentation'이라는 이름의 열을 추가하고, 각 텍스트 요소의 내용과 유형을 바탕으로 
    해당 텍스트의 유형을 결정하는 'determine_text_type' 함수를 사용합니다.
    
    매개변수:
    - location_df: 텍스트 유형을 결정할 DataFrame입니다.
    
    반환값:
    - 반환 값은 없으며, 입력된 DataFrame에 직접 변경을 적용합니다.
    """
    
    # Initialize a new column with default value 'Normal'
    location_df['Indentation'] = 'Normal'
    
    # Update the 'Text Type' based on the text content and type
    location_df['Indentation'] = location_df.apply(determine_text_type, axis=1)


# In[ ]:


def reset_id(df):
    """
    위치 데이터를 전처리합니다. 'dataframe'을 reset Id 합니다.
    
    매개변수:
    - location_df: 전처리할 위치 데이터를 포함하는 DataFrame입니다.
    
    반환값:
    - DataFrame: reset Id 과정을 거친 DataFrame을 반환합니다.
    """
    df=df.reset_index()
    df=df.rename({'index':'Unique Id'},axis=1)
    return df

