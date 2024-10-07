#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import spacy
import pandas as pd
import docx
import re


# In[ ]:


# Load the spaCy English language model
nlp = spacy.load("en_core_web_sm")

def natural_language_processing(docx_file):
    doc = docx.Document(docx_file)
    npl=[]
    index = 0

    for paragraph in doc.paragraphs:
        if paragraph !='':
            doc = nlp(paragraph.text)
            for token in doc:
                if token.text!='':
                    token_texts=token.text,
                    lemma_texts=token.lemma_
                    pos_tags=token.pos_
                    dependency_relations=token.dep_
                    dependency_heads=token.head.text
                    ner_tags=token.ent_type_ if token.ent_type_ else "-"
                
                    npl.append({
                    'Paragraph Id':index,
                    "Token": token_texts,
                    "Lemma": lemma_texts,
                    "POS": pos_tags,
                    "Dependency": dependency_relations,
                    "Dependency Head": dependency_heads,
                    "NER": ner_tags})
        index += 1
        
    nlp_df = pd.DataFrame(npl)
    return nlp_df

def nlp_index_cleanup(df):
    df=df.reset_index()
    df=df.rename({'index':'Nlp Id'}, axis=1)
    return df

def token_purification(nlp):
    extracted_values = [re.sub(r'[(),]', '', str(item)).strip() for item in nlp['Token']]
    cleaned_data = [item.replace("'", "") for item in extracted_values]
    return cleaned_data

def letter_case(text):
    if text.islower():
        return 'Lower'
    elif text.isupper():
        return 'Upper'
    else:
        return 'Include both'
    
def Apply_case_functions(nlp,texts):
    nlp['Letter Case']='None'
    nlp['Text']=texts
    english=nlp[nlp['Kind']=='English']
    
    for idx,val in english.iterrows():
        letter_case_val=letter_case(val['Text'])
        nlp.loc[idx,'Letter Case']=letter_case_val
    
def classify_text_character(text):
    
    """Classify individual characters in the given text."""
    
    english_pattern = re.compile(r'^[A-Za-z]+$')
    special_characters_pattern = re.compile(r'[!@#$%^&*(),.?":{}|<>\-]')
    number_pattern = re.compile(r'^\d+(\.\d+)?$')
    
    if re.match(special_characters_pattern, text):
        return 'Special Chars & Punct'
    elif re.match(english_pattern,text):
        return 'English'
    elif re.fullmatch(number_pattern, text):
        return 'Number'
    else:
        return 'contain variety things'

def extract_nlp_case(nlp):
    
    """Extract and classify runs in the given dataframe."""
    
    texts=token_purification(nlp)
    nlp['Kind'] = [classify_text_character(text) for text in texts]
    Apply_case_functions(nlp,texts)
    # Assuming 'Run' dataframe has NaNs, replacing them with 'None'
    nlp.fillna('None', inplace=True)

    return nlp

