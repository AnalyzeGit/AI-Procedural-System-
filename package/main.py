import pandas as pd
import pyodbc
from textAnalysis import process_text, combined_extraction

# Database connection and reading data into DataFrame
mdb_file_path = r'C:\Users\in2du\PycharmProjects\semanticAnalysis\procedureDB.mdb'
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=' + mdb_file_path + ';'
)
with pyodbc.connect(conn_str) as conn:
    query = "SELECT * FROM [PARAGRAPH]"
    df = pd.read_sql(query, conn)


# Clean and process DataFrame
df['content'] = df['content'].str.replace('\x0b', '')
df[['pos_tags', 'parse_tree', 'ner']] = df['content'].apply(process_text).apply(pd.Series)
df[['ParagraphType','ActionVerb', 'TagetObject']] = df['content'].apply(combined_extraction).apply(pd.Series)

# Save the results
df.to_csv('procedureContent.tsv', sep='\t', index=False)
