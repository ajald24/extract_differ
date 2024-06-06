import pandas as pd
import spacy
import ginza
# from collections import Counter
# import string
import streamlit as st

# 日本語ストップワードの読み込み
# with open('stopwords.txt', 'r', encoding='utf-8') as file:
#     stopwords = set(file.read().splitlines())

# SpaCyのGinzaモデルの読み込み
nlp = spacy.load('ja_ginza')

def extract_nouns(text):
    doc = nlp(text)
    # words = [token.text for token in doc if token.pos_ in ('NOUN', 'PROPN', 'ADJ','NUM') and token.text not in stopwords and token.text not in string.punctuation]
    words = [token.text for token in doc if token.pos_ in ('NOUN', 'PROPN', 'ADJ','NUM') and token.text not in string.punctuation]
    # counter = Counter(words)
    # keywords = {word for word, freq in counter.items()}
    return words

def highlight_keywords(text, keywords):
    for keyword in keywords:
        text = text.replace(keyword, f'<{keyword}>')
    return text

def process_excel(input_file, output_file):
    if input_file is not None:
        # Excelファイルを読み込む
        df = pd.read_excel(input_file)

        st.write("Excelファイルの内容を読み込みました。")
        
        col1 = st.selectbox('列を選択1',df.columns)
        col2 = st.selectbox('列を選択2',df.columns)
    
        if st.button('実行'):
            # 各行の処理
            for index, row in df.iterrows():
                text1 = row[col1]
                text2 = row[col2]
                
                # キーフレーズを抽出
                keywords1 = extract_nouns(text1)
                keywords2 = extract_nouns(text2)
                
                # 片方の文章にしか含まれていないキーフレーズ
                unique_to_text1 = set(keywords1) - set(keywords2)
                unique_to_text2 = set(keywords2) - set(keywords1)
    
                # 文章に含まれているキーフレーズを除外
                cleaned_text1 = [txt for txt in unique_to_text1 if txt not in text2]
                cleaned_text2 = [txt for txt in unique_to_text2 if txt not in text1]
                
                # 強調表示
                highlighted_text1 = highlight_keywords(text1, cleaned_text1)
                highlighted_text2 = highlight_keywords(text2, cleaned_text2)
                
                # データフレームに保存
                df.at[index, 'Text1'] = highlighted_text1
                df.at[index, 'Text2'] = highlighted_text2
            
            # 結果をExcelファイルに保存
            df.to_excel(output_file, index=False)
            st.success(f"処理が完了しました。結果は{output_file}に保存されました。")
    else:
        st.warning("Excelファイルをアップロードしてください。")

# ファイルのパス
input_file = st.file_uploader('Excelファイルを選択', type='xlsx')
output_file = 'output.xlsx'

# Excelファイルの処理
process_excel(input_file, output_file)
