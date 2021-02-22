# -*- coding: utf-8 -*-
"""
Created on Thu Feb  4 19:46:02 2021

@author: takashi.shiozawa
"""

import pandas as pd

df = pd.read_csv("fraggle-result.tsv", sep='\t')#tab区切りで読み込み
df_tanimoto = df.sort_values("tanimoto", ascending=False)#降順に並び替え
df_tanimoto["tanimoto_rank"] = list(range(1,len(df)+1))#順位付け
df_s = df_tanimoto.sort_values("fraggle", ascending=False)#降順に並び替え
df_s["fraggle_rank"] = list(range(1,len(df)+1))#順位付け

PandasTools.AddMoleculeColumnToFrame(df_s, molCol='IMAGE', smilesCol='smiles')
PandasTools.SaveXlsxFromFrame(df_s, 'data_frame1.xlsx',molCol='IMAGE',size=(150,150))
