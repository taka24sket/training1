# -*- coding: utf-8 -*-
"""
Created on Mon Jan  4 11:06:56 2021

@author: takashi.shiozawa
"""



"""

1.作業ディレクトリの設定

フルパスの調べ方
作業ディレクトリのフォルダを「shift+右クリック」のメニューでパスのコピーを選択
spyderを用いているのであれば、右上のフォルダマークからディレクトリの選択可能

"""

#おおもとの作業(モジュールのimportとディレクトリの設定)
import numpy as np
import pandas as pd
import os

os.chdir("C:\\Users\\takashi.shiozawa\\Desktop\\Python")

"""

2.様々な名前や要素の指定

決める要素は「インポートするexcel名」、「それぞれの試験の活性の足切りの値」、「出力するexcel名」

"""

read_excel_name = "HCP-487.xlsx" #エクセル名
water_farming = 5
paddy_1 = 5
paddy_2 = 5
upland_soil = 5
upland_stem_leaf = 5
upland_soil_2 = 5
upland_stem_leaf_2 = 5 #活性値
sheet_names = [0,1,2] #シートの枚数分の数値を打ち込む 3枚なら[0,1,2]

#出力名
water_farming_excel = "water_farming.xlsx"
paddy_1_excel = "paddy_1.xlsx"
paddy_2_excel = "paddy_2.xlsx"
upland_soil_excel = "upland_soil.xlsx"
upland_stem_leaf_excel = "upland_stem_leaf.xlsx"
upland_soil_2_excel = "upland_soil_2.xlsx"
upland_stem_leaf_2_excel = "upland_stem_leaf_2.xlsx"

#カラム名
water_farming_column =["compound","concentration","ECHCG","LOLMU","LASCA","AMARE","SCPJO"]
paddy_1_column = ["compound","concentration","ORY2","ECHOR","MOOVP","SCPJO"]
paddy_2_column = ["compound","concentration","ORY2","ECHOR","MOOVP","SCPJO"]
upland_soil_column = ["compound","concentration","ECHCG","LOLMU","ABUTH","AMARE"]
upland_stem_leaf_column = ["compound","concentration","ECHCG","LOLMU","ABUTH","AMARE"]
upland_soil_2_column = ["compound","concentration","TRZAX","BRSNN","AVEFA","ALOMY","ZEAMX","GLXMA","ORYSA","ECHCG","POLLN","CHEAL","ABUTH","IPOLA","AMBEL"]
upland_stem_leaf_2_column = ["compound","concentration","TRZAX","BRSNN","AVEFA","ALOMY","ZEAMX","GLXMA","ORYSA","ECHCG","POLLN","CHEAL","ABUTH","IPOLA","AMBEL"]


#exlelファイルの読み込み
#excleの読み込み 実施試験の数に合わせてsheet_nameの部分の数字を減らす(4つなら[0,1,2,3],2つなら[0,1])この操作しないと途中で整合性が合わずに止まる
original_list = pd.read_excel(read_excel_name, header =None, sheet_name=sheet_names) #0はじまりでシートの枚数と対応させる
#畑作土壌1次と畑作茎葉1次のみならsheet_name = [0,1]、水田1次と畑作土壌1次と畑作茎葉2次ならsheet_name = [0,1,2]

"""
例文
上の読み込み時にsheet_nameで指定した数字と試験名を一致させる(畑作土壌1次と畑作茎葉1次のみなら)
list_upland_soil = original_list[0]
list_upland_stem_leaf = original_list[1]

#水田1次と畑作土壌1次と畑作茎葉2次なら
list_paddy_1 = original_list[0]
list_upland_soil = original_list[1]
list_upland_stem_leaf_2 = original_list[2]
"""
   
list_water_farming = original_list[0]
list_paddy_1 = original_list[0]
list_paddy_2 = original_list[1]
list_upland_soil = original_list[2]
list_upland_stem_leaf = original_list[4]
list_upland_soil_2 = original_list[5]
list_upland_stem_leaf_2 = original_list[6]


"""
以降はいじらなくてよい
"""

#%%
#水耕試験("ECHCG","LOLMU","LASCA","AMARE","SCPJO")
try:
    list_water_farming.columns = water_farming_column    
    list_water_farming = list_water_farming[list_water_farming["compound"] != "H"]#Hの入った行の除去
    list_water_farming = list_water_farming[list_water_farming["compound"] != "化合物"]#化合物の入った行の除去
    list_water_farming = list_water_farming.fillna(method="ffill")#上の文字で置換
    list_water_farming = list_water_farming.replace("-",0)

#型の変換(表に-が入っていると数字の部分が文字列と認識されてしまっているため
    list1 = ["concentration","ECHCG","LOLMU","LASCA","AMARE","SCPJO"]
    for i in list1:
        list_water_farming[i] = list_water_farming[i].astype(float)
        list_water_farming.dtypes

#草種ごとに分類
    list_water_farming_ECHCG = list_water_farming.loc[:,["compound","concentration","ECHCG"]]
    list_water_farming_LOLMU = list_water_farming.loc[:,["compound","concentration","LOLMU"]]
    list_water_farming_LASCA = list_water_farming.loc[:,["compound","concentration","LASCA"]]
    list_water_farming_AMARE = list_water_farming.loc[:,["compound","concentration","AMARE"]]
    list_water_farming_SCPJO = list_water_farming.loc[:,["compound","concentration","SCPJO"]]


#任意の活性以下のものを省く
    list_water_farming_ECHCG = list_water_farming_ECHCG[list_water_farming_ECHCG.ECHCG >= water_farming]
    list_water_farming_LOLMU = list_water_farming_LOLMU[list_water_farming_LOLMU.LOLMU >= water_farming]
    list_water_farming_LASCA = list_water_farming_LASCA[list_water_farming_LASCA.LASCA >= water_farming]
    list_water_farming_AMARE = list_water_farming_AMARE[list_water_farming_AMARE.AMARE >= water_farming]
    list_water_farming_SCPJO = list_water_farming_SCPJO[list_water_farming_SCPJO.SCPJO >= water_farming]

    #順番を逆にする(活性が同じ時により低い濃度を抽出するために必要)
    list_water_farming_ECHCG_2 = list_water_farming_ECHCG.sort_index(ascending=False)
    list_water_farming_LOLMU_2 = list_water_farming_LOLMU.sort_index(ascending=False)
    list_water_farming_LASCA_2 = list_water_farming_LASCA.sort_index(ascending=False)
    list_water_farming_AMARE_2 = list_water_farming_AMARE.sort_index(ascending=False)
    list_water_farming_SCPJO_2 = list_water_farming_SCPJO.sort_index(ascending=False)
    
#グループ化
    list_water_farming_ECHCG_3 = list_water_farming_ECHCG_2.groupby("compound")
    list_water_farming_LOLMU_3 = list_water_farming_LOLMU_2.groupby("compound")
    list_water_farming_LASCA_3 = list_water_farming_LASCA_2.groupby("compound")
    list_water_farming_AMARE_3 = list_water_farming_AMARE_2.groupby("compound")
    list_water_farming_SCPJO_3 = list_water_farming_SCPJO_2.groupby("compound")

#抽出
    list_water_farming_ECHCG_final = list_water_farming_ECHCG.loc[list_water_farming_ECHCG_3["ECHCG"].idxmin(),:]
    list_water_farming_LOLMU_final = list_water_farming_LOLMU.loc[list_water_farming_LOLMU_3["LOLMU"].idxmin(),:]
    list_water_farming_LASCA_final = list_water_farming_LASCA.loc[list_water_farming_LASCA_3["LASCA"].idxmin(),:]
    list_water_farming_AMARE_final = list_water_farming_AMARE.loc[list_water_farming_AMARE_3["AMARE"].idxmin(),:]
    list_water_farming_SCPJO_final = list_water_farming_SCPJO.loc[list_water_farming_SCPJO_3["SCPJO"].idxmin(),:]



#活性低いものもの抽出1
    list_water_farming_ECHCG_low_1 = list_water_farming.loc[:,["compound","concentration","ECHCG"]]
    list_water_farming_LOLMU_low_1 = list_water_farming.loc[:,["compound","concentration","LOLMU"]]
    list_water_farming_LASCA_low_1 = list_water_farming.loc[:,["compound","concentration","LASCA"]]
    list_water_farming_AMARE_low_1 = list_water_farming.loc[:,["compound","concentration","AMARE"]]
    list_water_farming_SCPJO_low_1 = list_water_farming.loc[:,["compound","concentration","SCPJO"]]



#活性低いものもの抽出2
    list_water_farming_ECHCG_low_2 = list_water_farming_ECHCG_low_1[list_water_farming_ECHCG_low_1["concentration"] == 200]
    list_water_farming_LOLMU_low_2 = list_water_farming_LOLMU_low_1[list_water_farming_LOLMU_low_1["concentration"] == 200]
    list_water_farming_LASCA_low_2 = list_water_farming_LASCA_low_1[list_water_farming_LASCA_low_1["concentration"] == 200]
    list_water_farming_AMARE_low_2 = list_water_farming_AMARE_low_1[list_water_farming_AMARE_low_1["concentration"] == 200]
    list_water_farming_SCPJO_low_2 = list_water_farming_SCPJO_low_1[list_water_farming_SCPJO_low_1["concentration"] == 200]

#活性低いものもの抽出3
    list_water_farming_ECHCG_low_3 = list_water_farming_ECHCG_low_2[list_water_farming_ECHCG_low_2.ECHCG < water_farming]
    list_water_farming_LOLMU_low_3 = list_water_farming_LOLMU_low_2[list_water_farming_LOLMU_low_2.LOLMU < water_farming]
    list_water_farming_LASCA_low_3 = list_water_farming_LASCA_low_2[list_water_farming_LASCA_low_2.LASCA < water_farming]
    list_water_farming_AMARE_low_3 = list_water_farming_AMARE_low_2[list_water_farming_AMARE_low_2.AMARE < water_farming]
    list_water_farming_SCPJO_low_3 = list_water_farming_SCPJO_low_2[list_water_farming_SCPJO_low_2.SCPJO < water_farming]
    

#活性低いものもの抽出4
    list_water_farming_ECHCG_low_4 = list_water_farming_ECHCG_low_3.groupby("compound")
    list_water_farming_LOLMU_low_4 = list_water_farming_LOLMU_low_3.groupby("compound")
    list_water_farming_LASCA_low_4 = list_water_farming_LASCA_low_3.groupby("compound")
    list_water_farming_AMARE_low_4 = list_water_farming_AMARE_low_3.groupby("compound")
    list_water_farming_SCPJO_low_4 = list_water_farming_SCPJO_low_3.groupby("compound")


#活性低いものもの抽出5
    list_water_farming_ECHCG_low_final = list_water_farming_ECHCG_low_3.loc[list_water_farming_ECHCG_low_4["ECHCG"].idxmax(),:]
    list_water_farming_LOLMU_low_final = list_water_farming_LOLMU_low_3.loc[list_water_farming_LOLMU_low_4["LOLMU"].idxmax(),:]
    list_water_farming_LASCA_low_final = list_water_farming_LASCA_low_3.loc[list_water_farming_LASCA_low_4["LASCA"].idxmax(),:]
    list_water_farming_AMARE_low_final = list_water_farming_AMARE_low_3.loc[list_water_farming_AMARE_low_4["AMARE"].idxmax(),:]
    list_water_farming_SCPJO_low_final = list_water_farming_SCPJO_low_3.loc[list_water_farming_SCPJO_low_4["SCPJO"].idxmax(),:]


#活性の高いものと低いものを結合
    list_water_farming_ECHCG_final_2 = pd.concat([list_water_farming_ECHCG_final, list_water_farming_ECHCG_low_final]) 
    list_water_farming_LOLMU_final_2 = pd.concat([list_water_farming_LOLMU_final, list_water_farming_LOLMU_low_final]) 
    list_water_farming_LASCA_final_2 = pd.concat([list_water_farming_LASCA_final, list_water_farming_LASCA_low_final]) 
    list_water_farming_AMARE_final_2 = pd.concat([list_water_farming_AMARE_final, list_water_farming_AMARE_low_final]) 
    list_water_farming_SCPJO_final_2 = pd.concat([list_water_farming_SCPJO_final, list_water_farming_SCPJO_low_final]) 

#複数ある場合を除く
    list_water_farming_ECHCG_final_3 = list_water_farming_ECHCG_final_2.groupby("compound")
    list_water_farming_LOLMU_final_3 = list_water_farming_LOLMU_final_2.groupby("compound")
    list_water_farming_LASCA_final_3 = list_water_farming_LASCA_final_2.groupby("compound")
    list_water_farming_AMARE_final_3 = list_water_farming_AMARE_final_2.groupby("compound")
    list_water_farming_SCPJO_final_3 = list_water_farming_SCPJO_final_2.groupby("compound")

    list_water_farming_ECHCG_final_4 = list_water_farming_ECHCG_final_2.loc[list_water_farming_ECHCG_final_3["ECHCG"].idxmax(),:]
    list_water_farming_LOLMU_final_4 = list_water_farming_LOLMU_final_2.loc[list_water_farming_LOLMU_final_3["LOLMU"].idxmax(),:]
    list_water_farming_LASCA_final_4 = list_water_farming_LASCA_final_2.loc[list_water_farming_LASCA_final_3["LASCA"].idxmax(),:]
    list_water_farming_AMARE_final_4 = list_water_farming_AMARE_final_2.loc[list_water_farming_AMARE_final_3["AMARE"].idxmax(),:]
    list_water_farming_SCPJO_final_4 = list_water_farming_SCPJO_final_2.loc[list_water_farming_SCPJO_final_3["SCPJO"].idxmax(),:]
    
#出力
    with pd.ExcelWriter(water_farming_excel) as writer:
        list_water_farming_ECHCG_final_4.to_excel(writer, sheet_name = "ECHCG", index=False)
        list_water_farming_LOLMU_final_4.to_excel(writer, sheet_name = "LOLMU", index=False)
        list_water_farming_LASCA_final_4.to_excel(writer, sheet_name = "LASCA", index=False)
        list_water_farming_AMARE_final_4.to_excel(writer, sheet_name = "AMARE", index=False)
        list_water_farming_SCPJO_final_4.to_excel(writer, sheet_name = "SCPJO", index=False)
except:
    pass





#%%
#水田1次("ORY2","ECHOR","MOOVP","SCPJO")
try:
    list_paddy_1.columns = paddy_1_column
    list_paddy_1 = list_paddy_1[list_paddy_1["compound"] != "H"]
    list_paddy_1 = list_paddy_1[list_paddy_1["compound"] != "化合物"]
    list_paddy_1 = list_paddy_1.fillna(method="ffill")
    list_paddy_1 = list_paddy_1.replace("-",0)

#型の変換(表に-が入っていると数字の部分が文字列と認識されてしまっているため
    list1 = ["concentration","ORY2","ECHOR","MOOVP","SCPJO"]
    for i in list1:
        list_paddy_1[i] = list_paddy_1[i].astype(float)
        list_paddy_1.dtypes

#草種ごとに分類
    list_paddy_1_ORY2 = list_paddy_1.loc[:,["compound","concentration","ORY2"]]
    list_paddy_1_ECHOR = list_paddy_1.loc[:,["compound","concentration","ECHOR"]]
    list_paddy_1_MOOVP = list_paddy_1.loc[:,["compound","concentration","MOOVP"]]
    list_paddy_1_SCPJO = list_paddy_1.loc[:,["compound","concentration","SCPJO"]]

#任意の活性以下のものを省く
    list_paddy_1_ORY2 = list_paddy_1_ORY2[list_paddy_1_ORY2.ORY2 >= paddy_1]
    list_paddy_1_ECHOR = list_paddy_1_ECHOR[list_paddy_1_ECHOR.ECHOR >= paddy_1]
    list_paddy_1_MOOVP = list_paddy_1_MOOVP[list_paddy_1_MOOVP.MOOVP >= paddy_1]
    list_paddy_1_SCPJO = list_paddy_1_SCPJO[list_paddy_1_SCPJO.SCPJO >= paddy_1]

#順番を逆にする
    list_paddy_1_ORY2_2 = list_paddy_1_ORY2.sort_index(ascending=False)
    list_paddy_1_ECHOR_2 = list_paddy_1_ECHOR.sort_index(ascending=False)
    list_paddy_1_MOOVP_2 = list_paddy_1_MOOVP.sort_index(ascending=False)
    list_paddy_1_SCPJO_2 = list_paddy_1_SCPJO.sort_index(ascending=False)

#グループ化
    list_paddy_1_ORY2_3 = list_paddy_1_ORY2_2.groupby("compound")
    list_paddy_1_ECHOR_3 = list_paddy_1_ECHOR_2.groupby("compound")
    list_paddy_1_MOOVP_3 = list_paddy_1_MOOVP_2.groupby("compound")
    list_paddy_1_SCPJO_3 = list_paddy_1_SCPJO_2.groupby("compound")

#抽出
    list_paddy_1_ORY2_final = list_paddy_1_ORY2.loc[list_paddy_1_ORY2_3["ORY2"].idxmin(),:]
    list_paddy_1_ECHOR_final = list_paddy_1_ECHOR.loc[list_paddy_1_ECHOR_3["ECHOR"].idxmin(),:]
    list_paddy_1_MOOVP_final = list_paddy_1_MOOVP.loc[list_paddy_1_MOOVP_3["MOOVP"].idxmin(),:]
    list_paddy_1_SCPJO_final = list_paddy_1_SCPJO.loc[list_paddy_1_SCPJO_3["SCPJO"].idxmin(),:]



#活性低いものもの抽出1
    list_paddy_1_ORY2_low_1 = list_paddy_1.loc[:,["compound","concentration","ORY2"]]
    list_paddy_1_ECHOR_low_1 = list_paddy_1.loc[:,["compound","concentration","ECHOR"]]
    list_paddy_1_MOOVP_low_1 = list_paddy_1.loc[:,["compound","concentration","MOOVP"]]
    list_paddy_1_SCPJO_low_1 = list_paddy_1.loc[:,["compound","concentration","SCPJO"]]


#活性低いものもの抽出2
    list_paddy_1_ORY2_low_2 = list_paddy_1_ORY2_low_1[list_paddy_1_ORY2_low_1["concentration"] == 100]
    list_paddy_1_ECHOR_low_2 = list_paddy_1_ECHOR_low_1[list_paddy_1_ECHOR_low_1["concentration"] == 100]
    list_paddy_1_MOOVP_low_2 = list_paddy_1_MOOVP_low_1[list_paddy_1_MOOVP_low_1["concentration"] == 100]
    list_paddy_1_SCPJO_low_2 = list_paddy_1_SCPJO_low_1[list_paddy_1_SCPJO_low_1["concentration"] == 100]

#活性低いものもの抽出3
    list_paddy_1_ORY2_low_3 = list_paddy_1_ORY2_low_2[list_paddy_1_ORY2_low_2.ORY2 < paddy_1]
    list_paddy_1_ECHOR_low_3 = list_paddy_1_ECHOR_low_2[list_paddy_1_ECHOR_low_2.ECHOR < paddy_1]
    list_paddy_1_MOOVP_low_3 = list_paddy_1_MOOVP_low_2[list_paddy_1_MOOVP_low_2.MOOVP < paddy_1]
    list_paddy_1_SCPJO_low_3 = list_paddy_1_SCPJO_low_2[list_paddy_1_SCPJO_low_2.SCPJO < paddy_1]

#活性低いものもの抽出4
    list_paddy_1_ORY2_low_4 = list_paddy_1_ORY2_low_3.groupby("compound")
    list_paddy_1_ECHOR_low_4 = list_paddy_1_ECHOR_low_3.groupby("compound")
    list_paddy_1_MOOVP_low_4 = list_paddy_1_MOOVP_low_3.groupby("compound")
    list_paddy_1_SCPJO_low_4 = list_paddy_1_SCPJO_low_3.groupby("compound")

#活性低いものもの抽出5
    list_paddy_1_ORY2_low_final = list_paddy_1_ORY2_low_3.loc[list_paddy_1_ORY2_low_4["ORY2"].idxmax(),:]
    list_paddy_1_ECHOR_low_final = list_paddy_1_ECHOR_low_3.loc[list_paddy_1_ECHOR_low_4["ECHOR"].idxmax(),:]
    list_paddy_1_MOOVP_low_final = list_paddy_1_MOOVP_low_3.loc[list_paddy_1_MOOVP_low_4["MOOVP"].idxmax(),:]
    list_paddy_1_SCPJO_low_final = list_paddy_1_SCPJO_low_3.loc[list_paddy_1_SCPJO_low_4["SCPJO"].idxmax(),:]


#活性の高いものと低いものを結合
    list_paddy_1_ORY2_final_2 = pd.concat([list_paddy_1_ORY2_final, list_paddy_1_ORY2_low_final]) 
    list_paddy_1_ECHOR_final_2 = pd.concat([list_paddy_1_ECHOR_final, list_paddy_1_ECHOR_low_final]) 
    list_paddy_1_MOOVP_final_2 = pd.concat([list_paddy_1_MOOVP_final, list_paddy_1_MOOVP_low_final]) 
    list_paddy_1_SCPJO_final_2 = pd.concat([list_paddy_1_SCPJO_final, list_paddy_1_SCPJO_low_final]) 


#複数ある場合を除く
    list_paddy_1_ORY2_final_3 = list_paddy_1_ORY2_final_2.groupby("compound")
    list_paddy_1_ECHOR_final_3 = list_paddy_1_ECHOR_final_2.groupby("compound")
    list_paddy_1_MOOVP_final_3 = list_paddy_1_MOOVP_final_2.groupby("compound")
    list_paddy_1_SCPJO_final_3 = list_paddy_1_SCPJO_final_2.groupby("compound")


    list_paddy_1_ORY2_final_4 = list_paddy_1_ORY2_final_2.loc[list_paddy_1_ORY2_final_3["ORY2"].idxmax(),:]
    list_paddy_1_ECHOR_final_4 = list_paddy_1_ECHOR_final_2.loc[list_paddy_1_ECHOR_final_3["ECHOR"].idxmax(),:]
    list_paddy_1_MOOVP_final_4 = list_paddy_1_MOOVP_final_2.loc[list_paddy_1_MOOVP_final_3["MOOVP"].idxmax(),:]
    list_paddy_1_SCPJO_final_4 = list_paddy_1_SCPJO_final_2.loc[list_paddy_1_SCPJO_final_3["SCPJO"].idxmax(),:]


#出力
    with pd.ExcelWriter(paddy_1_excel) as writer:
        list_paddy_1_ORY2_final_4.to_excel(writer, sheet_name = "ORY2", index=False)
        list_paddy_1_ECHOR_final_4.to_excel(writer, sheet_name = "ECHOR", index=False)
        list_paddy_1_MOOVP_final_4.to_excel(writer, sheet_name = "MOOVP", index=False)
        list_paddy_1_SCPJO_final_4.to_excel(writer, sheet_name = "SCPJO", index=False)
except:
    pass    





#%%
#水田2次("ORY2","ECHOR","MOOVP","SCPJO")

try:
    list_paddy_2.columns = paddy_2_column
    list_paddy_2 = list_paddy_2[list_paddy_2["compound"] != "H"]
    list_paddy_2 = list_paddy_2[list_paddy_2["compound"] != "化合物"]
    list_paddy_2 = list_paddy_2.fillna(method="ffill")
    list_paddy_2 = list_paddy_2.replace("-",0)

#型の変換(表に-が入っていると数字の部分が文字列と認識されてしまっているため
    list2 = ["concentration","ORY2","ECHOR","MOOVP","SCPJO"]
    for i in list2:
        list_paddy_2[i] = list_paddy_2[i].astype(float)
        list_paddy_2.dtypes

#草種ごとに分類
    list_paddy_2_ORY2 = list_paddy_2.loc[:,["compound","concentration","ORY2"]]
    list_paddy_2_ECHOR = list_paddy_2.loc[:,["compound","concentration","ECHOR"]]
    list_paddy_2_MOOVP = list_paddy_2.loc[:,["compound","concentration","MOOVP"]]
    list_paddy_2_SCPJO = list_paddy_2.loc[:,["compound","concentration","SCPJO"]]

#任意の活性以下のものを省く
    list_paddy_2_ORY2 = list_paddy_2_ORY2[list_paddy_2_ORY2.ORY2 >= paddy_2]
    list_paddy_2_ECHOR = list_paddy_2_ECHOR[list_paddy_2_ECHOR.ECHOR >= paddy_2]
    list_paddy_2_MOOVP = list_paddy_2_MOOVP[list_paddy_2_MOOVP.MOOVP >= paddy_2]
    list_paddy_2_SCPJO = list_paddy_2_SCPJO[list_paddy_2_SCPJO.SCPJO >= paddy_2]

#順番を逆にする
    list_paddy_2_ORY2_2 = list_paddy_2_ORY2.sort_index(ascending=False)
    list_paddy_2_ECHOR_2 = list_paddy_2_ECHOR.sort_index(ascending=False)
    list_paddy_2_MOOVP_2 = list_paddy_2_MOOVP.sort_index(ascending=False)
    list_paddy_2_SCPJO_2 = list_paddy_2_SCPJO.sort_index(ascending=False)

#グループ化
    list_paddy_2_ORY2_3 = list_paddy_2_ORY2_2.groupby("compound")
    list_paddy_2_ECHOR_3 = list_paddy_2_ECHOR_2.groupby("compound")
    list_paddy_2_MOOVP_3 = list_paddy_2_MOOVP_2.groupby("compound")
    list_paddy_2_SCPJO_3 = list_paddy_2_SCPJO_2.groupby("compound")

#抽出
    list_paddy_2_ORY2_final = list_paddy_2_ORY2.loc[list_paddy_2_ORY2_3["ORY2"].idxmin(),:]
    list_paddy_2_ECHOR_final = list_paddy_2_ECHOR.loc[list_paddy_2_ECHOR_3["ECHOR"].idxmin(),:]
    list_paddy_2_MOOVP_final = list_paddy_2_MOOVP.loc[list_paddy_2_MOOVP_3["MOOVP"].idxmin(),:]
    list_paddy_2_SCPJO_final = list_paddy_2_SCPJO.loc[list_paddy_2_SCPJO_3["SCPJO"].idxmin(),:]


#活性低いものもの抽出1
    list_paddy_2_ORY2_low_1 = list_paddy_2.loc[:,["compound","concentration","ORY2"]]
    list_paddy_2_ECHOR_low_1 = list_paddy_2.loc[:,["compound","concentration","ECHOR"]]
    list_paddy_2_MOOVP_low_1 = list_paddy_2.loc[:,["compound","concentration","MOOVP"]]
    list_paddy_2_SCPJO_low_1 = list_paddy_2.loc[:,["compound","concentration","SCPJO"]]


#活性低いものもの抽出2
    list_paddy_2_ORY2_low_2 = list_paddy_2_ORY2_low_1[list_paddy_2_ORY2_low_1["concentration"] == 100]
    list_paddy_2_ECHOR_low_2 = list_paddy_2_ECHOR_low_1[list_paddy_2_ECHOR_low_1["concentration"] == 100]
    list_paddy_2_MOOVP_low_2 = list_paddy_2_MOOVP_low_1[list_paddy_2_MOOVP_low_1["concentration"] == 100]
    list_paddy_2_SCPJO_low_2 = list_paddy_2_SCPJO_low_1[list_paddy_2_SCPJO_low_1["concentration"] == 100]

#活性低いものもの抽出3
    list_paddy_2_ORY2_low_3 = list_paddy_2_ORY2_low_2[list_paddy_2_ORY2_low_2.ORY2 < paddy_2]
    list_paddy_2_ECHOR_low_3 = list_paddy_2_ECHOR_low_2[list_paddy_2_ECHOR_low_2.ECHOR < paddy_2]
    list_paddy_2_MOOVP_low_3 = list_paddy_2_MOOVP_low_2[list_paddy_2_MOOVP_low_2.MOOVP < paddy_2]
    list_paddy_2_SCPJO_low_3 = list_paddy_2_SCPJO_low_2[list_paddy_2_SCPJO_low_2.SCPJO < paddy_2]

#活性低いものもの抽出4
    list_paddy_2_ORY2_low_4 = list_paddy_2_ORY2_low_3.groupby("compound")
    list_paddy_2_ECHOR_low_4 = list_paddy_2_ECHOR_low_3.groupby("compound")
    list_paddy_2_MOOVP_low_4 = list_paddy_2_MOOVP_low_3.groupby("compound")
    list_paddy_2_SCPJO_low_4 = list_paddy_2_SCPJO_low_3.groupby("compound")

#活性低いものもの抽出5
    list_paddy_2_ORY2_low_final = list_paddy_2_ORY2_low_3.loc[list_paddy_2_ORY2_low_4["ORY2"].idxmax(),:]
    list_paddy_2_ECHOR_low_final = list_paddy_2_ECHOR_low_3.loc[list_paddy_2_ECHOR_low_4["ECHOR"].idxmax(),:]
    list_paddy_2_MOOVP_low_final = list_paddy_2_MOOVP_low_3.loc[list_paddy_2_MOOVP_low_4["MOOVP"].idxmax(),:]
    list_paddy_2_SCPJO_low_final = list_paddy_2_SCPJO_low_3.loc[list_paddy_2_SCPJO_low_4["SCPJO"].idxmax(),:]


#活性の高いものと低いものを結合
    list_paddy_2_ORY2_final_2 = pd.concat([list_paddy_2_ORY2_final, list_paddy_2_ORY2_low_final]) 
    list_paddy_2_ECHOR_final_2 = pd.concat([list_paddy_2_ECHOR_final, list_paddy_2_ECHOR_low_final]) 
    list_paddy_2_MOOVP_final_2 = pd.concat([list_paddy_2_MOOVP_final, list_paddy_2_MOOVP_low_final]) 
    list_paddy_2_SCPJO_final_2 = pd.concat([list_paddy_2_SCPJO_final, list_paddy_2_SCPJO_low_final]) 


#複数ある場合を除く
    list_paddy_2_ORY2_final_3 = list_paddy_2_ORY2_final_2.groupby("compound")
    list_paddy_2_ECHOR_final_3 = list_paddy_2_ECHOR_final_2.groupby("compound")
    list_paddy_2_MOOVP_final_3 = list_paddy_2_MOOVP_final_2.groupby("compound")
    list_paddy_2_SCPJO_final_3 = list_paddy_2_SCPJO_final_2.groupby("compound")


    list_paddy_2_ORY2_final_4 = list_paddy_2_ORY2_final_2.loc[list_paddy_2_ORY2_final_3["ORY2"].idxmax(),:]
    list_paddy_2_ECHOR_final_4 = list_paddy_2_ECHOR_final_2.loc[list_paddy_2_ECHOR_final_3["ECHOR"].idxmax(),:]
    list_paddy_2_MOOVP_final_4 = list_paddy_2_MOOVP_final_2.loc[list_paddy_2_MOOVP_final_3["MOOVP"].idxmax(),:]
    list_paddy_2_SCPJO_final_4 = list_paddy_2_SCPJO_final_2.loc[list_paddy_2_SCPJO_final_3["SCPJO"].idxmax(),:]


#出力
    with pd.ExcelWriter(paddy_2_excel) as writer:
        list_paddy_2_ORY2_final_4.to_excel(writer, sheet_name = "ORY2", index=False)
        list_paddy_2_ECHOR_final_4.to_excel(writer, sheet_name = "ECHOR", index=False)
        list_paddy_2_MOOVP_final_4.to_excel(writer, sheet_name = "MOOVP", index=False)
        list_paddy_2_SCPJO_final_4.to_excel(writer, sheet_name = "SCPJO", index=False)
except:
    pass    





#%%
#畑作_土壌1次("ECHCG","LOLMU","ABUTH","AMARE")
try:
    list_upland_soil.columns = upland_soil_column
    list_upland_soil = list_upland_soil[list_upland_soil["compound"] != "H"]
    list_upland_soil = list_upland_soil[list_upland_soil["compound"] != "化合物"]
    list_upland_soil = list_upland_soil.fillna(method="ffill")
    list_upland_soil = list_upland_soil.replace("-",0)

#型の変換(表に-が入っていると数字の部分が文字列と認識されてしまっているため
    list4 = ["concentration","ECHCG","LOLMU","ABUTH","AMARE"]
    for i in list4:
        list_upland_soil[i] = list_upland_soil[i].astype(float)
        list_upland_soil.dtypes

#草種ごとに分類
    list_upland_soil_ECHCG = list_upland_soil.loc[:,["compound","concentration","ECHCG"]]
    list_upland_soil_LOLMU = list_upland_soil.loc[:,["compound","concentration","LOLMU"]]
    list_upland_soil_ABUTH = list_upland_soil.loc[:,["compound","concentration","ABUTH"]]
    list_upland_soil_AMARE = list_upland_soil.loc[:,["compound","concentration","AMARE"]]

#任意の活性以下のものを省く
    list_upland_soil_ECHCG = list_upland_soil_ECHCG[list_upland_soil_ECHCG.ECHCG >= upland_soil]
    list_upland_soil_LOLMU = list_upland_soil_LOLMU[list_upland_soil_LOLMU.LOLMU >= upland_soil]
    list_upland_soil_ABUTH = list_upland_soil_ABUTH[list_upland_soil_ABUTH.ABUTH >= upland_soil]
    list_upland_soil_AMARE = list_upland_soil_AMARE[list_upland_soil_AMARE.AMARE >= upland_soil]

#順番を逆にする
    list_upland_soil_ECHCG_2 = list_upland_soil_ECHCG.sort_index(ascending=False)
    list_upland_soil_LOLMU_2 = list_upland_soil_LOLMU.sort_index(ascending=False)
    list_upland_soil_ABUTH_2 = list_upland_soil_ABUTH.sort_index(ascending=False)
    list_upland_soil_AMARE_2 = list_upland_soil_AMARE.sort_index(ascending=False)

#グループ化
    list_upland_soil_ECHCG_3 = list_upland_soil_ECHCG_2.groupby("compound")
    list_upland_soil_LOLMU_3 = list_upland_soil_LOLMU_2.groupby("compound")
    list_upland_soil_ABUTH_3 = list_upland_soil_ABUTH_2.groupby("compound")
    list_upland_soil_AMARE_3 = list_upland_soil_AMARE_2.groupby("compound")


#抽出
    list_upland_soil_ECHCG_final = list_upland_soil_ECHCG.loc[list_upland_soil_ECHCG_3["ECHCG"].idxmin(),:]
    list_upland_soil_LOLMU_final = list_upland_soil_LOLMU.loc[list_upland_soil_LOLMU_3["LOLMU"].idxmin(),:]
    list_upland_soil_ABUTH_final = list_upland_soil_ABUTH.loc[list_upland_soil_ABUTH_3["ABUTH"].idxmin(),:]
    list_upland_soil_AMARE_final = list_upland_soil_AMARE.loc[list_upland_soil_AMARE_3["AMARE"].idxmin(),:]



#活性低いものもの抽出1
    list_upland_soil_ECHCG_low_1 = list_upland_soil.loc[:,["compound","concentration","ECHCG"]]
    list_upland_soil_LOLMU_low_1 = list_upland_soil.loc[:,["compound","concentration","LOLMU"]]
    list_upland_soil_ABUTH_low_1 = list_upland_soil.loc[:,["compound","concentration","ABUTH"]]
    list_upland_soil_AMARE_low_1 = list_upland_soil.loc[:,["compound","concentration","AMARE"]]


#活性低いものもの抽出2
    list_upland_soil_ECHCG_low_2 = list_upland_soil_ECHCG_low_1[list_upland_soil_ECHCG_low_1["concentration"] == 100]
    list_upland_soil_LOLMU_low_2 = list_upland_soil_LOLMU_low_1[list_upland_soil_LOLMU_low_1["concentration"] == 100]
    list_upland_soil_ABUTH_low_2 = list_upland_soil_ABUTH_low_1[list_upland_soil_ABUTH_low_1["concentration"] == 100]
    list_upland_soil_AMARE_low_2 = list_upland_soil_AMARE_low_1[list_upland_soil_AMARE_low_1["concentration"] == 100]

#活性低いものもの抽出3
    list_upland_soil_ECHCG_low_3 = list_upland_soil_ECHCG_low_2[list_upland_soil_ECHCG_low_2.ECHCG < upland_soil]
    list_upland_soil_LOLMU_low_3 = list_upland_soil_LOLMU_low_2[list_upland_soil_LOLMU_low_2.LOLMU < upland_soil]
    list_upland_soil_ABUTH_low_3 = list_upland_soil_ABUTH_low_2[list_upland_soil_ABUTH_low_2.ABUTH < upland_soil]
    list_upland_soil_AMARE_low_3 = list_upland_soil_AMARE_low_2[list_upland_soil_AMARE_low_2.AMARE < upland_soil]

#活性低いものもの抽出4
    list_upland_soil_ECHCG_low_4 = list_upland_soil_ECHCG_low_3.groupby("compound")
    list_upland_soil_LOLMU_low_4 = list_upland_soil_LOLMU_low_3.groupby("compound")
    list_upland_soil_ABUTH_low_4 = list_upland_soil_ABUTH_low_3.groupby("compound")
    list_upland_soil_AMARE_low_4 = list_upland_soil_AMARE_low_3.groupby("compound")

#活性低いものもの抽出5
    list_upland_soil_ECHCG_low_final = list_upland_soil_ECHCG_low_3.loc[list_upland_soil_ECHCG_low_4["ECHCG"].idxmax(),:]
    list_upland_soil_LOLMU_low_final = list_upland_soil_LOLMU_low_3.loc[list_upland_soil_LOLMU_low_4["LOLMU"].idxmax(),:]
    list_upland_soil_ABUTH_low_final = list_upland_soil_ABUTH_low_3.loc[list_upland_soil_ABUTH_low_4["ABUTH"].idxmax(),:]
    list_upland_soil_AMARE_low_final = list_upland_soil_AMARE_low_3.loc[list_upland_soil_AMARE_low_4["AMARE"].idxmax(),:]


#活性の高いものと低いものを結合
    list_upland_soil_ECHCG_final_2 = pd.concat([list_upland_soil_ECHCG_final, list_upland_soil_ECHCG_low_final]) 
    list_upland_soil_LOLMU_final_2 = pd.concat([list_upland_soil_LOLMU_final, list_upland_soil_LOLMU_final]) 
    list_upland_soil_ABUTH_final_2 = pd.concat([list_upland_soil_ABUTH_final, list_upland_soil_ABUTH_low_final]) 
    list_upland_soil_AMARE_final_2 = pd.concat([list_upland_soil_AMARE_final, list_upland_soil_AMARE_low_final]) 

#複数ある場合を除く
    list_upland_soil_ECHCG_final_3 = list_upland_soil_ECHCG_final_2.groupby("compound")
    list_upland_soil_LOLMU_final_3 = list_upland_soil_LOLMU_final_2.groupby("compound")
    list_upland_soil_ABUTH_final_3 = list_upland_soil_ABUTH_final_2.groupby("compound")
    list_upland_soil_AMARE_final_3 = list_upland_soil_AMARE_final_2.groupby("compound")



    list_upland_soil_ECHCG_final_4 = list_upland_soil_ECHCG_final_2.loc[list_upland_soil_ECHCG_final_3["ECHCG"].idxmax(),:]
    list_upland_soil_LOLMU_final_4 = list_upland_soil_LOLMU_final_2.loc[list_upland_soil_LOLMU_final_3["LOLMU"].idxmax(),:]
    list_upland_soil_ABUTH_final_4 = list_upland_soil_ABUTH_final_2.loc[list_upland_soil_ABUTH_final_3["ABUTH"].idxmax(),:]
    list_upland_soil_AMARE_final_4 = list_upland_soil_AMARE_final_2.loc[list_upland_soil_AMARE_final_3["AMARE"].idxmax(),:]

#出力
    with pd.ExcelWriter(upland_soil_excel) as writer:
        list_upland_soil_ECHCG_final_4.to_excel(writer, sheet_name = "ECHCG", index=False)
        list_upland_soil_LOLMU_final_4.to_excel(writer, sheet_name = "LOLMU", index=False)
        list_upland_soil_ABUTH_final_4.to_excel(writer, sheet_name = "ABUTH", index=False)
        list_upland_soil_AMARE_final_4.to_excel(writer, sheet_name = "AMARE", index=False)
except:
    pass    






#%%
#畑作_茎葉1次("ECHCG","LOLMU","ABUTH","AMARE")
try:
    list_upland_stem_leaf.columns =upland_stem_leaf_column
    list_upland_stem_leaf = list_upland_stem_leaf[list_upland_stem_leaf["compound"] != "H"]
    list_upland_stem_leaf = list_upland_stem_leaf[list_upland_stem_leaf["compound"] != "化合物"]
    list_upland_stem_leaf = list_upland_stem_leaf.fillna(method="ffill")
    list_upland_stem_leaf = list_upland_stem_leaf.replace("-",0)

#型の変換(表に-が入っていると数字の部分が文字列と認識されてしまっているため
    list5 = ["concentration","ECHCG","LOLMU","ABUTH","AMARE"]
    for i in list5:
        list_upland_stem_leaf[i] = list_upland_stem_leaf[i].astype(float)
        list_upland_stem_leaf.dtypes

#草種ごとに分類
    list_upland_stem_leaf_ECHCG = list_upland_stem_leaf.loc[:,["compound","concentration","ECHCG"]]
    list_upland_stem_leaf_LOLMU = list_upland_stem_leaf.loc[:,["compound","concentration","LOLMU"]]
    list_upland_stem_leaf_ABUTH = list_upland_stem_leaf.loc[:,["compound","concentration","ABUTH"]]
    list_upland_stem_leaf_AMARE = list_upland_stem_leaf.loc[:,["compound","concentration","AMARE"]]

#任意の活性以下のものを省く
    list_upland_stem_leaf_ECHCG = list_upland_stem_leaf_ECHCG[list_upland_stem_leaf_ECHCG.ECHCG >= upland_stem_leaf]
    list_upland_stem_leaf_LOLMU = list_upland_stem_leaf_LOLMU[list_upland_stem_leaf_LOLMU.LOLMU >= upland_stem_leaf]
    list_upland_stem_leaf_ABUTH = list_upland_stem_leaf_ABUTH[list_upland_stem_leaf_ABUTH.ABUTH >= upland_stem_leaf]
    list_upland_stem_leaf_AMARE = list_upland_stem_leaf_AMARE[list_upland_stem_leaf_AMARE.AMARE >= upland_stem_leaf]

#順番を逆にする
    list_upland_stem_leaf_ECHCG_2 = list_upland_stem_leaf_ECHCG.sort_index(ascending=False)
    list_upland_stem_leaf_LOLMU_2 = list_upland_stem_leaf_LOLMU.sort_index(ascending=False)
    list_upland_stem_leaf_ABUTH_2 = list_upland_stem_leaf_ABUTH.sort_index(ascending=False)
    list_upland_stem_leaf_AMARE_2 = list_upland_stem_leaf_AMARE.sort_index(ascending=False)

#グループ化
    list_upland_stem_leaf_ECHCG_3 = list_upland_stem_leaf_ECHCG_2.groupby("compound")
    list_upland_stem_leaf_LOLMU_3 = list_upland_stem_leaf_LOLMU_2.groupby("compound")
    list_upland_stem_leaf_ABUTH_3 = list_upland_stem_leaf_ABUTH_2.groupby("compound")
    list_upland_stem_leaf_AMARE_3 = list_upland_stem_leaf_AMARE_2.groupby("compound")


#抽出
    list_upland_stem_leaf_ECHCG_final = list_upland_stem_leaf_ECHCG.loc[list_upland_stem_leaf_ECHCG_3["ECHCG"].idxmin(),:]
    list_upland_stem_leaf_LOLMU_final = list_upland_stem_leaf_LOLMU.loc[list_upland_stem_leaf_LOLMU_3["LOLMU"].idxmin(),:]
    list_upland_stem_leaf_ABUTH_final = list_upland_stem_leaf_ABUTH.loc[list_upland_stem_leaf_ABUTH_3["ABUTH"].idxmin(),:]
    list_upland_stem_leaf_AMARE_final = list_upland_stem_leaf_AMARE.loc[list_upland_stem_leaf_AMARE_3["AMARE"].idxmin(),:]



#活性低いものもの抽出1
    list_upland_stem_leaf_ECHCG_low_1 = list_upland_stem_leaf.loc[:,["compound","concentration","ECHCG"]]
    list_upland_stem_leaf_LOLMU_low_1 = list_upland_stem_leaf.loc[:,["compound","concentration","LOLMU"]]
    list_upland_stem_leaf_ABUTH_low_1 = list_upland_stem_leaf.loc[:,["compound","concentration","ABUTH"]]
    list_upland_stem_leaf_AMARE_low_1 = list_upland_stem_leaf.loc[:,["compound","concentration","AMARE"]]


#活性低いものもの抽出2
    list_upland_stem_leaf_ECHCG_low_2 = list_upland_stem_leaf_ECHCG_low_1[list_upland_stem_leaf_ECHCG_low_1["concentration"] == 100]
    list_upland_stem_leaf_LOLMU_low_2 = list_upland_stem_leaf_LOLMU_low_1[list_upland_stem_leaf_LOLMU_low_1["concentration"] == 100]
    list_upland_stem_leaf_ABUTH_low_2 = list_upland_stem_leaf_ABUTH_low_1[list_upland_stem_leaf_ABUTH_low_1["concentration"] == 100]
    list_upland_stem_leaf_AMARE_low_2 = list_upland_stem_leaf_AMARE_low_1[list_upland_stem_leaf_AMARE_low_1["concentration"] == 100]

#活性低いものもの抽出3
    list_upland_stem_leaf_ECHCG_low_3 = list_upland_stem_leaf_ECHCG_low_2[list_upland_stem_leaf_ECHCG_low_2.ECHCG < upland_stem_leaf]
    list_upland_stem_leaf_LOLMU_low_3 = list_upland_stem_leaf_LOLMU_low_2[list_upland_stem_leaf_LOLMU_low_2.LOLMU < upland_stem_leaf]
    list_upland_stem_leaf_ABUTH_low_3 = list_upland_stem_leaf_ABUTH_low_2[list_upland_stem_leaf_ABUTH_low_2.ABUTH < upland_stem_leaf]
    list_upland_stem_leaf_AMARE_low_3 = list_upland_stem_leaf_AMARE_low_2[list_upland_stem_leaf_AMARE_low_2.AMARE < upland_stem_leaf]
    
#活性低いものもの抽出4
    list_upland_stem_leaf_ECHCG_low_4 = list_upland_stem_leaf_ECHCG_low_3.groupby("compound")
    list_upland_stem_leaf_LOLMU_low_4 = list_upland_stem_leaf_LOLMU_low_3.groupby("compound")
    list_upland_stem_leaf_ABUTH_low_4= list_upland_stem_leaf_ABUTH_low_3.groupby("compound")
    list_upland_stem_leaf_AMARE_low_4 = list_upland_stem_leaf_AMARE_low_3.groupby("compound")

#活性低いものもの抽出5
    list_upland_stem_leaf_ECHCG_low_final = list_upland_stem_leaf_ECHCG_low_3.loc[list_upland_stem_leaf_ECHCG_low_4["ECHCG"].idxmax(),:]
    list_upland_stem_leaf_LOLMU_low_final = list_upland_stem_leaf_LOLMU_low_3.loc[list_upland_stem_leaf_LOLMU_low_4["LOLMU"].idxmax(),:]
    list_upland_stem_leaf_ABUTH_low_final = list_upland_stem_leaf_ABUTH_low_3.loc[list_upland_stem_leaf_ABUTH_low_4["ABUTH"].idxmax(),:]
    list_upland_stem_leaf_AMARE_low_final = list_upland_stem_leaf_AMARE_low_3.loc[list_upland_stem_leaf_AMARE_low_4["AMARE"].idxmax(),:]


#活性の高いものと低いものを結合
    list_upland_stem_leaf_ECHCG_final_2 = pd.concat([list_upland_stem_leaf_ECHCG_final, list_upland_stem_leaf_ECHCG_low_final]) 
    list_upland_stem_leaf_LOLMU_final_2 = pd.concat([list_upland_stem_leaf_LOLMU_final, list_upland_stem_leaf_LOLMU_final]) 
    list_upland_stem_leaf_ABUTH_final_2 = pd.concat([list_upland_stem_leaf_ABUTH_final, list_upland_stem_leaf_ABUTH_low_final]) 
    list_upland_stem_leaf_AMARE_final_2 = pd.concat([list_upland_stem_leaf_AMARE_final, list_upland_stem_leaf_AMARE_low_final]) 

#複数ある場合を除く
    list_upland_stem_leaf_ECHCG_final_3 = list_upland_stem_leaf_ECHCG_final_2.groupby("compound")
    list_upland_stem_leaf_LOLMU_final_3 = list_upland_stem_leaf_LOLMU_final_2.groupby("compound")
    list_upland_stem_leaf_ABUTH_final_3 = list_upland_stem_leaf_ABUTH_final_2.groupby("compound")
    list_upland_stem_leaf_AMARE_final_3 = list_upland_stem_leaf_AMARE_final_2.groupby("compound")

    list_upland_stem_leaf_ECHCG_final_4 = list_upland_stem_leaf_ECHCG_final_2.loc[list_upland_stem_leaf_ECHCG_final_3["ECHCG"].idxmax(),:]
    list_upland_stem_leaf_LOLMU_final_4 = list_upland_stem_leaf_LOLMU_final_2.loc[list_upland_stem_leaf_LOLMU_final_3["LOLMU"].idxmax(),:]
    list_upland_stem_leaf_ABUTH_final_4 = list_upland_stem_leaf_ABUTH_final_2.loc[list_upland_stem_leaf_ABUTH_final_3["ABUTH"].idxmax(),:]
    list_upland_stem_leaf_AMARE_final_4 = list_upland_stem_leaf_AMARE_final_2.loc[list_upland_stem_leaf_AMARE_final_3["AMARE"].idxmax(),:]


#出力
    with pd.ExcelWriter(upland_stem_leaf_excel) as writer:
        list_upland_stem_leaf_ECHCG_final_4.to_excel(writer, sheet_name = "ECHCG", index=False)
        list_upland_stem_leaf_LOLMU_final_4.to_excel(writer, sheet_name = "LOLMU", index=False)
        list_upland_stem_leaf_ABUTH_final_4.to_excel(writer, sheet_name = "ABUTH", index=False)
        list_upland_stem_leaf_AMARE_final_4.to_excel(writer, sheet_name = "AMARE", index=False)
except:
    pass    
  
    
    
    
    
    
    
    
#%%  
#畑作_土壌2次("TRZAX","BRSNN","AVEFA","ALOMY","ZEAMX","GLXMA","ORYSA","ECHCG","POLLN","CHEAL","ABUTH","IPOLA","AMBEL")
try:
    list_upland_soil_2.columns = upland_soil_2_column
    list_upland_soil_2 = list_upland_soil_2[list_upland_soil_2["compound"] != "H"]
    list_upland_soil_2 = list_upland_soil_2[list_upland_soil_2["compound"] != "化合物"]
    list_upland_soil_2 = list_upland_soil_2.fillna(method="ffill")
    list_upland_soil_2= list_upland_soil_2.replace("-",0)

#型の変換(表に-が入っていると数字の部分が文字列と認識されてしまっているため
    list6 = ["concentration","TRZAX","BRSNN","AVEFA","ALOMY","ZEAMX","GLXMA","ORYSA","ECHCG","POLLN","CHEAL","ABUTH","IPOLA","AMBEL"]
    for i in list6:
        list_upland_soil_2[i] = list_upland_soil_2[i].astype(float)
        list_upland_soil_2.dtypes

#草種ごとに分類
    list_upland_soil_2_TRZAX = list_upland_soil_2.loc[:,["compound","concentration","TRZAX"]]
    list_upland_soil_2_BRSNN = list_upland_soil_2.loc[:,["compound","concentration","BRSNN"]]
    list_upland_soil_2_AVEFA = list_upland_soil_2.loc[:,["compound","concentration","AVEFA"]]
    list_upland_soil_2_ALOMY = list_upland_soil_2.loc[:,["compound","concentration","ALOMY"]]
    list_upland_soil_2_ZEAMX = list_upland_soil_2.loc[:,["compound","concentration","ZEAMX"]]
    list_upland_soil_2_GLXMA = list_upland_soil_2.loc[:,["compound","concentration","GLXMA"]]
    list_upland_soil_2_ORYSA = list_upland_soil_2.loc[:,["compound","concentration","ORYSA"]]
    list_upland_soil_2_ECHCG = list_upland_soil_2.loc[:,["compound","concentration","ECHCG"]]
    list_upland_soil_2_POLLN = list_upland_soil_2.loc[:,["compound","concentration","POLLN"]]
    list_upland_soil_2_CHEAL = list_upland_soil_2.loc[:,["compound","concentration","CHEAL"]]
    list_upland_soil_2_ABUTH = list_upland_soil_2.loc[:,["compound","concentration","ABUTH"]]
    list_upland_soil_2_IPOLA = list_upland_soil_2.loc[:,["compound","concentration","IPOLA"]]
    list_upland_soil_2_AMBEL = list_upland_soil_2.loc[:,["compound","concentration","AMBEL"]]



#任意の活性以下のものを省く
    list_upland_soil_2_TRZAX = list_upland_soil_2_TRZAX[list_upland_soil_2_TRZAX.TRZAX >= upland_soil_2]
    list_upland_soil_2_BRSNN = list_upland_soil_2_BRSNN[list_upland_soil_2_BRSNN.BRSNN >= upland_soil_2]
    list_upland_soil_2_AVEFA = list_upland_soil_2_AVEFA[list_upland_soil_2_AVEFA.AVEFA >= upland_soil_2]
    list_upland_soil_2_ALOMY = list_upland_soil_2_ALOMY[list_upland_soil_2_ALOMY.ALOMY >= upland_soil_2]
    list_upland_soil_2_ZEAMX = list_upland_soil_2_ZEAMX[list_upland_soil_2_ZEAMX.ZEAMX >= upland_soil_2]
    list_upland_soil_2_GLXMA = list_upland_soil_2_GLXMA[list_upland_soil_2_GLXMA.GLXMA >= upland_soil_2]
    list_upland_soil_2_ORYSA = list_upland_soil_2_ORYSA[list_upland_soil_2_ORYSA.ORYSA >= upland_soil_2]
    list_upland_soil_2_ECHCG = list_upland_soil_2_ECHCG[list_upland_soil_2_ECHCG.ECHCG >= upland_soil_2]
    list_upland_soil_2_POLLN = list_upland_soil_2_POLLN[list_upland_soil_2_POLLN.POLLN >= upland_soil_2]
    list_upland_soil_2_CHEAL = list_upland_soil_2_CHEAL[list_upland_soil_2_CHEAL.CHEAL >= upland_soil_2]
    list_upland_soil_2_ABUTH = list_upland_soil_2_ABUTH[list_upland_soil_2_ABUTH.ABUTH >= upland_soil_2]
    list_upland_soil_2_IPOLA = list_upland_soil_2_IPOLA[list_upland_soil_2_IPOLA.IPOLA >= upland_soil_2]
    list_upland_soil_2_AMBEL = list_upland_soil_2_AMBEL[list_upland_soil_2_AMBEL.AMBEL >= upland_soil_2]

#順番を逆にする
    list_upland_soil_2_TRZAX_2 = list_upland_soil_2_TRZAX.sort_index(ascending=False)
    list_upland_soil_2_BRSNN_2 = list_upland_soil_2_BRSNN.sort_index(ascending=False)
    list_upland_soil_2_AVEFA_2 = list_upland_soil_2_AVEFA.sort_index(ascending=False)
    list_upland_soil_2_ALOMY_2 = list_upland_soil_2_ALOMY.sort_index(ascending=False)
    list_upland_soil_2_ZEAMX_2 = list_upland_soil_2_ZEAMX.sort_index(ascending=False)
    list_upland_soil_2_GLXMA_2 = list_upland_soil_2_GLXMA.sort_index(ascending=False)
    list_upland_soil_2_ORYSA_2 = list_upland_soil_2_ORYSA.sort_index(ascending=False)
    list_upland_soil_2_ECHCG_2 = list_upland_soil_2_ECHCG.sort_index(ascending=False)
    list_upland_soil_2_POLLN_2 = list_upland_soil_2_POLLN.sort_index(ascending=False)
    list_upland_soil_2_CHEAL_2 = list_upland_soil_2_CHEAL.sort_index(ascending=False)
    list_upland_soil_2_ABUTH_2 = list_upland_soil_2_ABUTH.sort_index(ascending=False)
    list_upland_soil_2_IPOLA_2 = list_upland_soil_2_IPOLA.sort_index(ascending=False)
    list_upland_soil_2_AMBEL_2 = list_upland_soil_2_AMBEL.sort_index(ascending=False)
    

#グループ化
    list_upland_soil_2_TRZAX_3 = list_upland_soil_2_TRZAX_2.groupby("compound")
    list_upland_soil_2_BRSNN_3 = list_upland_soil_2_BRSNN_2.groupby("compound")
    list_upland_soil_2_AVEFA_3 = list_upland_soil_2_AVEFA_2.groupby("compound")
    list_upland_soil_2_ALOMY_3 = list_upland_soil_2_ALOMY_2.groupby("compound")
    list_upland_soil_2_ZEAMX_3 = list_upland_soil_2_ZEAMX_2.groupby("compound")
    list_upland_soil_2_GLXMA_3 = list_upland_soil_2_GLXMA_2.groupby("compound")
    list_upland_soil_2_ORYSA_3 = list_upland_soil_2_ORYSA_2.groupby("compound")
    list_upland_soil_2_ECHCG_3 = list_upland_soil_2_ECHCG_2.groupby("compound")
    list_upland_soil_2_POLLN_3 = list_upland_soil_2_POLLN_2.groupby("compound")
    list_upland_soil_2_CHEAL_3 = list_upland_soil_2_CHEAL_2.groupby("compound")
    list_upland_soil_2_ABUTH_3 = list_upland_soil_2_ABUTH_2.groupby("compound")
    list_upland_soil_2_IPOLA_3 = list_upland_soil_2_IPOLA_2.groupby("compound")
    list_upland_soil_2_AMBEL_3 = list_upland_soil_2_AMBEL_2.groupby("compound")
    


#抽出
    list_upland_soil_2_TRZAX_final = list_upland_soil_2_TRZAX.loc[list_upland_soil_2_TRZAX_3["TRZAX"].idxmin(),:]
    list_upland_soil_2_BRSNN_final = list_upland_soil_2_BRSNN.loc[list_upland_soil_2_BRSNN_3["BRSNN"].idxmin(),:]
    list_upland_soil_2_AVEFA_final = list_upland_soil_2_AVEFA.loc[list_upland_soil_2_AVEFA_3["AVEFA"].idxmin(),:]
    list_upland_soil_2_ALOMY_final = list_upland_soil_2_ALOMY.loc[list_upland_soil_2_ALOMY_3["ALOMY"].idxmin(),:]
    list_upland_soil_2_ZEAMX_final = list_upland_soil_2_ZEAMX.loc[list_upland_soil_2_ZEAMX_3["ZEAMX"].idxmin(),:]
    list_upland_soil_2_GLXMA_final = list_upland_soil_2_GLXMA.loc[list_upland_soil_2_GLXMA_3["GLXMA"].idxmin(),:]
    list_upland_soil_2_ORYSA_final = list_upland_soil_2_ORYSA.loc[list_upland_soil_2_ORYSA_3["ORYSA"].idxmin(),:]
    list_upland_soil_2_ECHCG_final = list_upland_soil_2_ECHCG.loc[list_upland_soil_2_ECHCG_3["ECHCG"].idxmin(),:]
    list_upland_soil_2_POLLN_final = list_upland_soil_2_POLLN.loc[list_upland_soil_2_POLLN_3["POLLN"].idxmin(),:]
    list_upland_soil_2_CHEAL_final = list_upland_soil_2_CHEAL.loc[list_upland_soil_2_CHEAL_3["CHEAL"].idxmin(),:]
    list_upland_soil_2_ABUTH_final = list_upland_soil_2_ABUTH.loc[list_upland_soil_2_ABUTH_3["ABUTH"].idxmin(),:]
    list_upland_soil_2_IPOLA_final = list_upland_soil_2_IPOLA.loc[list_upland_soil_2_IPOLA_3["IPOLA"].idxmin(),:]
    list_upland_soil_2_AMBEL_final = list_upland_soil_2_AMBEL.loc[list_upland_soil_2_AMBEL_3["AMBEL"].idxmin(),:]


#活性低いものもの抽出1
    list_upland_soil_2_TRZAX_low_1 = list_upland_soil_2.loc[:,["compound","concentration","TRZAX"]]
    list_upland_soil_2_BRSNN_low_1 = list_upland_soil_2.loc[:,["compound","concentration","BRSNN"]]
    list_upland_soil_2_AVEFA_low_1 = list_upland_soil_2.loc[:,["compound","concentration","AVEFA"]]
    list_upland_soil_2_ALOMY_low_1 = list_upland_soil_2.loc[:,["compound","concentration","ALOMY"]]
    list_upland_soil_2_ZEAMX_low_1 = list_upland_soil_2.loc[:,["compound","concentration","ZEAMX"]]
    list_upland_soil_2_GLXMA_low_1 = list_upland_soil_2.loc[:,["compound","concentration","GLXMA"]]
    list_upland_soil_2_ORYSA_low_1 = list_upland_soil_2.loc[:,["compound","concentration","ORYSA"]]
    list_upland_soil_2_ECHCG_low_1 = list_upland_soil_2.loc[:,["compound","concentration","ECHCG"]]
    list_upland_soil_2_POLLN_low_1 = list_upland_soil_2.loc[:,["compound","concentration","POLLN"]]
    list_upland_soil_2_CHEAL_low_1 = list_upland_soil_2.loc[:,["compound","concentration","CHEAL"]]
    list_upland_soil_2_ABUTH_low_1 = list_upland_soil_2.loc[:,["compound","concentration","ABUTH"]]
    list_upland_soil_2_IPOLA_low_1 = list_upland_soil_2.loc[:,["compound","concentration","IPOLA"]]
    list_upland_soil_2_AMBEL_low_1 = list_upland_soil_2.loc[:,["compound","concentration","AMBEL"]]



#活性低いものもの抽出2
    list_upland_soil_2_TRZAX_low_2 = list_upland_soil_2_TRZAX_low_1[list_upland_soil_2_TRZAX_low_1["concentration"] == 100]
    list_upland_soil_2_BRSNN_low_2 = list_upland_soil_2_BRSNN_low_1[list_upland_soil_2_BRSNN_low_1["concentration"] == 100]
    list_upland_soil_2_AVEFA_low_2 = list_upland_soil_2_AVEFA_low_1[list_upland_soil_2_AVEFA_low_1["concentration"] == 100]
    list_upland_soil_2_ALOMY_low_2 = list_upland_soil_2_ALOMY_low_1[list_upland_soil_2_ALOMY_low_1["concentration"] == 100]
    list_upland_soil_2_ZEAMX_low_2 = list_upland_soil_2_ZEAMX_low_1[list_upland_soil_2_ZEAMX_low_1["concentration"] == 100]
    list_upland_soil_2_GLXMA_low_2 = list_upland_soil_2_GLXMA_low_1[list_upland_soil_2_GLXMA_low_1["concentration"] == 100]
    list_upland_soil_2_ORYSA_low_2 = list_upland_soil_2_ORYSA_low_1[list_upland_soil_2_ORYSA_low_1["concentration"] == 100]
    list_upland_soil_2_ECHCG_low_2 = list_upland_soil_2_ECHCG_low_1[list_upland_soil_2_ECHCG_low_1["concentration"] == 100]
    list_upland_soil_2_POLLN_low_2 = list_upland_soil_2_POLLN_low_1[list_upland_soil_2_POLLN_low_1["concentration"] == 100]
    list_upland_soil_2_CHEAL_low_2 = list_upland_soil_2_CHEAL_low_1[list_upland_soil_2_CHEAL_low_1["concentration"] == 100]
    list_upland_soil_2_ABUTH_low_2 = list_upland_soil_2_ABUTH_low_1[list_upland_soil_2_ABUTH_low_1["concentration"] == 100]
    list_upland_soil_2_IPOLA_low_2 = list_upland_soil_2_IPOLA_low_1[list_upland_soil_2_IPOLA_low_1["concentration"] == 100]
    list_upland_soil_2_AMBEL_low_2 = list_upland_soil_2_AMBEL_low_1[list_upland_soil_2_AMBEL_low_1["concentration"] == 100]



#活性低いものもの抽出3
    list_upland_soil_2_TRZAX_low_3 = list_upland_soil_2_TRZAX_low_2[list_upland_soil_2_TRZAX_low_2.TRZAX < upland_soil_2]
    list_upland_soil_2_BRSNN_low_3 = list_upland_soil_2_BRSNN_low_2[list_upland_soil_2_BRSNN_low_2.BRSNN < upland_soil_2]
    list_upland_soil_2_AVEFA_low_3 = list_upland_soil_2_AVEFA_low_2[list_upland_soil_2_AVEFA_low_2.AVEFA < upland_soil_2]
    list_upland_soil_2_ALOMY_low_3 = list_upland_soil_2_ALOMY_low_2[list_upland_soil_2_ALOMY_low_2.ALOMY < upland_soil_2]
    list_upland_soil_2_ZEAMX_low_3 = list_upland_soil_2_ZEAMX_low_2[list_upland_soil_2_ZEAMX_low_2.ZEAMX < upland_soil_2]
    list_upland_soil_2_GLXMA_low_3 = list_upland_soil_2_GLXMA_low_2[list_upland_soil_2_GLXMA_low_2.GLXMA < upland_soil_2]
    list_upland_soil_2_ORYSA_low_3 = list_upland_soil_2_ORYSA_low_2[list_upland_soil_2_ORYSA_low_2.ORYSA < upland_soil_2]
    list_upland_soil_2_ECHCG_low_3 = list_upland_soil_2_ECHCG_low_2[list_upland_soil_2_ECHCG_low_2.ECHCG < upland_soil_2]
    list_upland_soil_2_POLLN_low_3 = list_upland_soil_2_POLLN_low_2[list_upland_soil_2_POLLN_low_2.POLLN < upland_soil_2]
    list_upland_soil_2_CHEAL_low_3 = list_upland_soil_2_CHEAL_low_2[list_upland_soil_2_CHEAL_low_2.CHEAL < upland_soil_2]
    list_upland_soil_2_ABUTH_low_3 = list_upland_soil_2_ABUTH_low_2[list_upland_soil_2_ABUTH_low_2.ABUTH < upland_soil_2]
    list_upland_soil_2_IPOLA_low_3 = list_upland_soil_2_IPOLA_low_2[list_upland_soil_2_IPOLA_low_2.IPOLA < upland_soil_2]
    list_upland_soil_2_AMBEL_low_3 = list_upland_soil_2_AMBEL_low_2[list_upland_soil_2_AMBEL_low_2.AMBEL < upland_soil_2]

#活性低いものもの抽出4
    list_upland_soil_2_TRZAX_low_4 = list_upland_soil_2_TRZAX_low_3.groupby("compound")
    list_upland_soil_2_BRSNN_low_4 = list_upland_soil_2_BRSNN_low_3.groupby("compound")
    list_upland_soil_2_AVEFA_low_4 = list_upland_soil_2_AVEFA_low_3.groupby("compound")
    list_upland_soil_2_ALOMY_low_4 = list_upland_soil_2_ALOMY_low_3.groupby("compound")
    list_upland_soil_2_ZEAMX_low_4 = list_upland_soil_2_ZEAMX_low_3.groupby("compound")
    list_upland_soil_2_GLXMA_low_4 = list_upland_soil_2_GLXMA_low_3.groupby("compound")
    list_upland_soil_2_ORYSA_low_4 = list_upland_soil_2_ORYSA_low_3.groupby("compound")
    list_upland_soil_2_ECHCG_low_4 = list_upland_soil_2_ECHCG_low_3.groupby("compound")
    list_upland_soil_2_POLLN_low_4 = list_upland_soil_2_POLLN_low_3.groupby("compound")
    list_upland_soil_2_CHEAL_low_4 = list_upland_soil_2_CHEAL_low_3.groupby("compound")
    list_upland_soil_2_ABUTH_low_4 = list_upland_soil_2_ABUTH_low_3.groupby("compound")
    list_upland_soil_2_IPOLA_low_4 = list_upland_soil_2_IPOLA_low_3.groupby("compound")
    list_upland_soil_2_AMBEL_low_4 = list_upland_soil_2_AMBEL_low_3.groupby("compound")

#活性低いものもの抽出5
    list_upland_soil_2_TRZAX_low_final = list_upland_soil_2_TRZAX_low_3.loc[list_upland_soil_2_TRZAX_low_4["TRZAX"].idxmax(),:]
    list_upland_soil_2_BRSNN_low_final = list_upland_soil_2_BRSNN_low_3.loc[list_upland_soil_2_BRSNN_low_4["BRSNN"].idxmax(),:]
    list_upland_soil_2_AVEFA_low_final = list_upland_soil_2_AVEFA_low_3.loc[list_upland_soil_2_AVEFA_low_4["AVEFA"].idxmax(),:]
    list_upland_soil_2_ALOMY_low_final = list_upland_soil_2_ALOMY_low_3.loc[list_upland_soil_2_ALOMY_low_4["ALOMY"].idxmax(),:]
    list_upland_soil_2_ZEAMX_low_final = list_upland_soil_2_ZEAMX_low_3.loc[list_upland_soil_2_ZEAMX_low_4["ZEAMX"].idxmax(),:]
    list_upland_soil_2_GLXMA_low_final = list_upland_soil_2_GLXMA_low_3.loc[list_upland_soil_2_GLXMA_low_4["GLXMA"].idxmax(),:]
    list_upland_soil_2_ORYSA_low_final = list_upland_soil_2_ORYSA_low_3.loc[list_upland_soil_2_ORYSA_low_4["ORYSA"].idxmax(),:]
    list_upland_soil_2_ECHCG_low_final = list_upland_soil_2_ECHCG_low_3.loc[list_upland_soil_2_ECHCG_low_4["ECHCG"].idxmax(),:]
    list_upland_soil_2_POLLN_low_final = list_upland_soil_2_POLLN_low_3.loc[list_upland_soil_2_POLLN_low_4["POLLN"].idxmax(),:]
    list_upland_soil_2_CHEAL_low_final = list_upland_soil_2_CHEAL_low_3.loc[list_upland_soil_2_CHEAL_low_4["CHEAL"].idxmax(),:]
    list_upland_soil_2_ABUTH_low_final = list_upland_soil_2_ABUTH_low_3.loc[list_upland_soil_2_ABUTH_low_4["ABUTH"].idxmax(),:]
    list_upland_soil_2_IPOLA_low_final = list_upland_soil_2_IPOLA_low_3.loc[list_upland_soil_2_IPOLA_low_4["IPOLA"].idxmax(),:]
    list_upland_soil_2_AMBEL_low_final = list_upland_soil_2_AMBEL_low_3.loc[list_upland_soil_2_AMBEL_low_4["AMBEL"].idxmax(),:]
    

#活性の高いものと低いものを結合
    list_upland_soil_2_TRZAX_final_2 = pd.concat([list_upland_soil_2_TRZAX_final, list_upland_soil_2_TRZAX_low_final])
    list_upland_soil_2_BRSNN_final_2 = pd.concat([list_upland_soil_2_BRSNN_final, list_upland_soil_2_BRSNN_low_final])
    list_upland_soil_2_AVEFA_final_2 = pd.concat([list_upland_soil_2_AVEFA_final, list_upland_soil_2_AVEFA_low_final])
    list_upland_soil_2_ALOMY_final_2 = pd.concat([list_upland_soil_2_ALOMY_final, list_upland_soil_2_ALOMY_low_final])
    list_upland_soil_2_ZEAMX_final_2 = pd.concat([list_upland_soil_2_ZEAMX_final, list_upland_soil_2_ZEAMX_low_final])
    list_upland_soil_2_GLXMA_final_2 = pd.concat([list_upland_soil_2_GLXMA_final, list_upland_soil_2_GLXMA_low_final])
    list_upland_soil_2_ORYSA_final_2 = pd.concat([list_upland_soil_2_ORYSA_final, list_upland_soil_2_ORYSA_low_final])
    list_upland_soil_2_ECHCG_final_2 = pd.concat([list_upland_soil_2_ECHCG_final, list_upland_soil_2_ECHCG_low_final])
    list_upland_soil_2_POLLN_final_2 = pd.concat([list_upland_soil_2_POLLN_final, list_upland_soil_2_POLLN_low_final])
    list_upland_soil_2_CHEAL_final_2 = pd.concat([list_upland_soil_2_CHEAL_final, list_upland_soil_2_CHEAL_low_final])
    list_upland_soil_2_ABUTH_final_2 = pd.concat([list_upland_soil_2_ABUTH_final, list_upland_soil_2_ABUTH_low_final])
    list_upland_soil_2_IPOLA_final_2 = pd.concat([list_upland_soil_2_IPOLA_final, list_upland_soil_2_IPOLA_low_final])
    list_upland_soil_2_AMBEL_final_2 = pd.concat([list_upland_soil_2_AMBEL_final, list_upland_soil_2_AMBEL_low_final])

#複数ある場合を除く
    list_upland_soil_2_TRZAX_final_3 = list_upland_soil_2_TRZAX_final_2.groupby("compound")
    list_upland_soil_2_BRSNN_final_3 = list_upland_soil_2_BRSNN_final_2.groupby("compound")
    list_upland_soil_2_AVEFA_final_3 = list_upland_soil_2_AVEFA_final_2.groupby("compound")
    list_upland_soil_2_ALOMY_final_3 = list_upland_soil_2_ALOMY_final_2.groupby("compound")
    list_upland_soil_2_ZEAMX_final_3 = list_upland_soil_2_ZEAMX_final_2.groupby("compound")
    list_upland_soil_2_GLXMA_final_3 = list_upland_soil_2_GLXMA_final_2.groupby("compound")
    list_upland_soil_2_ORYSA_final_3 = list_upland_soil_2_ORYSA_final_2.groupby("compound")
    list_upland_soil_2_ECHCG_final_3 = list_upland_soil_2_ECHCG_final_2.groupby("compound")
    list_upland_soil_2_POLLN_final_3 = list_upland_soil_2_POLLN_final_2.groupby("compound")
    list_upland_soil_2_CHEAL_final_3 = list_upland_soil_2_CHEAL_final_2.groupby("compound")
    list_upland_soil_2_ABUTH_final_3 = list_upland_soil_2_ABUTH_final_2.groupby("compound")
    list_upland_soil_2_IPOLA_final_3 = list_upland_soil_2_IPOLA_final_2.groupby("compound")
    list_upland_soil_2_AMBEL_final_3 = list_upland_soil_2_AMBEL_final_2.groupby("compound")


    list_upland_soil_2_TRZAX_final_4 = list_upland_soil_2_TRZAX_final_2.loc[list_upland_soil_2_TRZAX_final_3["TRZAX"].idxmax(),:]
    list_upland_soil_2_BRSNN_final_4 = list_upland_soil_2_BRSNN_final_2.loc[list_upland_soil_2_BRSNN_final_3["BRSNN"].idxmax(),:]
    list_upland_soil_2_AVEFA_final_4 = list_upland_soil_2_AVEFA_final_2.loc[list_upland_soil_2_AVEFA_final_3["AVEFA"].idxmax(),:]
    list_upland_soil_2_ALOMY_final_4 = list_upland_soil_2_ALOMY_final_2.loc[list_upland_soil_2_ALOMY_final_3["ALOMY"].idxmax(),:]
    list_upland_soil_2_ZEAMX_final_4 = list_upland_soil_2_ZEAMX_final_2.loc[list_upland_soil_2_ZEAMX_final_3["ZEAMX"].idxmax(),:]
    list_upland_soil_2_GLXMA_final_4 = list_upland_soil_2_GLXMA_final_2.loc[list_upland_soil_2_GLXMA_final_3["GLXMA"].idxmax(),:]
    list_upland_soil_2_ORYSA_final_4 = list_upland_soil_2_ORYSA_final_2.loc[list_upland_soil_2_ORYSA_final_3["ORYSA"].idxmax(),:]
    list_upland_soil_2_ECHCG_final_4 = list_upland_soil_2_ECHCG_final_2.loc[list_upland_soil_2_ECHCG_final_3["ECHCG"].idxmax(),:]
    list_upland_soil_2_POLLN_final_4 = list_upland_soil_2_POLLN_final_2.loc[list_upland_soil_2_POLLN_final_3["POLLN"].idxmax(),:]
    list_upland_soil_2_CHEAL_final_4 = list_upland_soil_2_CHEAL_final_2.loc[list_upland_soil_2_CHEAL_final_3["CHEAL"].idxmax(),:]
    list_upland_soil_2_ABUTH_final_4 = list_upland_soil_2_ABUTH_final_2.loc[list_upland_soil_2_ABUTH_final_3["ABUTH"].idxmax(),:]
    list_upland_soil_2_IPOLA_final_4 = list_upland_soil_2_IPOLA_final_2.loc[list_upland_soil_2_IPOLA_final_3["IPOLA"].idxmax(),:]
    list_upland_soil_2_AMBEL_final_4 = list_upland_soil_2_AMBEL_final_2.loc[list_upland_soil_2_AMBEL_final_3["AMBEL"].idxmax(),:]


#出力
    with pd.ExcelWriter(upland_soil_2_excel) as writer:
        list_upland_soil_2_TRZAX_final_4.to_excel(writer, sheet_name = "TRZAX", index=False)
        list_upland_soil_2_BRSNN_final_4.to_excel(writer, sheet_name = "BRSNN", index=False)
        list_upland_soil_2_AVEFA_final_4.to_excel(writer, sheet_name = "AVEFA", index=False)
        list_upland_soil_2_ALOMY_final_4.to_excel(writer, sheet_name = "ALOMY", index=False)
        list_upland_soil_2_ZEAMX_final_4.to_excel(writer, sheet_name = "ZEAMX", index=False)
        list_upland_soil_2_GLXMA_final_4.to_excel(writer, sheet_name = "GLXMA", index=False)
        list_upland_soil_2_ORYSA_final_4.to_excel(writer, sheet_name = "ORYSA", index=False)
        list_upland_soil_2_ECHCG_final_4.to_excel(writer, sheet_name = "ECHCG", index=False)
        list_upland_soil_2_POLLN_final_4.to_excel(writer, sheet_name = "POLLN", index=False)
        list_upland_soil_2_CHEAL_final_4.to_excel(writer, sheet_name = "CHEAL", index=False)
        list_upland_soil_2_ABUTH_final_4.to_excel(writer, sheet_name = "ABUTH", index=False)
        list_upland_soil_2_IPOLA_final_4.to_excel(writer, sheet_name = "IPOLA", index=False)
        list_upland_soil_2_AMBEL_final_4.to_excel(writer, sheet_name = "AMBEL", index=False)
except:
    pass    
    





#%%
#畑作_茎葉2次("TRZAX","BRSNN","AVEFA","ALOMY","ZEAMX","GLXMA","ORYSA","ECHCG","POLLN","CHEAL","ABUTH","IPOLA","AMBEL")
try:
    list_upland_stem_leaf_2.columns =["compound","concentration","TRZAX","BRSNN","AVEFA","ALOMY","ZEAMX","GLXMA","ORYSA","ECHCG","POLLN","CHEAL","ABUTH","IPOLA","AMBEL"]
    list_upland_stem_leaf_2 = list_upland_stem_leaf_2[list_upland_stem_leaf_2["compound"] != "H"]
    list_upland_stem_leaf_2 = list_upland_stem_leaf_2[list_upland_stem_leaf_2["compound"] != "化合物"]
    list_upland_stem_leaf_2 = list_upland_stem_leaf_2.fillna(method="ffill")
    list_upland_stem_leaf_2= list_upland_stem_leaf_2.replace("-",0)

#型の変換(表に-が入っていると数字の部分が文字列と認識されてしまっているため
    list7 = ["concentration","TRZAX","BRSNN","AVEFA","ALOMY","ZEAMX","GLXMA","ORYSA","ECHCG","POLLN","CHEAL","ABUTH","IPOLA","AMBEL"]
    for i in list7:
        list_upland_stem_leaf_2[i] = list_upland_stem_leaf_2[i].astype(float)
        list_upland_stem_leaf_2.dtypes

#草種ごとに分類
    list_upland_stem_leaf_2_TRZAX = list_upland_stem_leaf_2.loc[:,["compound","concentration","TRZAX"]]
    list_upland_stem_leaf_2_BRSNN = list_upland_stem_leaf_2.loc[:,["compound","concentration","BRSNN"]]
    list_upland_stem_leaf_2_AVEFA = list_upland_stem_leaf_2.loc[:,["compound","concentration","AVEFA"]]
    list_upland_stem_leaf_2_ALOMY = list_upland_stem_leaf_2.loc[:,["compound","concentration","ALOMY"]]
    list_upland_stem_leaf_2_ZEAMX = list_upland_stem_leaf_2.loc[:,["compound","concentration","ZEAMX"]]
    list_upland_stem_leaf_2_GLXMA = list_upland_stem_leaf_2.loc[:,["compound","concentration","GLXMA"]]
    list_upland_stem_leaf_2_ORYSA = list_upland_stem_leaf_2.loc[:,["compound","concentration","ORYSA"]]
    list_upland_stem_leaf_2_ECHCG = list_upland_stem_leaf_2.loc[:,["compound","concentration","ECHCG"]]
    list_upland_stem_leaf_2_POLLN = list_upland_stem_leaf_2.loc[:,["compound","concentration","POLLN"]]
    list_upland_stem_leaf_2_CHEAL = list_upland_stem_leaf_2.loc[:,["compound","concentration","CHEAL"]]
    list_upland_stem_leaf_2_ABUTH = list_upland_stem_leaf_2.loc[:,["compound","concentration","ABUTH"]]
    list_upland_stem_leaf_2_IPOLA = list_upland_stem_leaf_2.loc[:,["compound","concentration","IPOLA"]]
    list_upland_stem_leaf_2_AMBEL = list_upland_stem_leaf_2.loc[:,["compound","concentration","AMBEL"]]



#任意の活性以下のものを省く
    list_upland_stem_leaf_2_TRZAX = list_upland_stem_leaf_2_TRZAX[list_upland_stem_leaf_2_TRZAX.TRZAX >= upland_stem_leaf_2]
    list_upland_stem_leaf_2_BRSNN = list_upland_stem_leaf_2_BRSNN[list_upland_stem_leaf_2_BRSNN.BRSNN >= upland_stem_leaf_2]
    list_upland_stem_leaf_2_AVEFA = list_upland_stem_leaf_2_AVEFA[list_upland_stem_leaf_2_AVEFA.AVEFA >= upland_stem_leaf_2]
    list_upland_stem_leaf_2_ALOMY = list_upland_stem_leaf_2_ALOMY[list_upland_stem_leaf_2_ALOMY.ALOMY >= upland_stem_leaf_2]
    list_upland_stem_leaf_2_ZEAMX = list_upland_stem_leaf_2_ZEAMX[list_upland_stem_leaf_2_ZEAMX.ZEAMX >= upland_stem_leaf_2]
    list_upland_stem_leaf_2_GLXMA = list_upland_stem_leaf_2_GLXMA[list_upland_stem_leaf_2_GLXMA.GLXMA >= upland_stem_leaf_2]
    list_upland_stem_leaf_2_ORYSA = list_upland_stem_leaf_2_ORYSA[list_upland_stem_leaf_2_ORYSA.ORYSA >= upland_stem_leaf_2]
    list_upland_stem_leaf_2_ECHCG = list_upland_stem_leaf_2_ECHCG[list_upland_stem_leaf_2_ECHCG.ECHCG >= upland_stem_leaf_2]
    list_upland_stem_leaf_2_POLLN = list_upland_stem_leaf_2_POLLN[list_upland_stem_leaf_2_POLLN.POLLN >= upland_stem_leaf_2]
    list_upland_stem_leaf_2_CHEAL = list_upland_stem_leaf_2_CHEAL[list_upland_stem_leaf_2_CHEAL.CHEAL >= upland_stem_leaf_2]
    list_upland_stem_leaf_2_ABUTH = list_upland_stem_leaf_2_ABUTH[list_upland_stem_leaf_2_ABUTH.ABUTH >= upland_stem_leaf_2]
    list_upland_stem_leaf_2_IPOLA = list_upland_stem_leaf_2_IPOLA[list_upland_stem_leaf_2_IPOLA.IPOLA >= upland_stem_leaf_2]
    list_upland_stem_leaf_2_AMBEL = list_upland_stem_leaf_2_AMBEL[list_upland_stem_leaf_2_AMBEL.AMBEL >= upland_stem_leaf_2]

#順番を逆にする
    list_upland_stem_leaf_2_TRZAX_2 = list_upland_stem_leaf_2_TRZAX.sort_index(ascending=False)
    list_upland_stem_leaf_2_BRSNN_2 = list_upland_stem_leaf_2_BRSNN.sort_index(ascending=False)
    list_upland_stem_leaf_2_AVEFA_2 = list_upland_stem_leaf_2_AVEFA.sort_index(ascending=False)
    list_upland_stem_leaf_2_ALOMY_2 = list_upland_stem_leaf_2_ALOMY.sort_index(ascending=False)
    list_upland_stem_leaf_2_ZEAMX_2 = list_upland_stem_leaf_2_ZEAMX.sort_index(ascending=False)
    list_upland_stem_leaf_2_GLXMA_2 = list_upland_stem_leaf_2_GLXMA.sort_index(ascending=False)
    list_upland_stem_leaf_2_ORYSA_2 = list_upland_stem_leaf_2_ORYSA.sort_index(ascending=False)
    list_upland_stem_leaf_2_ECHCG_2 = list_upland_stem_leaf_2_ECHCG.sort_index(ascending=False)
    list_upland_stem_leaf_2_POLLN_2 = list_upland_stem_leaf_2_POLLN.sort_index(ascending=False)
    list_upland_stem_leaf_2_CHEAL_2 = list_upland_stem_leaf_2_CHEAL.sort_index(ascending=False)
    list_upland_stem_leaf_2_ABUTH_2 = list_upland_stem_leaf_2_ABUTH.sort_index(ascending=False)
    list_upland_stem_leaf_2_IPOLA_2 = list_upland_stem_leaf_2_IPOLA.sort_index(ascending=False)
    list_upland_stem_leaf_2_AMBEL_2 = list_upland_stem_leaf_2_AMBEL.sort_index(ascending=False)


#グループ化
    list_upland_stem_leaf_2_TRZAX_3 = list_upland_stem_leaf_2_TRZAX_2.groupby("compound")
    list_upland_stem_leaf_2_BRSNN_3 = list_upland_stem_leaf_2_BRSNN_2.groupby("compound")
    list_upland_stem_leaf_2_AVEFA_3 = list_upland_stem_leaf_2_AVEFA_2.groupby("compound")
    list_upland_stem_leaf_2_ALOMY_3 = list_upland_stem_leaf_2_ALOMY_2.groupby("compound")
    list_upland_stem_leaf_2_ZEAMX_3 = list_upland_stem_leaf_2_ZEAMX_2.groupby("compound")
    list_upland_stem_leaf_2_GLXMA_3 = list_upland_stem_leaf_2_GLXMA_2.groupby("compound")
    list_upland_stem_leaf_2_ORYSA_3 = list_upland_stem_leaf_2_ORYSA_2.groupby("compound")
    list_upland_stem_leaf_2_ECHCG_3 = list_upland_stem_leaf_2_ECHCG_2.groupby("compound")
    list_upland_stem_leaf_2_POLLN_3 = list_upland_stem_leaf_2_POLLN_2.groupby("compound")
    list_upland_stem_leaf_2_CHEAL_3 = list_upland_stem_leaf_2_CHEAL_2.groupby("compound")
    list_upland_stem_leaf_2_ABUTH_3 = list_upland_stem_leaf_2_ABUTH_2.groupby("compound")
    list_upland_stem_leaf_2_IPOLA_3 = list_upland_stem_leaf_2_IPOLA_2.groupby("compound")
    list_upland_stem_leaf_2_AMBEL_3 = list_upland_stem_leaf_2_AMBEL_2.groupby("compound")



#抽出
    list_upland_stem_leaf_2_TRZAX_final = list_upland_stem_leaf_2_TRZAX.loc[list_upland_stem_leaf_2_TRZAX_3["TRZAX"].idxmin(),:]
    list_upland_stem_leaf_2_BRSNN_final = list_upland_stem_leaf_2_BRSNN.loc[list_upland_stem_leaf_2_BRSNN_3["BRSNN"].idxmin(),:]
    list_upland_stem_leaf_2_AVEFA_final = list_upland_stem_leaf_2_AVEFA.loc[list_upland_stem_leaf_2_AVEFA_3["AVEFA"].idxmin(),:]
    list_upland_stem_leaf_2_ALOMY_final = list_upland_stem_leaf_2_ALOMY.loc[list_upland_stem_leaf_2_ALOMY_3["ALOMY"].idxmin(),:]
    list_upland_stem_leaf_2_ZEAMX_final = list_upland_stem_leaf_2_ZEAMX.loc[list_upland_stem_leaf_2_ZEAMX_3["ZEAMX"].idxmin(),:]
    list_upland_stem_leaf_2_GLXMA_final = list_upland_stem_leaf_2_GLXMA.loc[list_upland_stem_leaf_2_GLXMA_3["GLXMA"].idxmin(),:]
    list_upland_stem_leaf_2_ORYSA_final = list_upland_stem_leaf_2_ORYSA.loc[list_upland_stem_leaf_2_ORYSA_3["ORYSA"].idxmin(),:]
    list_upland_stem_leaf_2_ECHCG_final = list_upland_stem_leaf_2_ECHCG.loc[list_upland_stem_leaf_2_ECHCG_3["ECHCG"].idxmin(),:]
    list_upland_stem_leaf_2_POLLN_final = list_upland_stem_leaf_2_POLLN.loc[list_upland_stem_leaf_2_POLLN_3["POLLN"].idxmin(),:]
    list_upland_stem_leaf_2_CHEAL_final = list_upland_stem_leaf_2_CHEAL.loc[list_upland_stem_leaf_2_CHEAL_3["CHEAL"].idxmin(),:]
    list_upland_stem_leaf_2_ABUTH_final = list_upland_stem_leaf_2_ABUTH.loc[list_upland_stem_leaf_2_ABUTH_3["ABUTH"].idxmin(),:]
    list_upland_stem_leaf_2_IPOLA_final = list_upland_stem_leaf_2_IPOLA.loc[list_upland_stem_leaf_2_IPOLA_3["IPOLA"].idxmin(),:]
    list_upland_stem_leaf_2_AMBEL_final = list_upland_stem_leaf_2_AMBEL.loc[list_upland_stem_leaf_2_AMBEL_3["AMBEL"].idxmin(),:]
    


#活性低いものもの抽出1
    list_upland_stem_leaf_2_TRZAX_low_1 = list_upland_stem_leaf_2.loc[:,["compound","concentration","TRZAX"]]
    list_upland_stem_leaf_2_BRSNN_low_1 = list_upland_stem_leaf_2.loc[:,["compound","concentration","BRSNN"]]
    list_upland_stem_leaf_2_AVEFA_low_1 = list_upland_stem_leaf_2.loc[:,["compound","concentration","AVEFA"]]
    list_upland_stem_leaf_2_ALOMY_low_1 = list_upland_stem_leaf_2.loc[:,["compound","concentration","ALOMY"]]
    list_upland_stem_leaf_2_ZEAMX_low_1 = list_upland_stem_leaf_2.loc[:,["compound","concentration","ZEAMX"]]
    list_upland_stem_leaf_2_GLXMA_low_1 = list_upland_stem_leaf_2.loc[:,["compound","concentration","GLXMA"]]
    list_upland_stem_leaf_2_ORYSA_low_1 = list_upland_stem_leaf_2.loc[:,["compound","concentration","ORYSA"]]
    list_upland_stem_leaf_2_ECHCG_low_1 = list_upland_stem_leaf_2.loc[:,["compound","concentration","ECHCG"]]
    list_upland_stem_leaf_2_POLLN_low_1 = list_upland_stem_leaf_2.loc[:,["compound","concentration","POLLN"]]
    list_upland_stem_leaf_2_CHEAL_low_1 = list_upland_stem_leaf_2.loc[:,["compound","concentration","CHEAL"]]
    list_upland_stem_leaf_2_ABUTH_low_1 = list_upland_stem_leaf_2.loc[:,["compound","concentration","ABUTH"]]
    list_upland_stem_leaf_2_IPOLA_low_1 = list_upland_stem_leaf_2.loc[:,["compound","concentration","IPOLA"]]
    list_upland_stem_leaf_2_AMBEL_low_1 = list_upland_stem_leaf_2.loc[:,["compound","concentration","AMBEL"]]



#活性低いものもの抽出2
    list_upland_stem_leaf_2_TRZAX_low_2 = list_upland_stem_leaf_2_TRZAX_low_1[list_upland_stem_leaf_2_TRZAX_low_1["concentration"] == 100]
    list_upland_stem_leaf_2_BRSNN_low_2 = list_upland_stem_leaf_2_BRSNN_low_1[list_upland_stem_leaf_2_BRSNN_low_1["concentration"] == 100]
    list_upland_stem_leaf_2_AVEFA_low_2 = list_upland_stem_leaf_2_AVEFA_low_1[list_upland_stem_leaf_2_AVEFA_low_1["concentration"] == 100]
    list_upland_stem_leaf_2_ALOMY_low_2 = list_upland_stem_leaf_2_ALOMY_low_1[list_upland_stem_leaf_2_ALOMY_low_1["concentration"] == 100]
    list_upland_stem_leaf_2_ZEAMX_low_2 = list_upland_stem_leaf_2_ZEAMX_low_1[list_upland_stem_leaf_2_ZEAMX_low_1["concentration"] == 100]
    list_upland_stem_leaf_2_GLXMA_low_2 = list_upland_stem_leaf_2_GLXMA_low_1[list_upland_stem_leaf_2_GLXMA_low_1["concentration"] == 100]
    list_upland_stem_leaf_2_ORYSA_low_2 = list_upland_stem_leaf_2_ORYSA_low_1[list_upland_stem_leaf_2_ORYSA_low_1["concentration"] == 100]
    list_upland_stem_leaf_2_ECHCG_low_2 = list_upland_stem_leaf_2_ECHCG_low_1[list_upland_stem_leaf_2_ECHCG_low_1["concentration"] == 100]
    list_upland_stem_leaf_2_POLLN_low_2 = list_upland_stem_leaf_2_POLLN_low_1[list_upland_stem_leaf_2_POLLN_low_1["concentration"] == 100]
    list_upland_stem_leaf_2_CHEAL_low_2 = list_upland_stem_leaf_2_CHEAL_low_1[list_upland_stem_leaf_2_CHEAL_low_1["concentration"] == 100]
    list_upland_stem_leaf_2_ABUTH_low_2 = list_upland_stem_leaf_2_ABUTH_low_1[list_upland_stem_leaf_2_ABUTH_low_1["concentration"] == 100]
    list_upland_stem_leaf_2_IPOLA_low_2 = list_upland_stem_leaf_2_IPOLA_low_1[list_upland_stem_leaf_2_IPOLA_low_1["concentration"] == 100]
    list_upland_stem_leaf_2_AMBEL_low_2 = list_upland_stem_leaf_2_AMBEL_low_1[list_upland_stem_leaf_2_AMBEL_low_1["concentration"] == 100]



#活性低いものもの抽出3
    list_upland_stem_leaf_2_TRZAX_low_3 = list_upland_stem_leaf_2_TRZAX_low_2[list_upland_stem_leaf_2_TRZAX_low_2.TRZAX < upland_stem_leaf_2]
    list_upland_stem_leaf_2_BRSNN_low_3 = list_upland_stem_leaf_2_BRSNN_low_2[list_upland_stem_leaf_2_BRSNN_low_2.BRSNN < upland_stem_leaf_2]
    list_upland_stem_leaf_2_AVEFA_low_3 = list_upland_stem_leaf_2_AVEFA_low_2[list_upland_stem_leaf_2_AVEFA_low_2.AVEFA < upland_stem_leaf_2]
    list_upland_stem_leaf_2_ALOMY_low_3 = list_upland_stem_leaf_2_ALOMY_low_2[list_upland_stem_leaf_2_ALOMY_low_2.ALOMY < upland_stem_leaf_2]
    list_upland_stem_leaf_2_ZEAMX_low_3 = list_upland_stem_leaf_2_ZEAMX_low_2[list_upland_stem_leaf_2_ZEAMX_low_2.ZEAMX < upland_stem_leaf_2]
    list_upland_stem_leaf_2_GLXMA_low_3 = list_upland_stem_leaf_2_GLXMA_low_2[list_upland_stem_leaf_2_GLXMA_low_2.GLXMA < upland_stem_leaf_2]
    list_upland_stem_leaf_2_ORYSA_low_3 = list_upland_stem_leaf_2_ORYSA_low_2[list_upland_stem_leaf_2_ORYSA_low_2.ORYSA < upland_stem_leaf_2]
    list_upland_stem_leaf_2_ECHCG_low_3 = list_upland_stem_leaf_2_ECHCG_low_2[list_upland_stem_leaf_2_ECHCG_low_2.ECHCG < upland_stem_leaf_2]
    list_upland_stem_leaf_2_POLLN_low_3 = list_upland_stem_leaf_2_POLLN_low_2[list_upland_stem_leaf_2_POLLN_low_2.POLLN < upland_stem_leaf_2]
    list_upland_stem_leaf_2_CHEAL_low_3 = list_upland_stem_leaf_2_CHEAL_low_2[list_upland_stem_leaf_2_CHEAL_low_2.CHEAL < upland_stem_leaf_2]
    list_upland_stem_leaf_2_ABUTH_low_3 = list_upland_stem_leaf_2_ABUTH_low_2[list_upland_stem_leaf_2_ABUTH_low_2.ABUTH < upland_stem_leaf_2]
    list_upland_stem_leaf_2_IPOLA_low_3 = list_upland_stem_leaf_2_IPOLA_low_2[list_upland_stem_leaf_2_IPOLA_low_2.IPOLA < upland_stem_leaf_2]
    list_upland_stem_leaf_2_AMBEL_low_3 = list_upland_stem_leaf_2_AMBEL_low_2[list_upland_stem_leaf_2_AMBEL_low_2.AMBEL < upland_stem_leaf_2]

#活性低いものもの抽出4
    list_upland_stem_leaf_2_TRZAX_low_4 = list_upland_stem_leaf_2_TRZAX_low_3.groupby("compound")
    list_upland_stem_leaf_2_BRSNN_low_4 = list_upland_stem_leaf_2_BRSNN_low_3.groupby("compound")
    list_upland_stem_leaf_2_AVEFA_low_4 = list_upland_stem_leaf_2_AVEFA_low_3.groupby("compound")
    list_upland_stem_leaf_2_ALOMY_low_4 = list_upland_stem_leaf_2_ALOMY_low_3.groupby("compound")
    list_upland_stem_leaf_2_ZEAMX_low_4 = list_upland_stem_leaf_2_ZEAMX_low_3.groupby("compound")
    list_upland_stem_leaf_2_GLXMA_low_4 = list_upland_stem_leaf_2_GLXMA_low_3.groupby("compound")
    list_upland_stem_leaf_2_ORYSA_low_4 = list_upland_stem_leaf_2_ORYSA_low_3.groupby("compound")
    list_upland_stem_leaf_2_ECHCG_low_4 = list_upland_stem_leaf_2_ECHCG_low_3.groupby("compound")
    list_upland_stem_leaf_2_POLLN_low_4 = list_upland_stem_leaf_2_POLLN_low_3.groupby("compound")
    list_upland_stem_leaf_2_CHEAL_low_4 = list_upland_stem_leaf_2_CHEAL_low_3.groupby("compound")
    list_upland_stem_leaf_2_ABUTH_low_4 = list_upland_stem_leaf_2_ABUTH_low_3.groupby("compound")
    list_upland_stem_leaf_2_IPOLA_low_4 = list_upland_stem_leaf_2_IPOLA_low_3.groupby("compound")
    list_upland_stem_leaf_2_AMBEL_low_4 = list_upland_stem_leaf_2_AMBEL_low_3.groupby("compound")

#活性低いものもの抽出5
    list_upland_stem_leaf_2_TRZAX_low_final = list_upland_stem_leaf_2_TRZAX_low_3.loc[list_upland_stem_leaf_2_TRZAX_low_4["TRZAX"].idxmax(),:]
    list_upland_stem_leaf_2_BRSNN_low_final = list_upland_stem_leaf_2_BRSNN_low_3.loc[list_upland_stem_leaf_2_BRSNN_low_4["BRSNN"].idxmax(),:]
    list_upland_stem_leaf_2_AVEFA_low_final = list_upland_stem_leaf_2_AVEFA_low_3.loc[list_upland_stem_leaf_2_AVEFA_low_4["AVEFA"].idxmax(),:]
    list_upland_stem_leaf_2_ALOMY_low_final = list_upland_stem_leaf_2_ALOMY_low_3.loc[list_upland_stem_leaf_2_ALOMY_low_4["ALOMY"].idxmax(),:]
    list_upland_stem_leaf_2_ZEAMX_low_final = list_upland_stem_leaf_2_ZEAMX_low_3.loc[list_upland_stem_leaf_2_ZEAMX_low_4["ZEAMX"].idxmax(),:]
    list_upland_stem_leaf_2_GLXMA_low_final = list_upland_stem_leaf_2_GLXMA_low_3.loc[list_upland_stem_leaf_2_GLXMA_low_4["GLXMA"].idxmax(),:]
    list_upland_stem_leaf_2_ORYSA_low_final = list_upland_stem_leaf_2_ORYSA_low_3.loc[list_upland_stem_leaf_2_ORYSA_low_4["ORYSA"].idxmax(),:]
    list_upland_stem_leaf_2_ECHCG_low_final = list_upland_stem_leaf_2_ECHCG_low_3.loc[list_upland_stem_leaf_2_ECHCG_low_4["ECHCG"].idxmax(),:]
    list_upland_stem_leaf_2_POLLN_low_final = list_upland_stem_leaf_2_POLLN_low_3.loc[list_upland_stem_leaf_2_POLLN_low_4["POLLN"].idxmax(),:]
    list_upland_stem_leaf_2_CHEAL_low_final = list_upland_stem_leaf_2_CHEAL_low_3.loc[list_upland_stem_leaf_2_CHEAL_low_4["CHEAL"].idxmax(),:]
    list_upland_stem_leaf_2_ABUTH_low_final = list_upland_stem_leaf_2_ABUTH_low_3.loc[list_upland_stem_leaf_2_ABUTH_low_4["ABUTH"].idxmax(),:]
    list_upland_stem_leaf_2_IPOLA_low_final = list_upland_stem_leaf_2_IPOLA_low_3.loc[list_upland_stem_leaf_2_IPOLA_low_4["IPOLA"].idxmax(),:]
    list_upland_stem_leaf_2_AMBEL_low_final = list_upland_stem_leaf_2_AMBEL_low_3.loc[list_upland_stem_leaf_2_AMBEL_low_4["AMBEL"].idxmax(),:]


#活性の高いものと低いものを結合
    list_upland_stem_leaf_2_TRZAX_final_2 = pd.concat([list_upland_stem_leaf_2_TRZAX_final, list_upland_stem_leaf_2_TRZAX_low_final])
    list_upland_stem_leaf_2_BRSNN_final_2 = pd.concat([list_upland_stem_leaf_2_BRSNN_final, list_upland_stem_leaf_2_BRSNN_low_final])
    list_upland_stem_leaf_2_AVEFA_final_2 = pd.concat([list_upland_stem_leaf_2_AVEFA_final, list_upland_stem_leaf_2_AVEFA_low_final])
    list_upland_stem_leaf_2_ALOMY_final_2 = pd.concat([list_upland_stem_leaf_2_ALOMY_final, list_upland_stem_leaf_2_ALOMY_low_final])
    list_upland_stem_leaf_2_ZEAMX_final_2 = pd.concat([list_upland_stem_leaf_2_ZEAMX_final, list_upland_stem_leaf_2_ZEAMX_low_final])
    list_upland_stem_leaf_2_GLXMA_final_2 = pd.concat([list_upland_stem_leaf_2_GLXMA_final, list_upland_stem_leaf_2_GLXMA_low_final])
    list_upland_stem_leaf_2_ORYSA_final_2 = pd.concat([list_upland_stem_leaf_2_ORYSA_final, list_upland_stem_leaf_2_ORYSA_low_final])
    list_upland_stem_leaf_2_ECHCG_final_2 = pd.concat([list_upland_stem_leaf_2_ECHCG_final, list_upland_stem_leaf_2_ECHCG_low_final])
    list_upland_stem_leaf_2_POLLN_final_2 = pd.concat([list_upland_stem_leaf_2_POLLN_final, list_upland_stem_leaf_2_POLLN_low_final])
    list_upland_stem_leaf_2_CHEAL_final_2 = pd.concat([list_upland_stem_leaf_2_CHEAL_final, list_upland_stem_leaf_2_CHEAL_low_final])
    list_upland_stem_leaf_2_ABUTH_final_2 = pd.concat([list_upland_stem_leaf_2_ABUTH_final, list_upland_stem_leaf_2_ABUTH_low_final])
    list_upland_stem_leaf_2_IPOLA_final_2 = pd.concat([list_upland_stem_leaf_2_IPOLA_final, list_upland_stem_leaf_2_IPOLA_low_final])
    list_upland_stem_leaf_2_AMBEL_final_2 = pd.concat([list_upland_stem_leaf_2_AMBEL_final, list_upland_stem_leaf_2_AMBEL_low_final])



#複数ある場合を除く
    list_upland_stem_leaf_2_TRZAX_final_3 = list_upland_stem_leaf_2_TRZAX_final_2.groupby("compound")
    list_upland_stem_leaf_2_BRSNN_final_3 = list_upland_stem_leaf_2_BRSNN_final_2.groupby("compound")
    list_upland_stem_leaf_2_AVEFA_final_3 = list_upland_stem_leaf_2_AVEFA_final_2.groupby("compound")
    list_upland_stem_leaf_2_ALOMY_final_3 = list_upland_stem_leaf_2_ALOMY_final_2.groupby("compound")
    list_upland_stem_leaf_2_ZEAMX_final_3 = list_upland_stem_leaf_2_ZEAMX_final_2.groupby("compound")
    list_upland_stem_leaf_2_GLXMA_final_3 = list_upland_stem_leaf_2_GLXMA_final_2.groupby("compound")
    list_upland_stem_leaf_2_ORYSA_final_3 = list_upland_stem_leaf_2_ORYSA_final_2.groupby("compound")
    list_upland_stem_leaf_2_ECHCG_final_3 = list_upland_stem_leaf_2_ECHCG_final_2.groupby("compound")
    list_upland_stem_leaf_2_POLLN_final_3 = list_upland_stem_leaf_2_POLLN_final_2.groupby("compound")
    list_upland_stem_leaf_2_CHEAL_final_3 = list_upland_stem_leaf_2_CHEAL_final_2.groupby("compound")
    list_upland_stem_leaf_2_ABUTH_final_3 = list_upland_stem_leaf_2_ABUTH_final_2.groupby("compound")
    list_upland_stem_leaf_2_IPOLA_final_3 = list_upland_stem_leaf_2_IPOLA_final_2.groupby("compound")
    list_upland_stem_leaf_2_AMBEL_final_3 = list_upland_stem_leaf_2_AMBEL_final_2.groupby("compound")
    

    list_upland_stem_leaf_2_TRZAX_final_4 = list_upland_stem_leaf_2_TRZAX_final_2.loc[list_upland_stem_leaf_2_TRZAX_final_3["TRZAX"].idxmax(),:]
    list_upland_stem_leaf_2_BRSNN_final_4 = list_upland_stem_leaf_2_BRSNN_final_2.loc[list_upland_stem_leaf_2_BRSNN_final_3["BRSNN"].idxmax(),:]
    list_upland_stem_leaf_2_AVEFA_final_4 = list_upland_stem_leaf_2_AVEFA_final_2.loc[list_upland_stem_leaf_2_AVEFA_final_3["AVEFA"].idxmax(),:]
    list_upland_stem_leaf_2_ALOMY_final_4 = list_upland_stem_leaf_2_ALOMY_final_2.loc[list_upland_stem_leaf_2_ALOMY_final_3["ALOMY"].idxmax(),:]
    list_upland_stem_leaf_2_ZEAMX_final_4 = list_upland_stem_leaf_2_ZEAMX_final_2.loc[list_upland_stem_leaf_2_ZEAMX_final_3["ZEAMX"].idxmax(),:]
    list_upland_stem_leaf_2_GLXMA_final_4 = list_upland_stem_leaf_2_GLXMA_final_2.loc[list_upland_stem_leaf_2_GLXMA_final_3["GLXMA"].idxmax(),:]
    list_upland_stem_leaf_2_ORYSA_final_4 = list_upland_stem_leaf_2_ORYSA_final_2.loc[list_upland_stem_leaf_2_ORYSA_final_3["ORYSA"].idxmax(),:]
    list_upland_stem_leaf_2_ECHCG_final_4 = list_upland_stem_leaf_2_ECHCG_final_2.loc[list_upland_stem_leaf_2_ECHCG_final_3["ECHCG"].idxmax(),:]
    list_upland_stem_leaf_2_POLLN_final_4 = list_upland_stem_leaf_2_POLLN_final_2.loc[list_upland_stem_leaf_2_POLLN_final_3["POLLN"].idxmax(),:]
    list_upland_stem_leaf_2_CHEAL_final_4 = list_upland_stem_leaf_2_CHEAL_final_2.loc[list_upland_stem_leaf_2_CHEAL_final_3["CHEAL"].idxmax(),:]
    list_upland_stem_leaf_2_ABUTH_final_4 = list_upland_stem_leaf_2_ABUTH_final_2.loc[list_upland_stem_leaf_2_ABUTH_final_3["ABUTH"].idxmax(),:]
    list_upland_stem_leaf_2_IPOLA_final_4 = list_upland_stem_leaf_2_IPOLA_final_2.loc[list_upland_stem_leaf_2_IPOLA_final_3["IPOLA"].idxmax(),:]
    list_upland_stem_leaf_2_AMBEL_final_4 = list_upland_stem_leaf_2_AMBEL_final_2.loc[list_upland_stem_leaf_2_AMBEL_final_3["AMBEL"].idxmax(),:]


#出力
    with pd.ExcelWriter(upland_stem_leaf_2_excel) as writer:
        list_upland_stem_leaf_2_TRZAX_final_4.to_excel(writer, sheet_name = "TRZAX", index=False)
        list_upland_stem_leaf_2_BRSNN_final_4.to_excel(writer, sheet_name = "BRSNN", index=False)
        list_upland_stem_leaf_2_AVEFA_final_4.to_excel(writer, sheet_name = "AVEFA", index=False)
        list_upland_stem_leaf_2_ALOMY_final_4.to_excel(writer, sheet_name = "ALOMY", index=False)
        list_upland_stem_leaf_2_ZEAMX_final_4.to_excel(writer, sheet_name = "ZEAMX", index=False)
        list_upland_stem_leaf_2_GLXMA_final_4.to_excel(writer, sheet_name = "GLXMA", index=False)
        list_upland_stem_leaf_2_ORYSA_final_4.to_excel(writer, sheet_name = "ORYSA", index=False)
        list_upland_stem_leaf_2_ECHCG_final_4.to_excel(writer, sheet_name = "ECHCG", index=False)
        list_upland_stem_leaf_2_POLLN_final_4.to_excel(writer, sheet_name = "POLLN", index=False)
        list_upland_stem_leaf_2_CHEAL_final_4.to_excel(writer, sheet_name = "CHEAL", index=False)
        list_upland_stem_leaf_2_ABUTH_final_4.to_excel(writer, sheet_name = "ABUTH", index=False)
        list_upland_stem_leaf_2_IPOLA_final_4.to_excel(writer, sheet_name = "IPOLA", index=False)
        list_upland_stem_leaf_2_AMBEL_final_4.to_excel(writer, sheet_name = "AMBEL", index=False)
except:
    pass    


