# -*- coding: utf-8 -*-
"""
Created on Wed Apr  7 20:00:13 2021

@author: takashi.shiozawa
"""


from scipy.optimize import curve_fit
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd
import os
import math

#パスの指定
os.chdir("C:/Users/taka3/Desktop")

"""フィッティングの基礎"""
#線形近似
list_linear_x = range(0,20,2)
array_error = np.random.normal(size=len(list_linear_x))
array_x = np.array(list_linear_x)
array_y = array_x + array_error

sns.pointplot(x=array_x, y=array_y, join=False)

def linear_fit(x,a,b):
    return a*x + b

param, cov = curve_fit(linear_fit, array_x, array_y)

array_y_fit = param[0] * array_x + param[1]


#%%
"""IC50の算出"""

def IC50_fit(x,ic50,d):
    return 100/(1+(x/ic50)**d)

"""
もう一つの式(返り値はlogなので注意!!)
def IC50_fit_2(x,log_ic50,d):
    return 100/(1+10**((log_ic50-np.log10(x))*d))
"""


"""手入力の場合
concentration = np.array([1,10,100,1000])
activity = np.array([16,30,64,86])

param, cov = curve_fit(IC50_fit, concentration, activity)

print(param[0])

list_y =[]
for num in activity:
    list_y.append(100/(1+(num/param[0])**param[1]))

plt.scatter(concentration,np.array(list_y))
ax=plt.gca()
ax.set_xscale("log")
plt.show()
"""

"""Excelからの場合"""
df = pd.read_csv("IC_50.csv", header=None)

#array1 = np.array([[1,10,100,1000],[16,30,64,86],[18,30,62,80],[11,24,60,84],[18,24,61,81]])
#df = pd.DataFrame(array1)

concentration = np.array(df.iloc[0])#活性濃度の抜出し
df_2 = df.drop(df.index[[0]])#活性値の抜出し

ic_50 = []#IC50を入れるからのリストの作成
for i in range(0,len(df_2)):
    y = np.array(df_2.iloc[i])
    param, cov = curve_fit(IC50_fit, concentration, y)
    ic_50.append(param[0])#空のリストに代入


ic_50_2 = [round(ic_50[n],2) for n in range(len(ic_50))]
        
df_2["IC50"] = ic_50_2
index_name =["conc1","conc10","conc100","conc1000","IC50"]
df_3 = df_2.set_axis(index_name, axis=1)

"""シグモイド曲線の検証
p = np.linspace(1,1000,100)
q = 100/(1+(p/50)**-5)
plt.scatter(p,q)
ax=plt.gca()
ax.set_xscale("log")
plt.show()
"""

