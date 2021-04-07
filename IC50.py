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

dd = pd.read_csv("IC_50.csv", header=None)

def IC50_fit(x,ic50,d):
    return 100/(1+(x/ic50)**d)

#array1 = np.array([[1,10,100,1000],[16,30,64,86],[18,30,62,80],[11,24,60,84],[18,24,61,81]])
#df = pd.DataFrame(array1)

concentration = np.array(dd.iloc[0])

ic_50 = []
for i in range(1,len(dd)):
    y = np.array(dd.iloc[i])
    param2, cov2 = curve_fit(IC50_fit, concentration, y)
    ic_50.append(param2[0])
    
    
ic_50_2 = [round(ic_50[n],2) for n in range(len(ic_50))]
        
dd_2 = dd.drop(dd.index[[0]])
dd_2["IC50"] = ic_50_2
index_name =["conc1","conc10","conc100","conc1000","IC50"]
dd_3 = dd_2.set_axis(index_name, axis=1)
