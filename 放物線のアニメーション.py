# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

from scipy.integrate import odeint

def equation(x,t,a,b):
    return a*t + b

import numpy as np
t = np.arange(0,5,0.1) #0から5まで0.1刻みで解く
x0 = 0 #初期条件
a = 1
b = 2

x = odeint(equation,x0,t,args=(a,b))


import numpy as np
from scipy.integrate import odeint
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation

g = 9.8 #重力加速度[m/s^2]
L = 10 #弾丸の発射地点とお猿さんとの直線距離[m]
theta = np.pi/4 #弾丸の発射角度[rad]
v0 = 10 #弾丸の初速[10m/s]
interval = 10 #計算上の時間間隔[ms]

t = np.arange(0,L/v0,interval/1000)
y0 = [[0,v0*np.sin(theta)], #弾丸の初期条件(y,v) 
      [L*np.sin(theta),0]]  #弾丸の初期条件(y,v)

def equation(y,t,g): 
    ret = [y[1],-1*g] #上の式の右辺そのまま
    return ret

y1 = odeint(equation,y0[0],t,args=(g,))
y2 = odeint(equation,y0[1],t,args=(g,))




