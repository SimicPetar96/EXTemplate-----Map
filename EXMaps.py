import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import Workbook

InputPath=r'' # INSERT INPUT PATH!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#
df=pd.read_excel(InputPath,sheet_name='Map Data')
df_new=df.iloc[3:,:]
df_new.columns=df.iloc[1,:]
#print(df_new)

#RPM
rpm_col=df_new['n_EM']
rpm_col=rpm_col.to_numpy()
arr_rpm=np.array(list(rpm_col))
rpm_col_set=set(arr_rpm)
arr=np.array(list(rpm_col_set))
rpm_col_sorted=np.sort(arr)

#TQ
tq_col=df_new['T_EM']
tq_col=tq_col.to_numpy()
tq_col=np.array(list(tq_col))
arr_tq=np.round(tq_col,decimals=0)
arr_set=set(arr_tq)
tq_col=np.array(list(arr_set))
tq_col=np.sort(tq_col)
tq_col=np.reshape(tq_col,(len(tq_col),1))
tq_col=tq_col[::-1]

scale=max(abs(max(tq_col)),abs(min(tq_col)))
x=scale % 5
scale=scale+5-x

torque_axis=np.arange(-scale,scale,5)
torque_axis=torque_axis[::-1]
#Temp
var=df_new['Inv:Tj_Max']
var=var.to_numpy()

mat=np.zeros((len(torque_axis),len(rpm_col_sorted)))
#arr_rpm
#arr_tq

for k in range(len(var)):
    temp_var=var[k]
    speed=arr_rpm[k]
    torque=arr_tq[k]
    mat_i =0
    mat_j=0

    for j in range(len(rpm_col_sorted)):
        if speed==rpm_col_sorted[j]:
            mat_j=j
    for i in range(len(torque_axis)):
        if abs(torque-torque_axis[i])<5:
            x=abs(torque-torque_axis[i])

            if i+1<len(torque_axis):
                y=abs(torque-torque_axis[i+1])
            else:
                y=999999

            min_=min(x,y)
            if min_==x:
                mat_i=i
            else:
                mat_i=i+1

    mat[mat_i][mat_j]=temp_var

#column to be added -> torque_axis
#row to be added    -> rpm_col_sorted

rows,columns = mat.shape

torque_axis=np.reshape(torque_axis,(rows,1))
mat=np.hstack((torque_axis,mat))

rpm_col_sorted=np.reshape(rpm_col_sorted,(1,columns))
rpm_col_sorted=np.append([0],rpm_col_sorted)
mat=np.vstack((rpm_col_sorted,mat))

df=pd.DataFrame(mat)
df.to_excel(excel_writer=r"") # INSERT OUTPUT PATH!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
