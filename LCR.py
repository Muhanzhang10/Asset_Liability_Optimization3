#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jan 25 16:57:28 2021

@author: zhangmuhan
"""

import Data_Input as dp

def lcr(after):
    net = dp.p_all + dp.d_all
    r1 = []
    r2 = []
    r3 = []
    for i in range(len(net)):
        for j in range(len(net[i].lcr)):
            r1.append(net[i].lcr[j].r1)
            r2.append(net[i].lcr[j].r2)
            r3.append(net[i].lcr[j].r3)
    #programme
    #print(r1)
    r4 = {}
    r4["一级资产"] = [126.51]
    r4["2A"] = []
    r4["2B"] = []
    r4["现金流出"] = [3572.6 , 754.55, 214.8]
    r4["现金流入"] = [ 2908.83, 59.73]
    
    D1 = 0
    D11 = 0
    D111 = 0
    D112 = 0
    D113 = 0
    D2 = 0
    D21 = 0
    D3 = 0
    D31 = 0
    D41 = 0
    D411 = 0
    D42 = 0
    D421 = 0
    
    count = 0
    for i in range(len(net)):
        for j in range(len(net[i].lcr)):
            if net[i].lcr[j].attribute == "D":
                D1 += after[i] * r1[count]
                D11 += after[i] * r1[count] * r3[count][1 - 1] 
                D111 += after[i] * r1[count] * r3[count][1 - 1] * r3[count][2 - 1]
                D112 += after[i] * r1[count] * r3[count][1 - 1] * r3[count][3 - 1]
                D113 += after[i] * r1[count] * r3[count][1 - 1] * r3[count][4 - 1]
            if net[i].lcr[j].attribute == "E":
                if net[i].AL == "A":
                    D2 += after[i] * r1[count]
                    D21 += after[i] * r1[count] * r3[count][0]
                else:
                    D2 -= after[i] * r1[count]
                    D21 -= after[i] * r1[count] * r3[count][0]
            if net[i].lcr[j].attribute == "F":
                if net[i].AL == "A":
                    D3 += after[i] * r1[count]
                    D31 += after[i] * r1[count] * r3[count][0]
                else:
                    D3 -= after[i] * r1[count]
                    D31 -= after[i] * r1[count] * r3[count][0]
            if net[i].lcr[j].attribute == "G":
                if net[i].AL == "A":
                    D41 += after[i] * r1[count]
                    D411 += after[i] * r1[count] * r3[count][0]
                else:
                    D41 -= after[i] * r1[count]
                    D411 -= after[i] * r1[count] * r3[count][0]
            if net[i].lcr[j].attribute == "H":
                if net[i].AL == "A":
                    D42 += after[i] * r1[count]
                    D421 += after[i] * r1[count] * r3[count][0]
                else:
                    D42 -= after[i] * r1[count]
                    D421 -= after[i] * r1[count] * r3[count][0]
            count += 1
    

    
    total = []
    total_before = []
    #after = []
    after_before = []
    
    
    for i in range(len(dp.p_all)):
        total.append(dp.p_all[i].balance_base[month])
        total_before.append(dp.p_all[i].balance_base[month - 1])
    for i in range(len(dp.d_all)):
        total.append(dp.d_all[i].balance_base[month])
        total_before.append(dp.d_all[i].balance_base[month - 1])
    
    if month == 1:
        after_before = total_before
    else:
        for i in range(len(dp.p_all)):
            after_before.append(dp.p_all[i].balance_after[month - 1])
            #after.append(dp.p_all[i].balance_after[month])
        for i in range(len(dp.d_all)):
            after_before.append(dp.d_all[i].balance_after[month - 1])        
            
            
    
     
    KMAP = 0
    KMA = 0
    
    for i in range(len(net)):
        if net[i].is_mortgage == True:
            KMA += total[i] - total_before[i]
            KMAP += after[i] - after_before[i]


    
    A1 = D111 - D21 + D2 + D3 + D41 + D42 - D11
    A2 = D112 - D31
    A3 = D113 - D421 - D411
    


    
    C11 = 0
    count = 0
    for i in range(len(net)):
        for j in range(len(net[i].lcr)):
            if net[i].AL == "A":
                if net[i].lcr[j].asset_type == "一级":
                    C11 += after[i] * r1[count] * r2[count]
            count += 1
    C11 += sum(r4["一级资产"])
    C11 = C11 - (KMAP - KMA)
    #print(C11)


    C12 = 0
    count = 0
    for i in range(len(net)):
        for j in range(len(net[i].lcr)):
            if net[i].AL == "A":
                if net[i].lcr[j].asset_type == "2A":
                    C12 += after[i] * r1[count] * r2[count]
            count += 1
                
  
    C13 = 0
    count = 0
    for i in range(len(net)):
        for j in range(len(net[i].lcr)):
            if net[i].AL == "A":
                if net[i].lcr[j].asset_type == "2B":    
                    C13 += after[i] * r1[count] * r2[count]
            count += 1
                
            
    C22 = 0
    count = 0
    for i in range(len(net)):
        for j in range(len(net[i].lcr)):
            if net[i].AL == "A":
                if net[i].lcr[j].asset_type == "其他":
                    C22 += after[i] * r1[count] * r2[count]
                    print(after[i], r1[count], r2[count])
            count += 1
            
    C22 += sum(r4["现金流入"])





    C21 = 0 
    count = 0
    for i in range(len(net)):
        for j in range(len(net[i].lcr)):
            if net[i].AL == "L":
                C21 += after[i] * r1[count] * r2[count] 
            count += 1    
            
    C21 += sum(r4["现金流出"])


    A11 = max(A1 + C11, 0)
    A21 = A2 + C12
    A31 = A3 + C13
    A32 = max(A31 - 15/85*(A11 + A21), A31 - 15/60*A11, 0)
    A4 = max(A21 + A31 - A32 - 2/3*A11, 0)

    C1 = C11 + C12 + C13 - A32 - A4
    C2 = C21 - min(C22, 0.75 * C21)

    
    return  C1 / C2


month = 1
total = []
for i in range(len(dp.p_all)):
    total.append(dp.p_all[i].balance_base[month])  
for i in range(len(dp.d_all)):
    total.append(dp.d_all[i].balance_base[month])

import xlrd
path1 =  "/Users/zhangmuhan/Desktop/实习/My_code/First/result_after.xls"  
ExcelFile3 = xlrd.open_workbook(path1)
sheet3 = ExcelFile3.sheet_by_name('优化期末余额')
optimize = []
for column in range(1, 15):
    m = []
    for row in range(1, 41):
        if row == 19 or row == 20:
            continue
        m.append(sheet3.cell(row, column).value)
    optimize.append(m)
x0 = optimize[month]

print(lcr(x0))

    
    
      