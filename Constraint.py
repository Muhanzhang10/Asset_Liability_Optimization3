#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jan 18 17:35:02 2021

@author: zhangmuhan
"""

#Ai_Mj_B j月，基准期末余额   balance_base
#Ai_Mj_BP j月，优化后资产期末余额  balance_after
#A_M 该项资产存量到期结构

import numpy as np
from Data_Structure import Property_Debt
import Data_Input as dp
import pandas as pd
from pandas import ExcelWriter
import xlrd
from scipy.optimize import minimize
#import matplotlib.pyplot as plt
#import matplotlib.rcParams



#month: 1- 13
def run(month):
    total_month = 13  #这里之后要改！！！用于计算NII
    cut = len(dp.p_all)
    #total: 当月负债和资产的base总和
    #total_before:上月负债和资产的base的总和
    #after:当月负债和资产的优化总和
    #after_before：上月负债和资产的优化总和
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
            
    '''
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
    #after_before = optimize[month]
    '''
  
    expire_structure = []
    for i in range(len(dp.p_all)): 
        expire_structure.append(dp.p_all[i].expire_structure_shouzhi[month - 1])
    for i in range(len(dp.d_all)):
        expire_structure.append(dp.d_all[i].expire_structure_shouzhi[month - 1])    
    


    #优化前后资产总增量相等
    #equal
    def total_increment_capital(after): #after
        sum1 = 0
        sum2 = 0
        for i in range(len(total[:cut:])):
            sum1 += total[i] - total_before[i]
            sum2 += after[i] - after_before[i]
        return sum1 - sum2
    

    #equal
    def total_increment_debt(after): #after
        sum1 = 0
        sum2 = 0
        for i in range(len(total[:cut:]), len(total)):
            sum1 += total[i] - total_before[i]
            sum2 += after[i] - after_before[i]
        return sum1 - sum2
    
    
    #@para
    #rwaValue: used for adjust
    def rwa(after, value):
        sum1 = 0
        sum2 = 0
        for i in range(len(total[:cut:])):
            sum1 += (total[i] - total_before[i] + dp.p_all[i].expire_structure[month - 1]) * dp.p_all[i].RWA_weight
        for i in range(len(total[:cut:])):
            sum2 += (after[i] - after_before[i] + dp.p_all[i].expire_structure[month - 1]) * dp.p_all[i].RWA_weight
        return -(sum2 / sum1) + value
         
    
    #所有资产和负债都需大于0
    #@para
    #adjust1:期末余额0.05
    #adjust2:净增量 0.1
    def find_bound(adjust1, adjust2):
        bound = []
        for i in range(len(total[:cut:])):
            difference = total[i] - total_before[i] 
            if total[i] - total_before[i] != 0:
                difference = total[i] - total_before[i]
                up = max(difference * dp.p_all[i].range_up[month] + after_before[i] + adjust1, difference * dp.p_all[i].range_low[month] + after_before[i] + adjust1)
                low = min(difference * dp.p_all[i].range_up[month] + after_before[i] - adjust1, difference * dp.p_all[i].range_low[month] + after_before[i] - adjust1)
                low = max(low, 0)
            else:
                up = max(after_before[i] * (dp.p_all[i].range_low_spe[month] + 1) + adjust2, after_before[i] * (dp.p_all[i].range_up_spe[month] + 1) + adjust2)
                low = min(after_before[i] * (dp.p_all[i].range_low_spe[month] + 1) - adjust2, after_before[i] * (dp.p_all[i].range_up_spe[month] + 1) - adjust2)   
                low = max(low, 0)
            bound.append((low, up))
        for i in range(len(total[:cut:]), len(total)):

            difference = total[i] - total_before[i]
            if total[i] - total_before[i] != 0:
                difference = total[i] - total_before[i]
                up = max(difference * dp.d_all[i - cut].range_up[month] + after_before[i] + adjust1, difference * dp.d_all[i - cut].range_low[month] + after_before[i] + adjust1)
                low = min(difference * dp.d_all[i - cut].range_up[month] + after_before[i] - adjust1, difference * dp.d_all[i - cut].range_low[month] + after_before[i] - adjust1)
                low = max(low, 0)
            else:
                up = max(after_before[i] * (dp.d_all[i - cut].range_low_spe[month] + 1) + adjust2, after_before[i] * (dp.d_all[i - cut].range_up_spe[month] + 1) + adjust2)
                low = min(after_before[i] * (dp.d_all[i - cut].range_low_spe[month] + 1) - adjust2, after_before[i] * (dp.d_all[i - cut].range_up_spe[month] + 1) - adjust2)   
                low = max(low, 0)
            bound.append((low, up))
    
        return bound
    
    
    def lcr(after, value, is_base):
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
        
        if is_base == False:
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
        

        #return  C1 / C2 - 1.2
        
        if is_base == True:
            return C1 / C2

        return C1 / C2 - value
    
    
    
    
    
    
    #利息收支
    def NII_interest():
        b1_total = 0
        b2_total = 0

        for i in range(len(total[:cut:])):
            if dp.p_all[i].is_current:
                b1_total += expire_structure[i] * dp.p_all[i].balance_shouzhi[month - 1]
            else:
                b2_total += expire_structure[i] * dp.p_all[i].balance_shouzhi[month - 1] / 1200
        for i in range(len(total[:cut:]), len(total)):
            if dp.d_all[i - cut].is_current:
                b1_total -= expire_structure[i] * dp.d_all[i - cut].balance_shouzhi[month - 1]
            else:
                b2_total -= expire_structure[i] * dp.d_all[i - cut].balance_shouzhi[month - 1] / 1200
        b1_total = b1_total * (month - 0.5) / 1200
        
        return b1_total + b2_total
        

    def NII(after):
        temp = 0
        a1_total = 0
        a2_total = 0

        for i in range(len(total[:cut:])):
            if dp.p_all[i].is_current:
                a1_total += (after[i] - after_before[i] + dp.p_all[i].expire_structure[month - 1]) * dp.p_all[i].rate_new[month - 1]
            else:
                temp1 = 0
                temp2 = after[i] - after_before[i] + dp.p_all[i].expire_structure[month - 1]
                sigma = 0.5 * dp.p_all[i].rate_new[month - 1]
                for j in range(month, total_month):
                    sigma += dp.p_all[i].rate_new[j - 1]
                temp1 = temp2 * sigma
                a2_total += temp1
        for i in range(len(total[:cut:]), len(total)):
            if dp.d_all[i - cut].is_current:
                temp += (after[i] - after_before[i] + dp.d_all[i - cut].expire_structure[month - 1]) * dp.d_all[i - cut].rate_new[month - 1]
                a1_total -= (after[i] - after_before[i] + dp.d_all[i - cut].expire_structure[month - 1]) * dp.d_all[i - cut].rate_new[month - 1]
            else:
                temp1 = 0
                temp2 = after[i] - after_before[i] + dp.d_all[i - cut].expire_structure[month - 1]
                sigma = 0.5 * dp.d_all[i - cut].rate_new[month - 1]
                for j in range(month, total_month):
                    sigma += dp.d_all[i - cut].rate_new[j - 1]
                temp1 = temp2 * sigma
                a2_total -= temp1
        a2_total = a2_total / 1200
        a1_total = a1_total * (total_month + 1 - month - 0.5) / 1200  #这里total month要改！

        return (a1_total + a2_total + NII_interest()) * -1 #因为之后函数只能找minimum，所以要乘-1来找maximum 
    
    
    def NII_before(after):
        temp = 0
        a1_total = 0
        a2_total = 0

        for i in range(len(total[:cut:])):
            if dp.p_all[i].is_current:
                a1_total += (after[i] - total_before[i] + dp.p_all[i].expire_structure[month - 1]) * dp.p_all[i].rate_new[month - 1]
            else:
                temp1 = 0
                temp2 = after[i] - total_before[i] + dp.p_all[i].expire_structure[month - 1]
                sigma = 0.5 * dp.p_all[i].rate_new[month - 1]
                for j in range(month, total_month):
                    sigma += dp.p_all[i].rate_new[j - 1]
                temp1 = temp2 * sigma
                a2_total += temp1
        for i in range(len(total[:cut:]), len(total)):
            if dp.d_all[i - cut].is_current:
                temp += (after[i] - total_before[i] + dp.d_all[i - cut].expire_structure[month - 1]) * dp.d_all[i - cut].rate_new[month - 1]
                a1_total -= (after[i] - total_before[i] + dp.d_all[i - cut].expire_structure[month - 1]) * dp.d_all[i - cut].rate_new[month - 1]
            else:
                temp1 = 0
                temp2 = after[i] - total_before[i] + dp.d_all[i - cut].expire_structure[month - 1]
                sigma = 0.5 * dp.d_all[i - cut].rate_new[month - 1]
                for j in range(month, total_month):
                    sigma += dp.d_all[i - cut].rate_new[j - 1]
                temp1 = temp2 * sigma
                a2_total -= temp1
        a2_total = a2_total / 1200
        a1_total = a1_total * (total_month + 1 - month - 0.5) / 1200  #这里total month要改！


        return (a1_total + a2_total + NII_interest()) * -1 #因为之后函数只能找minimum，所以要乘-1来找maximum
    
    
    rwaValue = 1.05
    lcrValue = 1.2
    adjust1 = 0
    adjust2 = 0
    
    def lcr_Cal(after):
        return lcr(after, lcrValue, False)

    def rwa_Cal(after):
        
        return rwa(after, rwaValue)
    
    
    
    
    bound = find_bound(adjust1, adjust2)  #lower and upper bound for all assets included

    cons = []
    cons.append({'type': 'eq', 'fun': total_increment_capital})
    cons.append({'type': 'eq', 'fun': total_increment_debt})
    cons.append({'type': 'ineq', 'fun': rwa_Cal})
    cons.append({'type': 'ineq', 'fun': lcr_Cal})
    
    def is_Success(message):
        if message == "`gtol` termination condition is satisfied.":
            return True
        return False
    
    #调整优化
    while True:
        res = minimize(NII, total, method = 'trust-constr', bounds = bound, constraints = cons, options = {"maxiter": 300})
        if is_Success(res.message) == True:
            break
        

        print("Move to LCR")
        if lcrValue != 1:
            lcrValue -= 0.05
            if lcrValue < 1:
                lcrValue = 1
            res = minimize(NII, total, method = 'trust-constr', bounds = bound, constraints = cons, options = {"maxiter": 300})
            if is_Success(res.message) == True:
                break

        
        print("Move to adjust")
        adjust1 += 0.05
        adjust2 += 0.1
        bound = find_bound(adjust1, adjust2)
        res = minimize(NII, total, method = 'trust-constr', bounds = bound, constraints = cons, options = {"maxiter": 300})
        if is_Success(res.message) == True:
            break
        
        print("Move to RWA")
        rwaValue += 0.02
        res = minimize(NII, total, method = 'trust-constr', bounds = bound, constraints = cons, options = {"maxiter": 300})
        if is_Success(res.message) == True:
            break

    

    for i in range(len(res.x[:cut:])):
        dp.p_all[i].balance_after[month] = res.x[i]
    for i in range(len(res.x[:cut:]), len(res.x)):
        dp.d_all[i - cut].balance_after[month] = res.x[i]
    

    print("Month", month)
    print("资产前后差值", total_increment_capital(res.x))
    print("负债前后差值", total_increment_debt(res.x))
    print("rwa指标", rwaValue - rwa_Cal(res.x))
    print("LCR", lcr_Cal(res.x) + lcrValue)
    print("NII_after", res.fun)
    print("NII_before", NII_before(total))
    print(res.message)
    print()

    constraints = [rwaValue, lcrValue, adjust1, adjust2]
    return res.x, res.fun, total, NII_before(total), lcr(total, 0, True), lcr_Cal(res.x) + lcrValue, 1.05 - 0.05, rwaValue - rwa_Cal(res.x), -NII_before(total) - NII_interest(), -res.fun - NII_interest(), bound, constraints



    
    


def increment(month):
    lis = [] #毛增量优化
    lis2 = [] #毛增量基准
    lis3 = [] #净增量优化
    lis4 = [] #净增量基准
    for i in range(len(dp.p_all)):
        num = dp.p_all[i].balance_after[month] - dp.p_all[i].balance_after[month - 1] + dp.p_all[i].expire_structure[month - 1]
        num2 = dp.p_all[i].balance_base[month] - dp.p_all[i].balance_base[month - 1] + dp.p_all[i].expire_structure[month - 1]
        num3 = dp.p_all[i].balance_after[month] - dp.p_all[i].balance_after[month - 1]
        num4 = dp.p_all[i].balance_base[month] - dp.p_all[i].balance_base[month - 1]
        lis.append(num)
        lis2.append(num2)
        lis3.append(num3)
        lis4.append(num4)
    lis.append('')
    lis2.append('')
    lis3.append('')
    lis4.append('')
    for i in range(len(dp.d_all)):
        num = dp.d_all[i].balance_after[month] - dp.d_all[i].balance_after[month - 1] + dp.d_all[i].expire_structure[month - 1]
        num2 = dp.d_all[i].balance_base[month] - dp.d_all[i].balance_base[month - 1] + dp.d_all[i].expire_structure[month - 1]
        num3 = dp.d_all[i].balance_after[month] - dp.d_all[i].balance_after[month - 1] 
        num4 = dp.d_all[i].balance_base[month] - dp.d_all[i].balance_base[month - 1]
        lis.append(num)
        lis2.append(num2)
        lis3.append(num3)
        lis4.append(num4)
    return lis, lis2, lis3, lis4


def create_file():

    path = '/Users/zhangmuhan/Desktop/实习/My_code/Result' + '/结果.xlsx'
    writer = ExcelWriter(path)
    dic1 = {} #balance_after
    dic2 = {} #NII before and after
    dic3 = {}
    balance_Base = {}
    increment_Base = {}
    increment_Base2 = {}
    increment_after2 = {}
    LCR = {}
    RWA = {}
    expire_Structure = {}
    interest = {} #新业务利息收支
    bound_max = {} #上限
    bound_min = {} #下限
    constraints = {}
    NII_before = []
    NII_after = []
    LCR_before = []
    LCR_after = []
    RWA_before = []
    RWA_after = []
    interest_before = []
    interest_after = []
    RWA_bound = []
    LCR_bound = []
    bound_adjust1 = []
    bound_adjust2 = []
    
    
    
    
    #add_name
    name_list = []
    for i in range(len(dp.p_all)):
        name_list.append(dp.p_all[i].name)
    name_list.append('')
    for i in range(len(dp.d_all)):
        name_list.append(dp.d_all[i].name)
    dic1['Item_name'] = name_list
    dic3['Item_name'] = name_list
    balance_Base['Item_name'] = name_list
    increment_Base['Item_name'] = name_list
    increment_Base2['Item_name'] = name_list
    increment_after2['Item_name'] = name_list
    expire_Structure['Item_name'] = name_list
    bound_max['Item_name'] = name_list
    bound_min['Item_name'] = name_list
    
    
    
    for month in range(1, 13):
        result = run(month)
        NII_after.append(-result[1])
        NII_before.append(-result[3])
        LCR_before.append(result[4])
        LCR_after.append(result[5])
        RWA_before.append(result[6])
        RWA_after.append(result[7])
        interest_before.append(result[8])
        interest_after.append(result[9])
        bound = result[10]
        RWA_bound.append(result[11][0])
        LCR_bound.append(result[11][1])
        bound_adjust1.append(result[11][2])
        bound_adjust2.append(result[11][3])
        up = []
        low = []
        for item in range(len(dp.p_all) + len(dp.d_all)):
            up.append(bound[item][1])
            low.append(bound[item][0])
            if item == len(dp.p_all) - 1:
                up.append('')
                low.append('')
        bound_max[month] = up
        bound_min[month] = low
        
        
    #plt.plot([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12], NII_before)
    #plt.plot([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12], NII_after)
    #plt.xlabel('month')
    #plt.ylabel('NII')
    
    #plt.show()
        
    NII_after.append(sum(NII_after))
    NII_before.append(sum(NII_before))
    
    dic2 = {'month': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 'total'], 'NII_before': NII_before, 'NII_after': NII_after}
    LCR = {'month': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12], 'LCR_before': LCR_before, 'LCR_after': LCR_after}
    RWA = {'month': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12], 'RWA_before': RWA_before, 'RWA_after': RWA_after}
    interest = {'month': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12], 'interest_before': interest_before, 'interest_after': interest_after}
    constraints = {'month': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]}
    constraints['LCR_bound'] = LCR_bound
    constraints['RWA_bound'] = RWA_bound
    constraints['期末余额'] = bound_adjust1
    constraints['净增量'] = bound_adjust2
   
    
    #add_number
    for month in range(1, 13):
        month_list = []
        balance_Base_list = []
        expire_structure_list = []
        for i in range(len(dp.p_all)):
            month_list.append(dp.p_all[i].balance_after[month])
            balance_Base_list.append(dp.p_all[i].balance_base[month])
            expire_structure_list.append(dp.p_all[i].expire_structure[month - 1])
        month_list.append('')
        balance_Base_list.append('')
        expire_structure_list.append('')
        for i in range(len(dp.d_all)):
            month_list.append(dp.d_all[i].balance_after[month])
            balance_Base_list.append(dp.d_all[i].balance_base[month])
            expire_structure_list.append(dp.d_all[i].expire_structure[month - 1])
        dic1[month] = month_list
        balance_Base[month] = balance_Base_list
        expire_Structure[month] = expire_structure_list
        
    #add increment
    for month in range(1, 13):
        dic3[month] = increment(month)[0]
        increment_Base[month] = increment(month)[1]
        increment_Base2[month] = increment(month)[3]
        increment_after2[month] = increment(month)[2]
        
    

    
    dic1 = pd.DataFrame(dic1)
    dic2 = pd.DataFrame(dic2)
    LCR = pd.DataFrame(LCR)
    RWA = pd.DataFrame(RWA)
    dic3 = pd.DataFrame(dic3)
    balance_Base = pd.DataFrame(balance_Base)
    increment_Base = pd.DataFrame(increment_Base)
    increment_Base2 = pd.DataFrame(increment_Base2)
    increment_after2 = pd.DataFrame(increment_after2)
    expire_Structure = pd.DataFrame(expire_Structure)
    interest = pd.DataFrame(interest)
    bound_max = pd.DataFrame(bound_max)
    bound_min = pd.DataFrame(bound_min)
    constraints = pd.DataFrame(constraints)
    dic1.to_excel(writer, '各项资产负债优化结果', index = False)
    balance_Base.to_excel(writer, '各项资产负债基准结果', index = False)
    dic2.to_excel(writer, 'NII结果前后对比', index = False)
    dic3.to_excel(writer, '优化后毛增量', index = False)
    increment_Base.to_excel(writer, '基准毛增量', index = False)
    increment_Base2.to_excel(writer, '基准净增量', index = False)
    increment_after2.to_excel(writer, '优化后净增量', index = False)
    RWA.to_excel(writer, 'RWA结果前后对比')
    LCR.to_excel(writer, 'LCR结果前后对比')
    expire_Structure.to_excel(writer, '存量到期结构')
    interest.to_excel(writer, '新业务利息收支')
    bound_max.to_excel(writer, '上限')
    bound_min.to_excel(writer, '下限')
    constraints.to_excel(writer, '调整后限制')
    
    writer.save()


create_file()




