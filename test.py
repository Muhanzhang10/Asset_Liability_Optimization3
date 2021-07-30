#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jan 19 14:32:57 2021

@author: zhangmuhan
"""

import xlrd
import Data_Input as dp

path = "/Users/zhangmuhan/Desktop/实习/My_code/First/result_after.xls"

cut = len(dp.p_all)
ExcelFile = xlrd.open_workbook(path)
sheet = ExcelFile.sheet_by_name('优化期末余额')
sheet2 = ExcelFile.sheet_by_name('基准期末余额')

optimize = []
for column in range(1, 15):
    month = []
    for row in range(1, 41):
        if row == 19 or row == 20:
            continue
        month.append(sheet.cell(row, column).value)
    optimize.append(month)

optimize2 = []
for column in range(1, 15):
    month = []
    for row in range(1, 41):
        if row == 19 or row == 20:
            continue
        month.append(sheet2.cell(row, column).value)
    optimize2.append(month)



def run(month):
    total_month = 13
    
    
    
    total = []
    total_before = []
    after = []
    after_before = []
    
    for i in range(len(dp.p_all)):
        total.append(dp.p_all[i].balance_base[month])
        total_before.append(dp.p_all[i].balance_base[month - 1])
    for i in range(len(dp.d_all)):
        total.append(dp.d_all[i].balance_base[month])
        total_before.append(dp.d_all[i].balance_base[month - 1])


    
    expire_structure = []
    for i in range(len(dp.p_all)): 
        expire_structure.append(dp.p_all[i].expire_structure_shouzhi[month - 1])
    for i in range(len(dp.d_all)):
        expire_structure.append(dp.d_all[i].expire_structure_shouzhi[month - 1])    
    


    for i in range(len(optimize[0])):
        after.append(optimize[month][i])
        after_before.append(optimize[month - 1][i])
    
    
    
    def total_increment_capital(after): #after
        sum1 = 0
        sum2 = 0
        for i in range(len(total[:cut:])):
            sum1 += total[i] - total_before[i]
            sum2 += after[i] - after_before[i]
        return sum1 - sum2
 
    def total_increment_debt(after): #after
        sum1 = 0
        sum2 = 0
        for i in range(len(total[:cut:]), len(total)):
            sum1 += total[i] - total_before[i]
            sum2 += after[i] - after_before[i]
        return sum1 - sum2
    
    def rwa(after):
        sum1 = 0
        sum2 = 0
        for i in range(len(total[:cut:])):
            sum1 += (total[i] - total_before[i] + dp.p_all[i].expire_structure[month - 1]) * dp.p_all[i].RWA_weight
        for i in range(len(total[:cut:])):
            sum2 += (after[i] - after_before[i] + dp.p_all[i].expire_structure[month - 1]) * dp.p_all[i].RWA_weight
        return -(sum2 / sum1) + dp.RWA_Kr
    



    def find_bound():
        bound = []
        for i in range(len(total[:cut:])):
            difference = total[i] - total_before[i]
                
            if total[i] - total_before[i] != 0:
                difference = total[i] - total_before[i]
                up = max(difference * dp.p_all[i].range_up[month] + after_before[i], difference * dp.p_all[i].range_low[month] + after_before[i])
                low = min(difference * dp.p_all[i].range_up[month] + after_before[i], difference * dp.p_all[i].range_low[month] + after_before[i])
                low = max(low, 0)
            else:
                up = max(total[i] * (dp.p_all[i].range_low_spe[month] + 1), total[i] * (dp.p_all[i].range_up_spe[month] + 1))
                low = min(total[i] * (dp.p_all[i].range_low_spe[month] + 1), total[i] * (dp.p_all[i].range_up_spe[month] + 1))   
                low = max(low, 0)
            bound.append((low, up))
        for i in range(len(total[:cut:]), len(total)):
            difference = total[i] - total_before[i]
            if total[i] - total_before[i] != 0:
                difference = total[i] - total_before[i]
                up = max(difference * dp.d_all[i - cut].range_up[month] + after_before[i], difference * dp.d_all[i - cut].range_low[month] + after_before[i])
                low = min(difference * dp.d_all[i - cut].range_up[month] + after_before[i], difference * dp.d_all[i - cut].range_low[month] + after_before[i])
                low = max(low, 0)
            else:
                up = max(total[i] * (dp.d_all[i - cut].range_low_spe[month] + 1), total[i] * (dp.d_all[i - cut].range_up_spe[month] + 1))
                low = min(total[i] * (dp.d_all[i - cut].range_low_spe[month] + 1), total[i] * (dp.d_all[i - cut].range_up_spe[month] + 1))   
                low = max(low, 0)
            bound.append((low, up))
    
        return bound
    bound = find_bound()  #loower and upper bound for all assets included
    
    for i in range(len(bound)):
        if bound[i][0] - after[i] > 0.5: #or bound[i][1] < after[i]:
            print()
            print(bound[i][0], after[i], month)
            print()

    '''
    def NII():
        temp = 0
        a1_total = 0
        a2_total = 0
    
        for i in range(len(after_before[:cut:])):
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
        for i in range(len(after_before[:cut:]), len(after_before)):
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
    
        b1_total = 0
        b2_total = 0
    
        for i in range(len(after_before[:cut:])):
            if dp.p_all[i].is_current:
                b1_total += expire_structure[i] * dp.p_all[i].balance_shouzhi[month - 1]
            else:
                b2_total += expire_structure[i] * dp.p_all[i].balance_shouzhi[month - 1] / 1200
        for i in range(len(after_before[:cut:]), len(after_before)):
            if dp.d_all[i - cut].is_current:
                b1_total -= expire_structure[i] * dp.d_all[i - cut].balance_shouzhi[month - 1]
            else:
                b2_total -= expire_structure[i] * dp.d_all[i - cut].balance_shouzhi[month - 1] / 1200
        b1_total = b1_total * (month - 0.5) / 1200
    
        return (a1_total + a2_total + b1_total + b2_total) * -1 #因为之后函数只能找minimum，所以要乘-1来找maximum 
    '''
    
    

for i in range(1, 13):
    run(3)


