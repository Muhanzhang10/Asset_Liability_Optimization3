#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jan 19 10:42:33 2021

@author: zhangmuhan
"""

from Data_Structure import Property_Debt
from Data_Structure import LCR_structure
import xlrd

d_all = []
p_all = []

RWA_Kr = 1.05
    
def run(data_in_file, data_in_file2): 
    month_size = 13
    ExcelFile=xlrd.open_workbook(data_in_file)
    sheet=ExcelFile.sheet_by_name('存量业务到期结构')
    
    
    a1 = 0
    for i in range(100):
        if sheet.cell( i - 1, ord('B') - 65  ).value == '负债端':
            a1 = i
            break
    
    
    for i in range(1,a1):   
        if sheet.cell( i - 1, ord('C') - 65  ).value == 'P':
            name = sheet.cell( i - 1, ord('B') - 65  ).value
            p_all.append( Property_Debt(  name   ) )
            
    
    for j in range(a1,1000):
        try:
            if sheet.cell( j - 1, ord('C') - 65  ).value == 'P':
                
                name =   sheet.cell( j - 1, ord('B') - 65  ).value
                
                d_all.append(  Property_Debt(  name      ) )
        except BaseException:
            break
        
    #到期  收支的表
    sheet=ExcelFile.sheet_by_name('存量业务利息收支')
    for k in p_all : 
        a3 = 1
        for i in range(a3,100):
            if k.name == sheet.cell( i - 1, ord('A') - 65  ).value:        
                a2 = []
                a3 = i #定义新的起点，减少循环次数
                for j in range( ord('B') -65 , ord('B') -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.expire_structure_shouzhi = a2
                break
    
    
    for k in d_all : 
        a3 = 1
        for i in range(a3,100):
            if k.name == sheet.cell( i - 1, ord('A') - 65  ).value:        
                a2 = []
                a3 = i #定义新的起点，减少循环次数
                for j in range( ord('B') -65 , ord('B') -65 + month_size  ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.expire_structure_shouzhi = a2
                break
            
            




#到期  非收支的表
    sheet=ExcelFile.sheet_by_name('存量业务到期结构')
    for k in p_all : 
        a3 = 1
        for i in range(a3,100):
            if k.name == sheet.cell( i - 1, ord('B') - 65  ).value:        
                a2 = []
                a3 = i #定义新的起点，减少循环次数
                for j in range( ord('F') -65 , ord('F') -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.expire_structure = a2
                break
    
    
    for k in d_all : 
        a3 = 1
        for i in range(a3,100):
            if k.name == sheet.cell( i - 1, ord('B') - 65  ).value:        
                a2 = []
                a3 = i #定义新的起点，减少循环次数
                for j in range( ord('F') -65 , ord('F') -65 + month_size  ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.expire_structure = a2
                break


            
    
        
    # 为了给每一个 对象钟加 list    
    for k in p_all : 
        a2 = [0] * month_size
        k.status_after = a2
        a2 = [0] * month_size
        k.NII_after = a2
        a2 = [0] * month_size
        k.NII_base = a2
        a2 = [0] * month_size
        k.range_low_true = a2
        a2 = [0] * month_size
        k.range_up_true = a2
        
        a2 = [0] * (month_size + 1)
        k.balance_after = a2
        
        a2 = [0] * (month_size + 1)
        k.balance_after_lingshi = a2
        
        
        
        
        
        
    for k in d_all : 
        a2 = [0] * month_size
        k.status_after = a2
        a2 = [0] * month_size
        k.NII_after = a2
        a2 = [0] * month_size
        k.NII_base = a2
        a2 = [0] * month_size
        k.range_low_true = a2
        a2 = [0] * month_size
        k.range_up_true = a2
        
        a2 = [0] * (month_size + 1)
        k.balance_after = a2
        
        a2 = [0] * (month_size + 1)
        k.balance_after_lingshi = a2
        
        
        
        
        
        
        
     
    #存量业务到期收支
    sheet=ExcelFile.sheet_by_name('存量业务利息收支')
    
    a1 = 0
    for i in range(1000):
        if sheet.cell( 3 - 1, i  ).value == '存量业务到期收益率':
            a1 = i
            break
    
    for k in p_all : 
        a3 = 1
        for i in range(a3,100):
            if k.name == sheet.cell( i - 1, ord('A') - 65  ).value:        
                a2 = []
                a3 = i #定义新的起点，减少循环次数
                for j in range( a1 , a1 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.balance_shouzhi = a2
                break
    
    
    for k in d_all : 
        a3 = 1
        for i in range(a3,100):
            if k.name == sheet.cell( i - 1, ord('A') - 65  ).value:        
                a2 = []
                a3 = i #定义新的起点，减少循环次数
                for j in range( a1 , a1 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.balance_shouzhi = a2
                break
    
    
    

    
        
        
    #基准值
    sheet=ExcelFile.sheet_by_name('期末余额_基准情景')
    sheet2=ExcelFile.sheet_by_name('业务净增量_基准情景')
    for k in p_all : 
        a3 = 1
        for i in range(a3,100):
            if k.name == sheet.cell( i - 1, ord('B') - 65  ).value:        
                a2 = []
                a3 = i #定义新的起点，减少循环次数
                a2.append( sheet.cell( i-1, ord('C') -65  ).value -  sheet2.cell( i-1, ord('C') -65  ).value   )
                
                for j in range( ord('C') -65 , ord('C') -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.balance_base = a2
                break
    
    
    
    
    for k in d_all : 
        a3 = 1
        for i in range(a3,100):
            if k.name == sheet.cell( i - 1, ord('B') - 65  ).value:        
                a2 = []
                a3 = i #定义新的起点，减少循环次数
                
                a2.append( sheet.cell( i-1, ord('C') -65  ).value -  sheet2.cell( i-1, ord('C') -65  ).value   )
    
                for j in range( ord('C') -65 , ord('C') -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.balance_base = a2
                break
    
    

    
   
    
    
    
    
     #现在把第0个月的基准的值，当成第0个月优化的值    有问题
    
    for k in p_all : 
        k.balance_after[0] = k.balance_base[0]
        k.balance_after_lingshi[0] = k.balance_base[0]
    
    for k in d_all : 
        k.balance_after[0] = k.balance_base[0]
        k.balance_after_lingshi[0] = k.balance_base[0]
    
 
    
    
    
    
    ExcelFile=xlrd.open_workbook(data_in_file)

    
    #RWA_weight， 负债没有那个
    sheet=ExcelFile.sheet_by_name('RWA风险权重')
    
    for k in p_all : 
        a3 = 1
        for i in range(a3,100):
            if k.name == sheet.cell( i - 1, ord('B') - 65  ).value:        
            
                a3 = i #定义新的起点，减少循环次数
                
                if sheet.cell( i-1, ord('C') - 65  ).value == '':
                    k.RWA_weight = 0.0
                else:
                    k.RWA_weight =   sheet.cell( i-1, ord('C') - 65  ).value
                break
    
    

    
    sheet=ExcelFile.sheet_by_name('业务净增量变动限制设置')
    #上限 正常
    for k in p_all : 
        a3 = 1
        for i in range(a3,300):
            if k.name == sheet.cell( i - 1, ord('A') - 65  ).value:        
                a2 = []
                a4 = []
                a3 = i #定义新的起点，减少循环次数
                for j in range( ord('B') -65 , ord('B') -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.range_up = a2       
                for j in range( ord('Z') + 17 -65 ,  ord('Z') + 17 -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a4.append(0.0)
                    else:
                        a4.append(  sheet.cell( i-1, j  ).value  )
                k.range_low = a4
                break
    
    
    for k in d_all : 
        a3 = 1
        for i in range(a3,300):
            if k.name == sheet.cell( i - 1, ord('A') - 65  ).value:        
                a2 = []
                a4 = []
                a3 = i #定义新的起点，减少循环次数
                for j in range( ord('B') -65 , ord('B') -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.range_up = a2       
                for j in range( ord('Z') + 17 -65 ,  ord('Z') + 17 -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a4.append(0.0)
                    else:
                        a4.append(  sheet.cell( i-1, j  ).value  )
                k.range_low = a4
                break
    
    
    
    
    
    
    #上下限_基于期末余额变动上限设置
    number = 0
    for i in range(300):
        if sheet.cell( i - 1, ord('B') - 65  ).value == "基于期末余额变动上限设置":
            number = i
            break
    
    for k in p_all : 
        a3 = number
        for i in range(a3,300):
            if k.name == sheet.cell( i - 1, ord('A') - 65  ).value:        
                a2 = []
                a4 = []
                a3 = i #定义新的起点，减少循环次数
                for j in range( ord('B') -65 , ord('B') -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.range_up_spe = a2 
                for j in range( ord('Z') + 17 -65 ,  ord('Z') + 17 -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a4.append(0.0)
                    else:
                        a4.append(  sheet.cell( i-1, j  ).value  )
                k.range_low_spe = a4
                break
    
    
    
    for k in d_all : 
        a3 = number
        for i in range(a3,300):
            if k.name == sheet.cell( i - 1, ord('A') - 65  ).value:   
                a2 = []
                a4 = []
                a3 = i #定义新的起点，减少循环次数
                for j in range( ord('B') -65 , ord('B') -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.range_up_spe = a2 
                for j in range( ord('Z') + 17 -65 ,  ord('Z') + 17 -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a4.append(0.0)
                    else:
                        a4.append(  sheet.cell( i-1, j  ).value  )
                k.range_low_spe = a4
                break
    
    
    
            
 
    
    
    sheet=ExcelFile.sheet_by_name('新业务利率假设')
    for k in p_all : 
        a3 = 1
        for i in range(a3,300):
            if k.name == sheet.cell( i - 1, ord('B') - 65  ).value:        
                a2 = []
                a3 = i #定义新的起点，减少循环次数
                for j in range( ord('C') -65 , ord('C') -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.rate_new = a2
                break
    
    for k in d_all : 
        a3 = 1
        for i in range(a3,100):
            if k.name == sheet.cell( i - 1, ord('B') - 65  ).value:        
                a2 = []
                a3 = i #定义新的起点，减少循环次数
                for j in range( ord('C') -65 , ord('C') -65 + month_size ):
                    if sheet.cell( i-1, j  ).value == '':
                        a2.append(0.0)
                    else:
                        a2.append(  sheet.cell( i-1, j  ).value  )
                k.rate_new = a2
                break
    
    
    
    p_all[3].is_current = False
    p_all[10].is_current = False
    p_all[12].is_current = False
    d_all[4].is_current = False
    d_all[5].is_current = False
    d_all[9].is_current = False
    d_all[10].is_current = False
    d_all[11].is_current = False
    
    
    

    
    ExcelFile2 = xlrd.open_workbook(data_in_file2)
    

    sheet = ExcelFile2.sheet_by_name('属性')
    
    index = 0
    for row in range(1, 40):
        if row < 19:
            p_all[index].AL = sheet.cell(row, 1).value 
            
            if sheet.cell(row, 3).value == 'y':
                p_all[index].is_mortgage = True
            else:
                p_all[index].is_mortgage = False
                
            index += 1
            
        elif row == 19:
            index = 0
        else:
            d_all[index].AL = sheet.cell(row, 1).value
            
            if sheet.cell(row, 3).value == 'y':
                d_all[index].is_mortgage = True
            else:
                d_all[index].is_mortgage = False
            
            index += 1
    
    
    
    for i in range(len(p_all)):
        p_all[i].lcr = []
    for i in range(len(d_all)):
        d_all[i].lcr = []

    sheet = ExcelFile2.sheet_by_name('LCR项目系数')
    cut = len(p_all)
    objs = [LCR_structure() for i in range(144)]        
    for row in range(1, 145):
        index = int(sheet.cell(row, 0).value)
        objs[row - 1] = LCR_structure() 
        objs[row - 1].r2 = sheet.cell(row, 6).value
        objs[row - 1].r1 = sheet.cell(row, 7).value
        objs[row - 1].title = sheet.cell(row, 8).value
        objs[row - 1].asset_type = sheet.cell(row, 9).value
        objs[row - 1].attribute = sheet.cell(row, 10).value
        objs[row - 1].r3 = [0, 0, 0, 0]
        objs[row - 1].r3[0] = sheet.cell(row, 2).value
        if objs[row - 1].attribute == 'D':
            objs[row - 1].r3[1] = sheet.cell(row, 3).value
            objs[row - 1].r3[2] = sheet.cell(row, 4).value
            objs[row - 1].r3[3] = sheet.cell(row, 5).value
        
        if index <= cut:
            p_all[index - 1].lcr.append(objs[row - 1])
            #print(len(p_all[0].lcr), '0')
            #print(len(p_all[1].lcr), '1')
        else:
            d_all[index - cut - 1].lcr.append(objs[row - 1])
           
    

            
    
    return p_all, d_all




data_in_file = r"组合管理优化模板_v8.4_20181206.xlsm"
data_in_file2 = r"Excel模板涉及 （38）.xlsx"

p_all, d_all = run(data_in_file, data_in_file2)







    
    
    

