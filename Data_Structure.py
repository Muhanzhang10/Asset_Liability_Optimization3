#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jan 18 17:31:29 2021

@author: zhangmuhan
"""

class LCR_structure( object ):
    title = '' #lcr项目名称
    r1 = 0 #lcr折算系数
    r2 = 0 #lcr流出系数
    r3 = [0, 0, 0, 0] #押品比例系数
    attribute = 'i' #担保融资属性
    asset_type = '其它'#优质流动性资产
    '''
    def __init__(self, name ): #在这边弄没有用
            self.name = name
    '''
    
    

        
class Property_Debt( object ):  
        month_size = 13
        name = ''
        
        
        asset_type = "一级资产" #资产类型：一集资产, 2A, 2B, 其它
        is_mortgage = False #可抵质押融资属性
        AL = "A" #资产or负债
        
        
        expire_structure = [0] * month_size #常规到期结构，A_M，有0
        expire_structure_shouzhi = [0] * month_size  #收支到期结构，A_M,为了算NII的值，没有0
        balance_shouzhi = [0] * month_size  # 为了算NII的值，A_RM     
        rate_new = [] #新业务利率数量应该是 month_size  ，为了算NII，A_RM_N 
        
        is_current = True  #定期或者活期，定期为True，活期为False

        
        balance_base  = [0] * (month_size+1)  #数量是 month_size +1， 把初始状态作为第0个月，基准情况
        balance_after =  [0] * (month_size+1) #数量是 month_size +1， 把初始状态作为第0个月，优化之后的结果
        
        balance_after_lingshi =  [0] * (month_size+1) #和balance_afer 是同样的值，用作临时计算。
        
        
        range_up = []  #余额增量的上限系数 month_size 
        range_low = [] # 余额增量的下限系数 month_size
        range_up_spe  = []  #余额增量的上限 month_size 基于期末余额变动上限设置，在基准净增量为0的时候，用到这个系数
        range_low_spe = [] # 余额增量的下限 month_size 基于期末余额变动上限设置 在基准净增量为0的时候，用到这个系数
        
        range_low_true = [0] * month_size #由上限的系数算出的每月的不同变量的真实上限，为了展示和验证
        range_up_true = [0] * month_size #由上限的系数算出的每月的不同变量的真实下限，为了展示和验证
        
        
        NII_after = [0] * (month_size + 1) # 优化之后的每月的NII值。这个是当月的每一个变量都一样 ,多加的一个月 在最后面，用于存入总的NII值
        NII_base  = [0] * ( month_size +1 ) # 基准的每月的NII值 ,多加的一个月 在最后面，用于存入总的NII值
        status_after =  [0] * month_size #记录每月是否有解，这个也是每一个月的变量都一样
        
        LCR_after = [0] * month_size # 每月优化之后的LCR的值。 这个我没有弄的原因是， 不同的资产和负债的当月的LCR可以一样，所以我就没有 给每一个资产和负债都 给一个list ，而是公用一个size为13的list
        LCR_base =  [0] * month_size #每月基准LCR的值
        
        
        RWA_after = [0] * month_size #每月优化之后的RWA的值
        RWA_base = [0] * month_size # 每月基准的RWA的值
        RWA_weight = 0 # 每个资产的风险权重，一个资产一个值
        
        RWA_constraint_value = [1.05] * month_size    # RWA 约束值，,每一个月都可以不一样，这个程序目前取的是同样的值
        LCR_constraint_value = [1.2] * month_size    # LCR 约束值，,每一个月都可以不一样  这个程序目前取的是同样的值
        
        lcr = [] #every element is a LCR_structure object

        
        
        
        
        
        def __init__(self, name ): #在这边弄没有用
            self.name = name
        
                 