'''
使用注意事项：
1.name_list文件必须和次.py文件放在同一文件夹下
2.name_list文件的第一列必须为姓名（被抽对象）
3.name_list文件的第二列必须为活动参与情况（注意不是权重，和权重相反，活动参与情况越大，权重越小）
'''

import random
import pandas as pd

def random_choose(name_lst,weight,num):#抽取函数
    """
    列表抽取函数，根据权重（weight）在目标列表（name_list）抽不重复的抽取num个元素
    :param name_lst:被抽取列表
    :param weight:权重列表
    :param num:抽取个数
    :return:抽取结果列表
    """
    if len(name_lst) != len(weight):
        raise Exception("name_list len is not match weigth len")
    elif num > len(name_lst):
        raise Exception("num is too large")
    id_list = list(range(len(name_lst)))    #抽取人数id列表
    ret_id = []                             #抽取id
    ret = []                                #抽取结果
    for i in range(num):
        if sum(weight)==0:                  #判断是否所有权重都为0
            weight = [(0 if i in ret_id else 1) for i in range(weight)]
        choose_id = random.choices(id_list, weights = weight);
        ret_id.append(choose_id[0])
        ret.append(name_lst[choose_id[0]])
        weight[choose_id[0]] = 0
    return ret

def weight_create(org,times):
    '''
    权重生成算法，期待更好算法
    :param times:参与活动次数
    :return:权重列表
    '''
    weight = [max(times)+1-i for i in times]#简单生成权重
    if len(org.columns) > 2:#判断最后一列是否为最新一次的活动参加人员
        last_join = list(org.iloc[:,-1])
        ret = [(0 if last_join[i] == 1 else weight[i]) for i in range(len(weight))]
    else:
        ret = weight
    return ret

def save(address, action_name, name_list):
    """
    用于储存结果
    Parameters
    ----------
    address : char
        地址
    action_name : char
        活动名
    name_list : list
        活动参加人员
    Returns
    -------
    None.

    """
    org = pd.read_excel(address, index_col=0)
    action_list = [0 for i in range(len(org))]
    org.insert(loc=len(org.columns), column= action_name, value= action_list)
    for i in name_list:
        org.loc[i,"times"] += 1
        org.loc[i, action_name] = 1
    org.to_excel(address,sheet_name = "sheel1",index = True ,na_rep = 0,inf_rep = 0)

def choose_people(address, num, action_name = None):     #抽人函数
    '''
    从地址为address的文件中，抽取数量为num个元素
    注意：address文件中的第一列必须为被抽取对象（姓名），第二列必须为权重（活动参与次数）
    :param adderss:文件地址
    :param num:抽取人数
    :return:抽取结果
    '''
    org = pd.read_excel(address)        #文件读取
    name_lst = list(org.iloc[:,0])   #姓名列表
    times_lst = list(org.iloc[:,1])#参与活动次数列表
    weight = weight_create(org,times_lst)  #权重生成
    name_list = random_choose(name_lst,weight,num)
    if action_name != None:
        save(address, action_name, name_list)
    return name_list
        
        
if __name__ == "__main__":
    choose = choose_people("name_list.xlsx", 2)
    
    
    
    
    