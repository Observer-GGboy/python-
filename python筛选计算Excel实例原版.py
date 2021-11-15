import numpy as np
import pandas as pd
from datetime import datetime
import xlwings as xw
from math import ceil
import time
import msvcrt

''' 
# 从美式计数法如 1,999.13 转换为普通的数值,
# 传入值为一个美式计数法的数字字符串，
# 返回转换结果
'''
def AmericanNumber2NormalNumber(americanNumberStr):
    americanNumber = eval(americanNumberStr)# 去掉字符串引号
    res = 0 # 保存计算结果
    countOfThoundsSignal = americanNumberStr.count(",") # 计算逗号个数
    # 处理超过千位的数字
    if type(americanNumber) is tuple:
        weight = 3*countOfThoundsSignal #获取当前位数的权重:3*逗号的个数
        for i in range(countOfThoundsSignal):
            res += americanNumber[i]*(10**weight)
            weight = weight-3
        res += americanNumber[countOfThoundsSignal]
    else:
        res += americanNumber
    return res

#数据处理,转化为数字一维数组
def dataProccess(res_list):
    res = []
    for i in res_list:
        for j in i:
            for k in j:
                #处理美式类型数字
                if type(k) is str:
                    k = AmericanNumber2NormalNumber(k)
                res.append(abs(k))
    return np.array(res)

np.set_printoptions(suppress=True) #设置numpy为不用科学计数法输出
# 计算银行统计表的应收账款
def CalAllRecieveMoney():
    fundsOnAccountSheet = None
    try:
        fundsOnAccountSheet = pd.read_excel("日新销售台账新.xlsx", header=None, sheet_name="2021台账")
    except:
        print("未成功打开 日新销售台账新.xlsx 或 2021台账表格 程序异常退出")
        time.sleep(3)
        exit(-1)

    # 取出所需列数据
    sendOrRecieve = np.array(fundsOnAccountSheet.loc[1:, 0])
    sendDate = np.array(fundsOnAccountSheet.loc[1:, 6])
    recieveMoneyDate = np.array(fundsOnAccountSheet.loc[1:, 7])
    preRecieveMoney = np.array(fundsOnAccountSheet.loc[1:, 45])
    recieve = np.array(fundsOnAccountSheet.loc[1:, 22])
    # 创建改变的应收货款的位置和值的字典
    allRecieveMoney = 0
    # 计算行数
    count = 0
    # 选择发，有收/发货日期，没有收/付款日期，=销售收入-预收货款
    for i in range(len(sendOrRecieve)):
        if sendOrRecieve[i] == '发':
            # 找到不为空的发货日期
            if not sendDate[i] is np.nan:
                # 找到为空的收款日期:
                if recieveMoneyDate[i] is np.nan:
                    # 为空则计算 应收货款=销售货款-预付货款
                    # 判断预付货款是否为none:
                    if preRecieveMoney[i] is np.nan:
                        recieveMoney = ceil(recieve[i]) - 0
                        print("我是第" + str(i + 2) + "行的应收货款:我等于{}-0".format(ceil(recieve[i])))
                    else:
                        recieveMoney = ceil(recieve[i]) - preRecieveMoney[i]
                        print("我是第" + str(i + 2) + "行的应收货款:我等于{}-{}".format(ceil(recieve[i]), preRecieveMoney[i]))
                    allRecieveMoney += recieveMoney
                    count += 1

    print("共计行数:",count)
    print("应收货款为:", allRecieveMoney)
    return allRecieveMoney

# 计算银行统计表的应付账款
def CalPayment():
    fundsOnAccountSheet = None
    try:
        fundsOnAccountSheet = pd.read_excel("日新销售台账新.xlsx", header=None, sheet_name="2021台账")
    except:
        print("未成功打开 日新销售台账新.xlsx 或 2021台账表格 程序异常退出")
        time.sleep(3)
        exit(-1)

    # 取出所需列数据
    sendOrRecieve = np.array(fundsOnAccountSheet.loc[1:, 0])
    sendDate = np.array(fundsOnAccountSheet.loc[1:, 6])
    recieveMoneyDate = np.array(fundsOnAccountSheet.loc[1:, 7])
    payment = np.array(fundsOnAccountSheet.loc[1:, 9])
    cost = np.array(fundsOnAccountSheet.loc[1:20])
    allPayment = 0
    count = 0
    # 选择发，有收/发货日期，没有收/付款日期，应付货款=销售收入-预收货款
    for i in range(len(sendOrRecieve)):
        if sendOrRecieve[i] == '收':
            # 找到不为空的发货日期
            if not sendDate[i] is np.nan:
                # 找到为空的收款日期:
                if recieveMoneyDate[i] is np.nan:
                    # 判断若应付货款为none
                    if payment[i] is np.nan:
                        payment[i] = ceil(cost[i])
                        print("我是第", str(i + 2), "行的应付账款，我当前等于本行的采购成本", str(cost[i]))
                    else:
                        print("我是第", str(i + 2), "行的应付账款，我当前等于", payment[i])
                    allPayment += ceil(payment[i])
                    count += 1
    print("应付账款为:", allPayment)
    print("统计行数：",count)
    return allPayment

#计算三家银行的货款收入
def CalRecievePayment():
    # 变量声明部分
    ICIBSheetOfBandExcel = None
    ChinaBandSheetOfBandExcel = None
    YongHengSheetOfBandExcel = None

    res_ICIB_summary_send = None
    res_ICIB_summary_recieve = None
    res_ChinaBand_summary = None
    res_YongHeng_summary_send = None
    res_YongHeng_summary_receive = None
    # 读取对应表格数据
    try:
        ICIBSheetOfBandExcel = pd.read_excel("日新银行统计表2021.xlsx", header=1, sheet_name="工行收支表")
        ChinaBandSheetOfBandExcel = pd.read_excel("日新银行统计表2021.xlsx", header=1, sheet_name="中行收支表")
        YongHengSheetOfBandExcel = pd.read_excel("日新银行统计表2021.xlsx", header=2, sheet_name="华侨永亨")
    except:
        if fundsOnAccountSheet is None:
            print("未成功打开 日新销售台账新.xlsx 或 2021台账表格 程序异常退出")
        elif ICIBSheetOfBandExcel is None:
            print("未成功打开 日新银行统计表2021.xlsx 或 工行收支表 程序异常退出")
        elif ChinaBandSheetOfBandExcel is None:
            print("未成功打开 中行收支表 程序异常退出")
        elif YongHengSheetOfBandExcel is None:
            print("未成功打开 华侨永亨 程序异常退出")
        time.sleep(5)
        exit(-1)
    print("所需数据表格读取成功")

    # 筛选所需数据
    try:
        # 筛选工行表中 "摘要"为货款并且是"转出金额"不为零的"转出金额"
        res_ICIB_summary_send = np.array(ICIBSheetOfBandExcel.loc[(ICIBSheetOfBandExcel['摘要'] == "货款") & (
                ICIBSheetOfBandExcel["转出金额"] != 0.0), ["转出金额"]])
        # 筛选工行表中 "摘要"为货款并且是"转入金额"不为零的"转入金额"
        res_ICIB_summary_recieve = np.array(ICIBSheetOfBandExcel.loc[(ICIBSheetOfBandExcel['摘要'] == "货款") & (
                ICIBSheetOfBandExcel["转入金额"] != 0.0), ["转入金额"]])
        # 筛选中行表中 “交易附言[ Remark ]”为“货款”的交易金额
        res_ChinaBand_summary = np.array(ChinaBandSheetOfBandExcel.loc[
                                             (ChinaBandSheetOfBandExcel['交易附言[ Remark ]'] == "货款(网银转账，有误即退)") | (
                                                     ChinaBandSheetOfBandExcel['交易附言[ Remark ]'] == '货款'), [
                                                 "交易金额[ Trade Amount ]"]])
        # 筛选华侨永亨表中 "摘要"为"外币转账支出"或摘要为"SWIFT 转账支出"的"支出"
        res_YongHeng_summary_send = np.array(YongHengSheetOfBandExcel.loc[
                                                 (YongHengSheetOfBandExcel["摘要"] == "外币转账支出") | (
                                                         YongHengSheetOfBandExcel["摘要"] == "SWIFT 转账支出"), ["支出"]])
        # 筛选华侨永亨表中 "摘要"为"SWIFT 转账收入"的"收入"
        res_YongHeng_summary_receive = np.array(
            YongHengSheetOfBandExcel.loc[(YongHengSheetOfBandExcel["摘要"] == "SWIFT 转账收入"), ["收入"]])
    except:
        # 处理读取
        print('筛选步骤出错，请检查筛选条件对应的列名是否变化,筛选条件为:\n'
              '筛选工行表中 "摘要"为货款并且是"转出金额"不为零的"转出金额"\n'
              '筛选工行表中 "摘要"为货款并且是"转入金额"不为零的"转入金额"\n'
              '筛选中行表中 “交易附言[ Remark ]”为“货款”的"交易金额[ Trade Amount ]"\n'
              '筛选华侨永亨表中 "摘要"为"外币转账支出"或摘要为"SWIFT 转账支出"的"支出"\n'
              '筛选华侨永亨表中 "摘要"为"SWIFT 转账收入"的"收入"')
        time.sleep(5)
        exit(-1)

    # 整理中行表的货款收入和支出数据
    res_ChinaBand_summary_send = []
    res_ChinaBand_summary_recieve = []
    for i in res_ChinaBand_summary:
        # 交易金额为负,则加入货款支出,为正数则加入货款收入
        if i[0] < 0.0:
            res_ChinaBand_summary_send.append(i)
        elif i[0] > 0.0:
            res_ChinaBand_summary_recieve.append(i)

    # 整合支出货款数组
    res_send_list = np.array([res_ICIB_summary_send,
                              res_ChinaBand_summary_send,
                              res_YongHeng_summary_send], dtype=object)
    # 整合收入货款数组
    res_recieve_list = np.array([res_ChinaBand_summary_recieve,
                                 res_ICIB_summary_recieve,
                                 res_YongHeng_summary_receive], dtype=object)

    # 对支出货款列表进行数据格式化处理，生成一维的数组
    res_send_list_proccessed = np.array(dataProccess(res_send_list))
    res_recieve_proccessed = np.array(dataProccess(res_recieve_list))

    sum_send = np.around(np.sum(res_send_list_proccessed), 2)
    sum_recieve = np.around(np.sum(res_recieve_proccessed), 2)

    return [sum_send,sum_recieve]


def main():


    # 计算应收钱款
    allRecieveMoney = CalAllRecieveMoney()
    # 计算应付账款
    shouldSendMoney = CalPayment()
    # 计算货款收入
    allPayment = CalRecievePayment()
    sum_send = allPayment[0]
    sum_recieve = allPayment[1]

    # 计算货款支出
    # 初始化写入工作
    app = xw.App(visible=True, add_book=False)  # 程序可见，只打开不新建工作薄
    # app.display_alerts = False  # 警告关闭
    # app.screen_updating = False  # 屏幕更新关闭


    statisticExcel = None
    # 获取银行统计表
    try:
        statisticExcel = app.books.open("日新统计表2021实时.xlsx")
        bandSheet = statisticExcel.sheets["银行统计"]
        funds_on_account_sheet = statisticExcel.sheets["台账统计"]

        # 更新统计表的应付账款 E3 应收账款M3
        bandSheet.range('E3').value = allRecieveMoney
        bandSheet.range('F3').value = shouldSendMoney
        print("银行统计表更新完毕")

        funds_on_account_sheet.range('M2').value = sum_recieve
        funds_on_account_sheet.range('N2').value = sum_send
        print("台账表格更新完毕")
    except:
        print("打开 日新统计表2021实时.xlsx 失败,写入数据失败")
        exit(-2)
    finally:
        statisticExcel.save()  # 保存文件
        statisticExcel.close()  # 关闭文件
        app.quit()  # 关闭程序
        time.sleep(4)

main()