#coding=utf-8
import os
import pandas as pd
from scipy import stats
pd.set_option('max_rows', None)
#取财务报表每个单元格的首位数字
def Get_First_num(input):
    output = 0
    try:
        if input != 'nan' and len(input) > 0:
            input = float(input.replace(',', '')) #去除數字27,712,000,000的逗號
            input = abs(input)                    #將負數轉成正樹
            if input > 20 or input<-20:           #剔除過小值,如eps
                output = int(str(input)[0])
            else:
                pass
    except Exception as e:
        return
    return output
#計算各表的數字的數目
def Count_num(num_count, input):
    output = Get_First_num(input)
    if output == 1:         #如果字母為1
        num_count[0] += 1   #加一次
    elif output == 2:       #如果字母為2
        num_count[1] += 1   #加一次
    elif output == 3:       #如果字母為3
        num_count[2] += 1   #加一次
    elif output == 4:       #如果字母為4
        num_count[3] += 1   #加一次
    elif output == 5:       #如果字母為5
        num_count[4] += 1   #加一次
    elif output == 6:       #如果字母為6
        num_count[5] += 1   #加一次
    elif output == 7:       #如果字母為7
        num_count[6] += 1   #加一次
    elif output == 8:       #如果字母為8
        num_count[7] += 1   #加一次
    elif output == 9:       #如果字母為9
        num_count[8] += 1   #加一次
def compute_num(root,file):
    flow_num=[]
    #try:
    path_cash = os.path.join(root, file)
    cash_flow = pd.read_excel(path_cash, delimiter='\t')
    cash_flow = cash_flow[3:]
    cash_flow = cash_flow[:-10]
    for i in range(0, len(cash_flow.columns)):
        cash_flow.rename(columns={cash_flow.columns[i]:str(i)},inplace=True)
    for i in range(0, len(cash_flow.columns)):
        data = cash_flow.pop(str(i))
        num = [0, 0, 0, 0, 0, 0, 0, 0, 0]
        for t in data:
            Count_num(num, str(t))
        flow_num.append(num)
    # except Exception as e:
    #      #print(str(e))
    return flow_num
def Get_actual(season,cash_flow_num,profitsheet_num,Assetsheet_num):
    actual_num=[]                      #三張表首字母次數加總的結果
    for i in range(0, season):
        list1 = []
        for t in range(0,9):
            list1.append(profitsheet_num[i][t] + cash_flow_num[i][t]+ Assetsheet_num[i][t])
        actual_num.append(list1)
    return actual_num
def Get_expected(sum_sheet):#sum_sheet為三張表,其會計科目的首字母總數
    expected_benford=[]
    for i in sum_sheet:
        benford = [0.301, 0.1761, 0.1249, 0.0969, 0.0792, 0.0669,0.058, 0.0512, 0.0458]#本福特定律中各數的機率
        list1=[round(x*i,2)  for x in benford]
        expected_benford.append(list1)
    return expected_benford
def Chi_test(season,actual_num,expected_benford,P):
    finally_answer=[]
    for i in range(0, season):
        if stats.chisquare(actual_num[i],
                           f_exp=expected_benford[i])[1] > P:
            finally_answer.append(0)
        else:
            finally_answer.append(1)
    return finally_answer
#=======自訂選取區================================== #
season=4 #取選季節長度,若最新數據為2017第一季,則會往前抓取21季的數據
P=0.1     #90%信心水準
#=======列表存放區================================== #
profitsheet_num=[] #存放利潤表的首字母分佈   eg:[  9,  5,  4,  3,  2,  1,  5,  6,  7]
cash_flow_num=[]   #存放現金流量表的首字母分佈   [  3,  5,  6,  7,  8,  9,  2,  1,  4]
Assetsheet_num=[]  #存放資產負債表的首字母分佈   [  2,  3,  5,  6,  8,  9,  3,  2,  3]
actual_num=[]      #存放三表加總的首字母分      [ 14, 13, 15, 16, 18, 19, 10,  9, 14]
sum_sheet=[]       #存放三者加總的首字母分佈    [128,128,128,128,128,128,128,128,128]
name=[]            #上市公司的代碼
cannot_reject=[] #不拒絕次數加總
can_reject=[]    #拒絕次數加總
company=[]       #上市公司的代碼
#=======資料存放區================================== #
path ='/Users/Benfordslaw/data_test'#文件放置地址
#=======取出所有代碼================================ #
for root, dirs, files in os.walk(path):
    for file in files:
        name.append(file[0:6])
name_dict = {}.fromkeys(name).keys()
name_list =[]
for i in name_dict:
    name_list.append(i)
print(name_list)
name_list.remove('.DS_St')     #只有mac系統有 .DS_St
for root, dir, files in os.walk(path):
    for name in name_list:
        for file in files:
            if file[0:6] == str(name):
                if file[6:12]=='profit':
                    profitsheet_num=compute_num(root,file)
                elif file[6:10]=='cash':
                    cash_flow_num=compute_num(root,file)
                elif file[6:11]=='asset':
                    Assetsheet_num=compute_num(root,file)
                else:
                    pass
        print(name)
        print(len(cash_flow_num),len(profitsheet_num),len(Assetsheet_num))
        actual_num=Get_actual(season,cash_flow_num,profitsheet_num,Assetsheet_num)
        for i in range(0,season):
            sum_sheet.append(sum(actual_num[i]))
        expected_benford=Get_expected(sum_sheet)
        finally_answer=Chi_test(season,actual_num,expected_benford,P)
        sum_sheet.clear()
        actual_num.clear()
        expected_benford.clear()
        company.append(name)
        cannot_reject.append(finally_answer.count(0))
        can_reject.append(finally_answer.count(1))
dict = {"B_cannot_reject": cannot_reject,
        "can_reject":can_reject,
        "A_company": company,
        }
df_final = pd.DataFrame(dict)
print(df_final.sort_values("B_cannot_reject",ascending=False))
print(df_final["B_cannot_reject"].mean())
print(df_final["can_reject"].mean())
writer = pd.ExcelWriter('output.xlsx')
df_final.to_excel(writer,'Sheet1')
writer.save()