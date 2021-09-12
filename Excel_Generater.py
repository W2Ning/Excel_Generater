# -*- coding:utf-8 -*-
import os
import time
import xlwt
import random
import timeit 
import datetime
import calendar
from xlwt import Column
from chinese_calendar import is_workday


style1 = xlwt.XFStyle()
style2 = xlwt.XFStyle()
style3 = xlwt.XFStyle()

font = xlwt.Font()
font.name = '宋体'
font.height = 400
font.bold = False

style1.font = font
style2.font = font
style3.font= font

print('''

$$$$$$$$\                               $$\                                             
$$  _____|                              $$ |                                            
$$ |      $$\   $$\  $$$$$$$\  $$$$$$\  $$ |                                            
$$$$$\    \$$\ $$  |$$  _____|$$  __$$\ $$ |                                            
$$  __|    \$$$$  / $$ /      $$$$$$$$ |$$ |                                            
$$ |       $$  $$<  $$ |      $$   ____|$$ |                                            
$$$$$$$$\ $$  /\$$\ \$$$$$$$\ \$$$$$$$\ $$ |                                            
\________|\__/  \__| \_______| \_______|\__|                                            
                                                                                        
                                                                                        
                                                                                        
 $$$$$$\                                                    $$\                         
$$  __$$\                                                   $$ |                        
$$ /  \__| $$$$$$\  $$$$$$$\   $$$$$$\   $$$$$$\  $$$$$$\ $$$$$$\    $$$$$$\   $$$$$$\  
$$ |$$$$\ $$  __$$\ $$  __$$\ $$  __$$\ $$  __$$\ \____$$\\\_$$  _|  $$  __$$\ $$  __$$\ 
$$ |\_$$ |$$$$$$$$ |$$ |  $$ |$$$$$$$$ |$$ |  \__|$$$$$$$ | $$ |    $$$$$$$$ |$$ |  \__|
$$ |  $$ |$$   ____|$$ |  $$ |$$   ____|$$ |     $$  __$$ | $$ |$$\ $$   ____|$$ |      
\$$$$$$  |\$$$$$$$\ $$ |  $$ |\$$$$$$$\ $$ |     \$$$$$$$ | \$$$$  |\$$$$$$$\ $$ |      
 \______/  \_______|\__|  \__| \_______|\__|      \_______|  \____/  \_______|\__|      
                                                                                        
''')

year  = int(input(">>>>>>输入年: "))
month = int(input(">>>>>>输入月: "))
low   = int(input(">>>>>>输入最小值: "))
high  = int(input(">>>>>>输入最大值: "))

last_day = calendar.mdays[month]

workdays = []

for i in range(1,last_day+1):
    a = datetime.datetime(year,month,i)
    b = is_workday(a)
    if b:
        d = str(a)
        e = d.split(" ", 1)
        f = e[0]
        workdays.append(f)


count_days = len(workdays)
sum_1 = 0
sum_2 = 0
sum_3 = 0




new_workbook = xlwt.Workbook()
worksheet = new_workbook.add_sheet("sheet1")
no_1_col = worksheet.col(0)
no_1_col.width  = 512*20

no_2_col = worksheet.col(1)
no_2_col.width  = 256*20

no_3_col = worksheet.col(2)
no_3_col.width  = 256*20

no_5_col = worksheet.col(4)
no_5_col.width  = 512*20

no_6_col = worksheet.col(5)
no_6_col.width  = 256*20

no_7_col = worksheet.col(6)
no_7_col.width  = 256*20

no_8_col = worksheet.col(7)
no_8_col.width  = 512*20


worksheet.write(0,2,"金额",style3)
worksheet.write(0,6,"金额",style3)
worksheet.write(0,7,"总额",style3)

for i in range(0,count_days):
    worksheet.write(i+1, 0, workdays[i],style1)
    worksheet.write(i+1, 4, workdays[i],style1)
    a = str(random.randint(29,59))
    time_one = "7:"+ a
    worksheet.write(i+1, 1, time_one,style1)

    b = random.randint(low,high)
    money_one = str(b) 
    worksheet.write(i+1, 2, money_one,style3)

    c = str(random.randint(10,31))
    time_two = "18:"+ c
    worksheet.write(i+1, 5, time_two,style1)

    d = random.randint(low,high)
    money_two = str(d)
    worksheet.write(i+1, 6, money_two,style3)

    sum_1 += b
    sum_2 += d

    sum_3 = sum_1 + sum_2

worksheet.write(count_days+1, 2, sum_1, style1)
worksheet.write(count_days+1, 6, sum_2, style1)
worksheet.write(count_days+1, 7, sum_3, style3)



file_name = str(year) + "年_" +str(month) + "月.xls"

new_workbook.save(file_name)


# print(str(month) + "月份的工作日有" + str(len(workdays)) + "天")

print('''

 $$$$$$\                                                              
$$  __$$\                                                             
$$ /  \__|$$\   $$\  $$$$$$$\  $$$$$$$\  $$$$$$\   $$$$$$$\  $$$$$$$\ 
\$$$$$$\  $$ |  $$ |$$  _____|$$  _____|$$  __$$\ $$  _____|$$  _____|
 \____$$\ $$ |  $$ |$$ /      $$ /      $$$$$$$$ |\$$$$$$\  \$$$$$$\  
$$\   $$ |$$ |  $$ |$$ |      $$ |      $$   ____| \____$$\  \____$$\ 
\$$$$$$  |\$$$$$$  |\$$$$$$$\ \$$$$$$$\ \$$$$$$$\ $$$$$$$  |$$$$$$$  |
 \______/  \______/  \_______| \_______| \_______|\_______/ \_______/ 
                                                                      
                                                                      
                                                                      
''')


os.system(file_name) 
