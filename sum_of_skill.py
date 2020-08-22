import pandas as ps
from datetime import datetime, date

excel_data = ps.read_excel('Site_Capacity.xlsx', sheet_name = 'Sheet1')
skills = excel_data['Skill'].tolist()

while 1:
    # enter skill and check if the skill
    # doesn't exist in the table else enter again
    print("Please enter a skill:")
    skill = input() 
    if skill not in skills:
        continue

    # enter year and check if the year
    # makes sense else enter again
    while 1:
        print("Please enter a year:")
        try:
            year = int(input())
        except:
            continue
        if 1900 <= year <= date.today().year + 2:
            break

    # enter month and check if the month
    # makes sense else enter again
    while 1:
        print("Please enter a month:")
        try:
            month = int(input())
        except:
            continue
        if  1 <= month <= 12:
            break
    
    # get the corresponding value for the specified skill
    # from the table
    # if all inputs value are correct but nothing in the table
    # enter input again
    try:
        excel_list = excel_data.columns[3:].tolist()
        index = excel_list.index(datetime(year = year, month = month, day = 1))
        sum = 0
        excel_data = excel_data.fillna(0)
        for i in excel_data.index:
            if excel_data["Skill"][i] == skill:
                sum = sum + excel_data[excel_data.columns[index + 3]][i]
        print(sum)         
    except:
        continue
        
    # possibility to exit
    to_do = ""
    while to_do != "continue" and to_do != "exit":
        print("Continue or exit?")
        to_do = input().lower()
    
    if to_do == "continue":
        continue
    else:
        break




