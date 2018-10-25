import xlrd

all_data = [] #汇总的数据
fund_list = [] #for different fund
issue_type_list = [] # for issue type list
issue_country_list = [] # for issue country list
calculation_fund = {}
calculation_issue_type = {}
calculation_issue_country = {}


class Data:
    investment_fund = ""
    duration = 0
    weight = 0
    issue_type = ""
    issue_country = ""
    calculation = 0


def read_xlsfile(xls_filename):
    readbook = xlrd.open_workbook(xls_filename)
    sheet = readbook.sheet_by_name("Sheet1")  # 名字的方式
    nrows = sheet.nrows  # 行
    ncols = sheet.ncols  # 列
    for i in range(1,nrows):
        data = Data()
        data.investment_fund = sheet.cell_value(i, 0)
        data.duration = sheet.cell_value(i, 6)
        data.weight = sheet.cell_value(i, 7)
        data.issue_type = sheet.cell_value(i, 8)
        data.issue_country = sheet.cell_value(i,9)
        data.calculation = data.weight * data.duration
        all_data.append(data)


read_xlsfile("FOR VBA CODING.xlsx")
for i in range(len(all_data)):
    if all_data[i].investment_fund not in fund_list:
        fund_list.append(all_data[i].investment_fund)

    if all_data[i].issue_type not in issue_type_list:
        issue_type_list.append(all_data[i].issue_type)

    if all_data[i].issue_country not in issue_country_list:
        issue_country_list.append(all_data[i].issue_country)

for i in range(len(fund_list)):
    calculation_fund[fund_list[i]] = 0
    for j in range(len(all_data)):
        if all_data[j].investment_fund == fund_list[i]:
            temp_cal = calculation_fund[fund_list[i]]
            calculation_fund[fund_list[i]] = all_data[j].calculation + temp_cal

for i in range(len(issue_type_list)):
    calculation_issue_type[issue_type_list[i]] = 0
    for j in range(len(all_data)):
        if all_data[j].issue_type == issue_type_list[i]:
            temp_cal = calculation_issue_type[issue_type_list[i]]
            calculation_issue_type[issue_type_list[i]] = all_data[j].calculation + temp_cal

for i in range(len(issue_country_list)):
    calculation_issue_country[issue_country_list[i]] = 0
    for j in range(len(all_data)):
        if all_data[j].issue_country == issue_country_list[i]:
            temp_cal = calculation_issue_country[issue_country_list[i]]
            calculation_issue_country[issue_country_list[i]] = all_data[j].calculation + temp_cal


print(calculation_fund)
print(calculation_issue_type)
print(calculation_issue_country)
