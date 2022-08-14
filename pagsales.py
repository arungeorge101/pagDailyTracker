import sys
import pandas as pd
import os

fromDate = sys.argv[1]
toDate = sys.argv[2]

if(os.path.exists('/Users/arungeorge/Documents/Personal/PAGeorgeCo/2022/summary.xlsx')):
    os.remove('/Users/arungeorge/Documents/Personal/PAGeorgeCo/2022/summary.xlsx')

companyList = ['PAG-EKM','PAG-ALPY','Pages-Ekm','PAG-agy-EKM','PAG-Calicut']
writer = pd.ExcelWriter('/Users/arungeorge/Documents/Personal/PAGeorgeCo/2022/summary.xlsx', engine='xlsxwriter')

for singleCompany in companyList:
    daybookDF = pd.read_excel('/Users/arungeorge/Documents/Personal/PAGeorgeCo/2022/PAG_dayBook.xlsx',sheet_name=singleCompany)

    dateRange = pd.date_range(fromDate, toDate)

    days = []
    totalSale = []
    totalProfit =[]
    columns = ['day', 'amount', 'profit']

    for singleDate in dateRange:
        filter_dayBook = daybookDF[(daybookDF['Date']>= singleDate) & (daybookDF['Date'] <= singleDate)]
        daySale = filter_dayBook['Amount'].sum()
        dayProfit = filter_dayBook['Profit'].sum()

        days.append(singleDate.strftime("%m-%d-%y"))
        totalSale.append("{:.2f}".format(daySale))
        totalProfit.append("{:.2f}".format(dayProfit))

    consolidated_df = pd.DataFrame(list(zip(days,totalSale,totalProfit)), columns=columns)
    consolidated_df.to_excel(writer,sheet_name=singleCompany)

writer.save()

print('')


