import sys
import pandas as pd
import os

fromDate = sys.argv[1]
toDate = sys.argv[2]

if(os.path.exists('/Users/arungeorge/Documents/Personal/PAGeorgeCo/2022/summary.xlsx')):
    os.remove('/Users/arungeorge/Documents/Personal/PAGeorgeCo/2022/summary.xlsx')

companyList = ['PAG-EKM','PAG-ALPY','Pages-Ekm','PAG-Calicut']
writer = pd.ExcelWriter('/Users/arungeorge/Documents/Personal/PAGeorgeCo/2022/summary.xlsx', engine='xlsxwriter')

for singleCompany in companyList:
    
    daybookDF = pd.read_excel('/Users/arungeorge/Documents/Personal/PAGeorgeCo/2022/PAG_dayBook.xlsx',sheet_name=singleCompany)

    #combine the sales for ITC (pages and Agencies)
    if(singleCompany == "Pages-Ekm"):
        daybookDFAgy = pd.read_excel('/Users/arungeorge/Documents/Personal/PAGeorgeCo/2022/PAG_dayBook.xlsx',sheet_name="PAG-agy-EKM")

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
        totalSale.append(round(daySale,2))
        totalProfit.append(round(dayProfit,2))

        if(singleCompany == "Pages-Ekm"):
            filter_dayBookAgy = daybookDFAgy[(daybookDFAgy['Date']>= singleDate) & (daybookDFAgy['Date'] <= singleDate)]
            daySale = filter_dayBookAgy['Amount'].sum()
            dayProfit = filter_dayBookAgy['Profit'].sum()

            days.append(singleDate.strftime("%m-%d-%y"))
            totalSale.append(round(daySale,2))
            totalProfit.append(round(dayProfit,2))

    consolidated_df = pd.DataFrame(list(zip(days,totalSale,totalProfit)), columns=columns)
    
    if(singleCompany == "Pages-Ekm"):
        consolidated_df.to_excel(writer,sheet_name="ITC", index=False)
    else:
        consolidated_df.to_excel(writer,sheet_name=singleCompany, index=False)

writer.save()

print('processing completed for all companies from ' + fromDate + ' to ' + toDate)