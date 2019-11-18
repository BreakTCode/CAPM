import pandas as pd
import numpy as np
from scipy import stats
import xlsxwriter as xlw

#Variables
count = 1
MRGSPC = [0]
MRMSFT = [0]

#Read the different files for the data.
sheet1 = pd.read_excel('^GSPC (Full).xlsx', header=0)
sheet2 = pd.read_excel('MSFT (Full).xlsx', header=0)

#Create the new file, worksheet and Date format.
writer = xlw.Workbook('Microsoft CAPM.xlsx')
DateF = writer.add_format({'num_format': 'mm/dd/yy'})
worksheet1 = writer.add_worksheet('CAPM')

#Creates the header row and first values
MyHeader = ['Date', 'GSPC', 'Monthly Return', 'MSFT', 'Monthly Return']
worksheet1.write_row('A1', MyHeader)

MyRow = [sheet1['Adj Close'][0], "", sheet2['Adj Close'][0]]
worksheet1.write('A2', sheet1['Date'][0], DateF)
worksheet1.write_row('B2', MyRow)

#Loops through both files adding the date, the adjusted close value of the GSPC and MSFT, then the caluculated change.
while count < sheet1.shape[0]:
    MyRow = [sheet1['Adj Close'][count], "{:.2%}".format(sheet1['Adj Close'][count] / sheet1['Adj Close'][count - 1] - 1), sheet2['Adj Close'][count], "{:.2%}".format(sheet2['Adj Close'][count]/sheet2['Adj Close'][count - 1] - 1)]
    if MRMSFT[0] == 0:     
        MRGSPC[0] = sheet1['Adj Close'][count] / sheet1['Adj Close'][0]
        MRMSFT[0] = sheet2['Adj Close'][count] / sheet2['Adj Close'][0]
    else:
        MRGSPC.append(sheet1['Adj Close'][count] / sheet1['Adj Close'][count - 1])
        MRMSFT.append(sheet2['Adj Close'][count] / sheet2['Adj Close'][count - 1])
    worksheet1.write('A' + str(count + 1), sheet1['Date'][count], DateF)
    worksheet1.write_row('B' + str(count + 1), MyRow)    
    count += 1
    


#Writes the Average Close Value
count += 2
worksheet1.write('A' + str(count), "Average")
worksheet1.write('C' + str(count), stats.gmean(MRGSPC) - 1)
worksheet1.write('E' + str(count), stats.gmean(MRMSFT) - 1)

#Writes the Annual Average Close Value
count += 1
worksheet1.write('A' + str(count), "Average Annual")
worksheet1.write('C' + str(count), (pow(stats.gmean(MRGSPC), 12) - 1))
worksheet1.write('E' + str(count), (pow(stats.gmean(MRMSFT), 12) - 1))

#Writes the Variance
count += 1
worksheet1.write('A' + str(count), "Variance")
worksheet1.write('C' + str(count), np.var(MRGSPC,ddof=1))
worksheet1.write('E' + str(count), np.var(MRMSFT,ddof=1))

#Writes the Monthly Standard Deviation
count += 1
worksheet1.write('A' + str(count), "Monthly Standard Deviation")
worksheet1.write('C' + str(count), np.std(MRGSPC,ddof=1))
worksheet1.write('E' + str(count), np.std(MRMSFT,ddof=1))

#Writes the Annual Standard Deviation
count += 1
worksheet1.write('A' + str(count), "Annual Standard Deviation")
worksheet1.write('C' + str(count), pow(np.var(MRGSPC,ddof=1) * 12, .5) )
worksheet1.write('E' + str(count), pow(np.var(MRMSFT,ddof=1) * 12, .5) )

#Writes the Covariance
count += 1
worksheet1.write('A' + str(count), "Covariance")
worksheet1.write('C' + str(count), np.cov(MRGSPC,MRMSFT, ddof=0)[0][1])

#Writes the BETA
count += 1
worksheet1.write('A' + str(count), "BETA")
worksheet1.write('C' + str(count), np.cov(MRGSPC,MRMSFT, ddof=0)[0][1]/np.var(MRGSPC,ddof=1))

writer.close()
