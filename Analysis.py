'''
Shahzeb Khalid

'''
import statistics as stats
import xlrd
import numpy as num
import scipy
from scipy.stats import skew
import math
import matplotlib.pyplot as plt

# _-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_

# This is define a funtion for find Outliers
outliers = []
def detect_outlier(data):
    checkpoint = 3
    mean = num.mean(data)
    std =num.std(data)
    for y in data:
        z_score= (y - mean)/std
        if num.abs(z_score) > checkpoint:
            outliers.append(y)
    return outliers

# _-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_

# inisalizing file to vaiable
filename = "HmEqty.xlsx"
# This line is used to open the file and passing the info and storing it in the variable data
data = xlrd.open_workbook(filename)

# copy = pd.read_csv(gapminder_csv_url)

sheet = data.sheet_by_index(0)
# These variables are used to store the columns from the data set from the excel sheet.
loan, mortdue, value, yol, derog, delinq, clage, ninq, clno, debtinc = ([] for i in range(10))
# These declared vaiables are used for storing the data passed form the Outliers Funtion.
loan_outlier, mort_outlier, value_outlier, yol_outlier, derog_outlier = ([] for i in range(5))
delinq_outlier,clage_outlier, ninq_outlier, clno_outlier, debtinc_outlier = ([] for i in range(5))

# ∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞
''' This is loop is used to iterate thorugh the columns from the sheet and
    appends the colum one by one from the sheet when it called by hardcoding col numerbs '''
for i in range(sheet.nrows):
   if (i > 0):
       loan.append(sheet.cell_value(i,1))
       mortdue.append(sheet.cell_value(i,2))
       value.append(sheet.cell_value(i,3))
       yol.append(sheet.cell_value(i,6))
       derog.append(sheet.cell_value(i,7))
       delinq.append(sheet.cell_value(i,8))
       clage.append(sheet.cell_value(i,9))
       ninq.append(sheet.cell_value(i,10))
       clno.append(sheet.cell_value(i,11))
       debtinc.append(sheet.cell_value(i,12))



#_-_-_-_-_-_-_-_-_-_-_-_-_-_- Loan _-_-_-_-_-_-_-_-_-_-_-_-_-_-
loan_mean = num.mean(loan)
loan_MID = num.median(loan)
loan_SD = num.std(loan)
loan_var = num.var(loan)
loan_skew = skew(loan)

loan_min_qurt= num.quantile(loan, 0.25)
loan_max_qurt = num.quantile(loan, 0.75)

for i in range(len(loan)):
    min_loan = min(loan)
    max_loan = max(loan)
    loan_ran = max_loan - min_loan

# loan_cor = num.corrcoef(loan,value)

print("\n")
print(25*' ',"OUTLIERS")
loan_outlier = detect_outlier(loan)
print("Loan Outlier", 8*' ', loan_outlier)
outliers.clear()

#_-_-_-_-_-_-_-_-_-_-_-_-_-_- Mortdue _-_-_-_-_-_-_-_-_-_-_-_-_-_-
mort_mean = num.mean(mortdue)
mort_MID = num.median(mortdue)
mort_SD = num.std(mortdue)
mort_var = num.var(mortdue)
mort_skew = skew(mortdue)

mort_min_qurt= num.quantile(mortdue, 0.25)
mort_max_qurt = num.quantile(mortdue, 0.75)

for i in range(len(mortdue)):
    min_loan = min(mortdue)
    max_loan = max(mortdue)
    mort_ran = max_loan - min_loan

mort_outlier = detect_outlier(mortdue)
print("Mortdue Outlier", 5*' ', mort_outlier)
outliers.clear()

# _-_-_-_-_-_-_-_-_-_-_-_-_-_- Value _-_-_-_-_-_-_-_-_-_-_-_-_-_-
value_mean = num.mean(value)
value_MID = num.median(value)
value_SD = num.std(value)
value_var = num.var(value)
value_skew = skew(value)

value_min_qurt= num.quantile(value, 0.25)
value_max_qurt = num.quantile(value, 0.75)

for i in range(len(value)):
    min_loan = min(value)
    max_loan = max(value)
    value_ran = max_loan - min_loan

value_outlier = detect_outlier(value)
print("Value Outlier",  7*' ',value_outlier)
outliers.clear()

# _-_-_-_-_-_-_-_-_-_-_-_-_-_- YoJ _-_-_-_-_-_-_-_-_-_-_-_-_-_-


# this loop and chekc if there's any white space and reaplce it with zero
for j in range(len(yol)):
    if (yol[j] == '' or yol[j] == ' '):
        yol[j] = 0

yol_mean = num.mean(yol)

yol_MID = num.median(yol)
yol_SD = num.std(yol)
yol_var = num.var(yol)
yol_skew = skew(yol)

yol_min_qurt= num.quantile(yol, 0.25)
yol_max_qurt = num.quantile(yol, 0.75)


#  This loop is used to find the range
for i in range(len(yol)):
    min_loan = min(yol)
    max_loan = max(yol)
    yol_ran = max_loan - min_loan

yol_outlier = detect_outlier(yol)
print("YOJ Outlier", 9*' ', yol_outlier)
outliers.clear()

# _-_-_-_-_-_-_-_-_-_-_-_-_-_- Derog _-_-_-_-_-_-_-_-_-_-_-_-_-_-
derog_mean = num.mean(derog)
derog_MID = num.median(derog)
derog_SD = num.std(derog)
derog_var = num.var(derog)
derog_skew = skew(derog)

derog_min_qurt= num.quantile(derog, 0.25)
derog_max_qurt = num.quantile(derog, 0.75)

for i in range(len(derog)):
    min_loan = min(derog)
    max_loan = max(derog)
    derog_ran = max_loan - min_loan

derog_outlier = detect_outlier(derog)
print("DEROG Outlier", 7*' ', derog_outlier)
outliers.clear()

# _-_-_-_-_-_-_-_-_-_-_-_-_-_- Delinq _-_-_-_-_-_-_-_-_-_-_-_-_-_-
delinq_mean = num.mean(delinq)
delinq_MID = num.median(delinq)
delinq_SD = num.std(delinq)
delinq_var = num.var(delinq)
delinq_skew = skew(delinq)

delinq_min_qurt= num.quantile(delinq, 0.25)
delinq_max_qurt = num.quantile(delinq, 0.75)

for i in range(len(delinq)):
    min_loan = min(delinq)
    max_loan = max(delinq)
    delinq_ran = max_loan - min_loan

delinq_outlier = detect_outlier(delinq)
print("DELINQ Outlier",  6*' ',delinq_outlier)
outliers.clear()

# _-_-_-_-_-_-_-_-_-_-_-_-_-_- Clage _-_-_-_-_-_-_-_-_-_-_-_-_-_-
clage_mean = num.mean(clage)
clage_MID = num.median(clage)
clage_SD = num.std(clage)
clage_var = num.var(clage)
clage_skew = skew(clage)

clage_min_qurt= num.quantile(clage, 0.25)
clage_max_qurt = num.quantile(clage, 0.75)

for i in range(len(clage)):
    min_loan = min(clage)
    max_loan = max(clage)
    clage_ran = max_loan - min_loan

clage_outlier = detect_outlier(clage)
print("CLAGE Outlier",  7*' ',clage_outlier)
outliers.clear()

# _-_-_-_-_-_-_-_-_-_-_-_-_-_- Ninq_-_-_-_-_-_-_-_-_-_-_-_-_-_-
ninq_mean = num.mean(ninq)
ninq_MID = num.median(ninq)
ninq_SD = num.std(ninq)
ninq_var = num.var(ninq)
ninq_skew = skew(ninq)

ninq_min_qurt= num.quantile(ninq, 0.25)
ninq_max_qurt = num.quantile(ninq, 0.75)

for i in range(len(ninq)):
    min_loan = min(ninq)
    max_loan = max(ninq)
    ninq_ran = max_loan - min_loan

ninq_outlier = detect_outlier(ninq)
print("NINQ Outlier", 8*' ', ninq_outlier)
outliers.clear()

# _-_-_-_-_-_-_-_-_-_-_-_-_-_- Clno _-_-_-_-_-_-_-_-_-_-_-_-_-_-
clno_mean = num.mean(clno)
clno_MID = num.median(clno)
clno_SD = num.std(clno)
clno_var = num.var(clno)
clno_skew = skew(clno)

clno_min_qurt= num.quantile(clno, 0.25)
clno_max_qurt = num.quantile(clno, 0.75)

for i in range(len(clno)):
    min_loan = min(clno)
    max_loan = max(clno)
    clno_ran = max_loan - min_loan

clno_outlier = detect_outlier(clno)
print("CLNO Outlier", 8*' ', clno_outlier)
outliers.clear()

# _-_-_-_-_-_-_-_-_-_-_-_-_-_- Debtinc _-_-_-_-_-_-_-_-_-_-_-_-_-
debtinc_mean = num.mean(debtinc)
debtinc_MID = num.median(debtinc)
debtinc_SD = num.std(debtinc)
debtinc_var = num.var(debtinc)
debtinc_skew = skew(debtinc)

debtinc_min_qurt= num.quantile(debtinc, 0.25)
debtinc_max_qurt = num.quantile(debtinc, 0.75)

for i in range(len(debtinc)):
    min_loan = min(debtinc)
    max_loan = max(debtinc)
    debtinc_ran = max_loan - min_loan

# debtincoutlier = []
debtinc_outlier = detect_outlier(debtinc)
print("DEBTINC Outlier", 5*' ', debtinc_outlier)
outliers.clear()

# _-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
''' All of these are print Statments are used inorder to print the data calculated
    for example : MEAN, MEDAIN, Standarad Deviation, VARINCE, RAGE, SKEW '''

print("\n")
print("øøøøøøøøøøøøøøøøøøøøøøøøøøøøøøø  •••••••••••••••••••••••••••••••••••••••••  øøøøøøøøøøøøøøøøøøøøøøøøøøøøøøøø")
print("\n")
print("         ","LOAN",'    ',"MORTDUE",'    ',"VALUE",'     ',"YOJ",'   ',"DEROG",'  ' +\
"DELINQ",'  ',"CLAGE", '   ', "NINQ", '     ', "CLNO", '     ',"DEBTINC" )

print("MEAN     ",round(loan_mean,2), '  ',round(mort_mean,2),'  ', +\
round(value_mean,2),'  ',round(yol_mean,2),'  ',round(derog_mean,2),'  ',round(delinq_mean,2),'   ',+\
round(clage_mean,2),'   ',round(ninq_mean,2),'     ', +\
round(clno_mean,2),'     ',round(debtinc_mean,2))

# Medain Statment
print("MEDIAN   ",round(loan_MID,2),'   ' ,round(mort_MID,2),'   ',+\
round(value_MID,2),'   ',round(yol_MID,2),'   ',round(derog_MID,2),'   ',round(delinq_MID,2),'    ', +\
round(clage_MID,2),'   ',round(ninq_MID,2),'      ', +\
round(clno_MID,2),'      ',round(debtinc_MID,2))

# Standarad Deviation
print("Stan.Dev ",round(loan_SD,2),'  ',round(mort_SD,2),'  ',+\
round(value_SD,2), '   ',round(yol_SD,2),'  ',round(derog_SD,2),'  ',round(delinq_SD,2), +\
2*'  ',round(clage_SD,2),'   ',round(ninq_SD,2),2*'   ', +\
round(clno_SD,2),2*'   ' ,round(debtinc_SD,2))

# Range
print("Range    ",round(loan_ran,2),'   ', round(mort_ran,2),'  ', round(value_ran,2),'  ', +\
round(yol_ran,2),'  ',round(derog_ran,2),'   ' ,round(delinq_ran,2),'    ',round(clage_ran,2),'   ',+\
round(ninq_ran,2),'     ', round(clno_ran,2),'      ',round(debtinc_ran,2))

# Skew
print("Skewness", round(loan_skew,2),'     ', round(mort_skew,2),'      ', round(value_skew,2),'      ', +\
round(yol_skew,2),'  ',round(derog_skew,2),'  ' ,round(delinq_skew,2),'   ',round(clage_skew,2),'     ' , round(ninq_skew,2),'     ',+\
round(clno_skew,2),'      ',round(debtinc_skew,2))

# print("\n")
# print(" øøøøøøøøøøøøøøøøøøøøøøøøøøøøø  •••••••••••••••••••••••••••••••••••••••  øøøøøøøøøøøøøøøøøøøøøøøøøøøøøø ")
# print("\n")

print("\n")
print("  ∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚ VARIANCE ˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚ ")
print("\n")
# Varicance
print(18*' ',"Variance ")
print(" Loan",10*' ',round(loan_var,2),"\n","MORTDUE",5*' ',round(mort_var,2),"\n","VALUE",6*' ',+\
round(value_var,2),"\n","YOJ",16*' ',round(yol_var,2),"\n","DEROG",15*' ',round(derog_var,2),"\n","DELINQ",14*' ',+\
round(delinq_var,2),"\n","CLAGE",12*' ', round(clage_var,2),"\n","NINQ",16*' ',round(ninq_var,2),"\n","CLNO",15*' ', +\
round(clno_var,2),"\n","DEBTINC",12*' ',round(debtinc_var,2))
print("\n")
print(" ˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚∆˚  ")
print("\n")
print(10*' ',"  Low Quantile",10*' ', "High Quantile")
# Quantile
print("Loan", 10*' ',loan_min_qurt,16*' ' ,loan_max_qurt)
print("MORTDUE", 6*' ',mort_min_qurt,15*' ' ,mort_max_qurt)
print("VALUE", 8*' ',value_min_qurt,15*' ' ,value_max_qurt)
print("YOJ", 14*' ',yol_min_qurt,18*' ' ,yol_max_qurt)
print("DEROG", 12*' ',derog_min_qurt,19*' ' ,derog_max_qurt)
print("DELINQ", 11*' ',delinq_min_qurt,19*' ' ,delinq_max_qurt)
print("CLAGE", 3*' ',clage_min_qurt,10*' ' ,clage_max_qurt)
print("NINQ", 13*' ',ninq_min_qurt,19*' ' ,ninq_max_qurt)
print("CLNO", 12*' ',clno_min_qurt,18*' ' ,clno_max_qurt)
print("DEBTINC", 1*' ',debtinc_min_qurt,11*' ' ,debtinc_max_qurt)
print("\n")

print(" ˆ¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬ˆ CORRALATION ˆ¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚¬˚ˆ")
corraltion_1 = num.corrcoef(mortdue ,value)
print(" Correlation:", "Mortdue & Value \n",corraltion_1 )
print(" Correlation:", "Mortdue & Value \n", corraltion_1[0][1] )

# Scatter Plot for Mortdue and value
# plt.scatter(mortdue, value)
# plt.xlabel('Mortdue')
# plt.ylabel('value')
# plt.title('Mortgage & Value')
# plt.show()

# # the histogram of the Loan
# plt.hist(loan, facecolor='g')
# plt.xlabel('Loan')
# plt.ylabel('Frequency')
# plt.title('Histogram of Loan')
# plt.show()

# Box Plot for
# plt.boxplot(value)
# plt.xlabel('Value')
# plt.ylabel('value')
# plt.title('Mortgag & Value')
# plt.show()

# Bar Plot for
# width = 1/2.109288
# plt.bar(yol, clno, align = 'center', alpha =.35)
# plt.xlabel('Value')
# plt.show()
# plt.xlabel('Mortdue')
# plt.ylabel('value')
# plt.title('Mortgag & Value')
# plt.show()
