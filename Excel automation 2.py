import pandas as pd
import os
from openpyxl import load_workbook
import xlsxwriter
from shutil import copyfile


#C:\Users\user\Desktop\Incercare\office.xlsx - our filepath example

file = input("Please introduce the file path: ")
# Splitting the path into root and extension

workbook = load_workbook(file)
list1 =['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1', 'K1', 'L1', 'M1', 'N1', 'O1', 'P1', 'Q1', 'R1', 'S1', 'T1', 'U1', 'V1', 'W1', 'X1', 'Y1', 'Z1']
x = len(list1)
st=[]
sheet = workbook.active

for i in range(x):
    st.append(sheet[list1[i]].value)

res = []
for val in st:
    if val != None:
        res.append(val)

print(f'This are the columns from the current excel: {res}')

extension = os.path.splitext(file)[1] # extension
filename = os.path.splitext(file)[0]  # root

pth = os.path.dirname(file) # Returning the directory name of the file

newfile = os.path.join(pth, filename + '_2' + extension) # Creating the new file with index 2

df = pd.read_excel(file) # Reading the file
colpick = input("Select your column: ")


cols = list(set(df[colpick].values)) # Creating a list with the columns name
print(f'This is the list with the elements from the column: {cols}')


def sendtofile(cols): # Function that create the files in our directory
    for i in cols: # Iterating inside of the list for every cols name
        df[df[colpick] == i].to_excel("{}/{}.xlsx".format(pth, i), sheet_name=i, index=False)  # Creating the file and name it after the column name. If the column name is 'Normal' the excel file will be Normal.xlsx
    print("\nCompleted!")
    print("Thanks for using this program.")
    return

def sendtosheet(cols): # Function that create the sheets in our excel file
    copyfile(file, newfile) # Create a new copy of the file to not overwrite the data
    for j in cols:
        writer = pd.ExcelWriter(newfile, engine='openpyxl')
        # Creating separte sheets for every column name
        for myname in cols:
            mydf = df.loc[df[colpick] == myname]
            mydf.to_excel(writer, sheet_name = myname, index = False)
        writer.save()
    print('\nCompleted!')
    print('Thanks for using this program.')
    return


# Integrating our functions
print('Do you really want to do that?')
while True:
    x=input('Ready to procced? (Y/N): ').lower()
    if x == 'y':
        while True:
            s = input('Split into different sheets or file? (S/F): ').lower()
            if s == 'f':
                sendtofile(cols)
                break
            elif s == 's':
                sendtosheet(cols)
                break
            else:
                continue
        break
    elif x == 'n':
        print('\nThanks for using this program.')
        break



