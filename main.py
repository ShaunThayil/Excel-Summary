import pandas as pd
import openpyxl as opnxl
import functions as fun

def main():     #Starting Point Of Excecution
 print ("START!!") 
 #filename=raw_input("Enter The FileName:")
 #sheetname=raw_input("Enter Sheet Name/Index:")
 filename="AWS4 DATA 2012-14.xlsx"
 sheetname="Ist process"
 ex_sheet=pd.read_excel(filename,sheet_name=sheetname)
 ex_sheet.set_index("DATE",drop=False,inplace=True)  #To Set The Index To Date
 column_names=fun.col_names(ex_sheet)        
 index_list=fun.unique_index(ex_sheet)     #To Create A List Of Indices(rows) and Only Permit Unique Values
 wb2=opnxl.Workbook()                      #Opening the Workbook To where Data Is to be written
 savename=filename+"_DailyMean.xlsx"       #File Name To Be Saved
 fun.create_sheets(wb2,column_names,savename) #Create Sheets Corresponding To Each Column
     

 for i in column_names:
     new_sheet=wb2.get_sheet_by_name(i)
     fun.daily_means_max_min(wb2,ex_sheet,index_list,i,new_sheet,savename)
 print ("END!!!")    


if __name__=='__main__':
    main()