# -*- coding: utf-8 -*-
"""
Created on Tue Nov 19 10:51:30 2019
@author: Michael Frangos - MihalyImportant@yahoo.com
"""
import matplotlib.pyplot as plt
import pandas as pd
import os
import numpy as np

Main_Directory = "C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry"
Data_Directory_PSI = "C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Flyer Data/PSI/Cleaned" 
Data_Directory_PVI = "C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Flyer Data/PVI/Cleaned" 
Data_Directory_PVCI = "C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Flyer Data/PVCI/Cleaned" 

#Load Flyer Dates
os.chdir("C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Flyer Data")
PSI_2018_Flyer_Dates = pd.read_excel("PSI 2018 Flyer Date Ranges.xlsx")
PVI_2018_Flyer_Dates = pd.read_excel("PVI 2018 Flyer Date Ranges.xlsx")
PVCI_2018_Flyer_Dates = pd.read_excel("PVCI 2018 Flyer Date Ranges.xlsx")
PSI_2019_Flyer_Dates = pd.read_excel("PSI 2019 Flyer Date Ranges.xlsx")
PVI_2019_Flyer_Dates = pd.read_excel("PVI 2019 Flyer Date Ranges.xlsx")
PVCI_2019_Flyer_Dates = pd.read_excel("PVCI 2019 Flyer Date Ranges.xlsx")

def drop_useless_columns(dirty_data):
  useless_columns = ["Division", "Department","Class","Subclass","Flyer Margin Weekly","Pre Flyer Weekly","Flyer Margin Dollars","Pre Flyer Margin Dollars","Flyer Margin Dollars","Post Dollars","Flyer Dollars","Pre Dollars","Itemized Coop"]    #"Regular Margin Dollars" 
  output = dirty_data
  for useless_column in useless_columns:
    try:    
      output = output.drop(f"{useless_column}",axis = 1)
    except:
      pass
  return output


def clean_page_column(dirty_column):
    cleaned_column = []
    for value in dirty_column:
      try:
        if value.lower().strip().replace(" ","") == "front":
          value = 1
          cleaned_column.append(value)
        elif value == 1:
          value = 1
          cleaned_column.append(value)
        elif value.lower().strip() == "back":
          value = 2
          cleaned_column.append(value)
        elif value.replace("-"," ").lower().strip() == "page 2":
          value = 2
          cleaned_column.append(value)
        elif value.replace("-"," ").lower().strip() == "page 3":
          value = 3
          cleaned_column.append(value)
        elif value.replace("-"," ").lower().strip() == "page 4":
          value = 4
          cleaned_column.append(value)
        elif value.replace("-"," ").lower().strip() == "page 1":
          value = 1
          cleaned_column.append(value)
        elif value.replace("-"," ").lower().strip() == "page1":
          value = 1
          cleaned_column.append(value)
        elif value.replace(" - Select Stores","") == "INSTORE":
          value = "In Store Only"
          cleaned_column.append(value)
        elif "small" in value.lower():
          value = "Small Pets"
          cleaned_column.append(value)
        elif "aquatic" in value.lower():
          value = "Aquatics"
          cleaned_column.append(value)
        elif "store" in value.lower():
          value = "In Store Only"
          cleaned_column.append(value)
        elif type(value) == str:
          try:
            value = int(value)
            cleaned_column.append(value)
          except:
            print(value, " cannot be converted to Int")
            cleaned_column.append(value)
        else:
          print("unknown", value)
          cleaned_column.append(value)
      except:
        if (value != 1) and (value != 2) and (value != 3) and (value != 4):
          print("Exception: ", value)
          cleaned_column.append(value)
        else:
          cleaned_column.append(value) 
    
    return cleaned_column  




#Compile PSI Data
os.chdir(Data_Directory_PSI)
aggregate= pd.DataFrame()
PSI_column_list = [] #to verify integrity across files
for file in os.listdir(Data_Directory_PSI):
  #Identifies if the folder is truly an excel file
  if file[-5:] == ".xlsx":
    print("Current File: ", f"{Data_Directory_PSI}/{file}")
    PSI_File_Fiscal_Year = int(file[-12:][:4])
    PSI_File_Fiscal_Month = int(file[-7:][:2])
    temp_excel_data = pd.read_excel(f"{Data_Directory_PSI}/{file}")
    temp_excel_data = drop_useless_columns(temp_excel_data)
    temp_excel_data["Fiscal Period"] = PSI_File_Fiscal_Month
    temp_excel_data["Fiscal Year"] = PSI_File_Fiscal_Year
    PSI_column_list.append(list(temp_excel_data.columns) + [f"{Data_Directory_PSI}/{file}"[-12:]]) #To check column integrity later
    #temp_excel_data.columns = ["District","Store","Sales"]
    
    ### Debug columns
    #print("Number of columns = ", len(temp_excel_data.columns))
    
    #Clean division column with the new columns
    temp_excel_data["Division"] = temp_excel_data["PRB Division"]
    temp_excel_data["Class"] = temp_excel_data["PRB Class"]
    temp_excel_data["Sub Class"] = temp_excel_data["PRB Sub Class"]
    temp_excel_data["Department"] = temp_excel_data["PRB Department"]
    
    
    #Clean up column names through replacement to allow them to merge
    temp_excel_data.columns = [x.replace("Retail","Reg. Price") for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace("ClassName","Class Name") for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace("Flyer Page","Page").replace('Page\n','Page').replace("Page", "Flyer Page") for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace("Flyer Page","Page") for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace("Page","Flyer Page") for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace("Prior 13 Flyer Margin Dollars","Prior 13 Norm Flyer Margin Dollars") for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace("Regular Margin $","Regular Margin $ Per Unit") for x in temp_excel_data.columns] #NOTE THIS IS BEING CALCULATED BY MARK
    temp_excel_data.columns = [x.replace("Flyer Margin $",'Flyer Margin $ Per Unit') for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace("Attribute 1","Consumable Type") for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace('Flyer Units',"Flyer Quantity") for x in temp_excel_data] #Replace column names
    temp_excel_data.columns = [x.replace("Prior 13 Units","Pre Quantity") for x in temp_excel_data] #Replace column names
    temp_excel_data.columns = [x.replace("Follow 6 Units","Post 6 Weeks Quantity") for x in temp_excel_data] #Replace column names
    temp_excel_data.columns = [x.replace('Flyer $',"Garbage column") for x in temp_excel_data.columns] #Replace column names
    temp_excel_data.columns = [x.replace('Item Description',"Desc") for x in temp_excel_data.columns] #Replace column names
    temp_excel_data.columns = [x.replace('Desc',"Item Description") for x in temp_excel_data.columns] #Replace column names

    

    temp_excel_data["Flyer Page"] = clean_page_column(temp_excel_data["Flyer Page"])

#    Clean_Sku_Column = []
#    for value in temp_excel_data["SKU #"]:
#        try:
#          Clean_Sku_Column.append(int(value))
#        except:
#          Clean_Sku_Column.append(99999999999)
          
    if 'Start Date' in temp_excel_data.columns:
      number_of_PSI_campaign_days = ((temp_excel_data["End Date"] - temp_excel_data["Start Date"]).dt.total_seconds() / (24* 60 * 60)).astype(int)+1
      print("Campaign Days = ", number_of_PSI_campaign_days[0])
    elif PSI_File_Fiscal_Year == 2018: #Pull dates from Mark's provided calendars
      number_of_PSI_campaign_days = int(PSI_2018_Flyer_Dates.loc[PSI_2018_Flyer_Dates["Fiscal Period"] == PSI_File_Fiscal_Month]["# of days"])
      print("USING CALENDAR - Campaign Days = ", number_of_PSI_campaign_days, ".")    
    elif PSI_File_Fiscal_Year == 2019: #Pull dates from Mark's provided calendars
          number_of_PSI_campaign_days = int(PSI_2019_Flyer_Dates.loc[PSI_2019_Flyer_Dates["Fiscal Period"] == PSI_File_Fiscal_Month]["# of days"])
          print("USING CALENDAR - Campaign Days = ", number_of_PSI_campaign_days, ".")                                 
    else: #Default for 2019
      number_of_PSI_campaign_days = 27
      print("Campaign Days = ", number_of_PSI_campaign_days)
    
    #ReWrite of old column
    temp_excel_data['Flyer Margin Per Unit'] = temp_excel_data["Sale Price"] - temp_excel_data["Current Cost"]
    temp_excel_data['Vendor BB'] = temp_excel_data['Vendor BB'].fillna(0) #This was the missing line that caused a lot of issues for us due to null computed cells. 1/15/2020
    #temp_excel_data.columns = [x.replace("Prior 13 Norm Flyer Margin Dollars","Prior 13 Margin $ Per Week") for x in temp_excel_data.columns]
    #Formulate new columns for analysis
    temp_excel_data["Regular_units_per_day"] = temp_excel_data["Pre Quantity"] / 91
    temp_excel_data["Flyer Quantity Per Day"] = temp_excel_data["Flyer Quantity"]/number_of_PSI_campaign_days
    temp_excel_data["Post Quantity Per Day"] = temp_excel_data["Post 6 Weeks Quantity"]/number_of_PSI_campaign_days
    temp_excel_data["Incremental_unit_lift_per_day"] = temp_excel_data["Flyer Quantity Per Day"] - temp_excel_data["Regular_units_per_day"]
    temp_excel_data["Post Incremental_unit_lift_per_day"] = temp_excel_data["Post Quantity Per Day"] - temp_excel_data["Regular_units_per_day"]
    temp_excel_data["% unit lift"] = temp_excel_data["Flyer Quantity Per Day"]/temp_excel_data["Regular_units_per_day"]-1
    temp_excel_data["Incremental_Margin_Per_Day"] = (temp_excel_data["Flyer Quantity Per Day"]*temp_excel_data['Flyer Margin Per Unit']) - (temp_excel_data["Regular Margin $ Per Unit"] * temp_excel_data["Regular_units_per_day"])
    temp_excel_data["Regular Margin per Day"] = temp_excel_data["Pre Quantity"] * temp_excel_data["Regular Margin $ Per Unit"]/ 91 #91 days in 13 weeks
    temp_excel_data["Flyer Margin Per Day"] = temp_excel_data["Flyer Quantity"] * temp_excel_data['Flyer Margin Per Unit'] /number_of_PSI_campaign_days
    temp_excel_data["Incremental Margin Per Week"] = temp_excel_data["Incremental_Margin_Per_Day"]*7
    temp_excel_data["Total Incremental Margin"] = temp_excel_data["Incremental_Margin_Per_Day"]*number_of_PSI_campaign_days
    temp_excel_data["Following 6 Weeks Units Per day"] = temp_excel_data["Post 6 Weeks Quantity"]/42
    temp_excel_data["Total Deal Amount"] = temp_excel_data["Flyer Quantity"] * pd.to_numeric(temp_excel_data["Deal $"].replace("I",""))
    temp_excel_data["Total Campaign Sales (Net of Deal)"] = temp_excel_data['Sale Price'] * temp_excel_data['Flyer Quantity']
    temp_excel_data["Total Vendor Reimbursement"] =  temp_excel_data['Vendor BB'] * temp_excel_data['Flyer Quantity']
    temp_excel_data["Incremental Sales per Day"] = (temp_excel_data["Flyer Quantity Per Day"] - temp_excel_data["Regular_units_per_day"]) * temp_excel_data['Sale Price']
    temp_excel_data["Total Incremental Sales"] = temp_excel_data["Incremental Sales per Day"] * number_of_PSI_campaign_days
    temp_excel_data["Total COGS During Campaign"] = temp_excel_data["Current Cost"] * temp_excel_data['Flyer Quantity']
    temp_excel_data["Margin Per Unit Lift"] = temp_excel_data['Flyer Margin Per Unit'] - temp_excel_data["Regular Margin $ Per Unit"]
    temp_excel_data["Post Margin $ Per Day"] = temp_excel_data["Regular Margin $ Per Unit"] * temp_excel_data["Post 6 Weeks Quantity"] / (6*7)
    temp_excel_data["Incremental Units During Flight"] = temp_excel_data["Incremental_unit_lift_per_day"] *number_of_PSI_campaign_days
    temp_excel_data["Incremental Units After Flight"] = temp_excel_data["Incremental_unit_lift_per_day"] *42
    temp_excel_data["Incremental Margin After Reimbursement"] =  temp_excel_data["Total Incremental Margin"] + (temp_excel_data["Vendor BB"] * temp_excel_data["Flyer Quantity"])
    temp_excel_data["Campaign Days"] = number_of_PSI_campaign_days
    temp_excel_data["Estimated_Regular_Margin_During_Campaign"] = temp_excel_data["Regular Margin $ Per Unit"] * temp_excel_data["Regular_units_per_day"] * temp_excel_data["Campaign Days"]
    temp_excel_data["Flyer Margin Dollars"] = temp_excel_data["Total Campaign Sales (Net of Deal)"] - temp_excel_data["Total COGS During Campaign"] + temp_excel_data["Total Vendor Reimbursement"] #PVI & PVCI only
    temp_excel_data["Flyer Margin Dollars Before Reimbursement"] = temp_excel_data["Total Campaign Sales (Net of Deal)"] - temp_excel_data["Total COGS During Campaign"]
    temp_excel_data["Reimbursement for Incremental Units"] = temp_excel_data["Incremental Units During Flight"] * temp_excel_data['Vendor BB']
    #temp_excel_data["Neg. Inc. Margin Reas.on"] = 
    
        #Put all the pivots here.
    values = ["Placement Fee",'Total Deal Amount', "Regular_units_per_day","Flyer Quantity Per Day",'Post Quantity Per Day',"Following 6 Weeks Units Per day","Total Incremental Sales","Total Incremental Margin",
              "Total Campaign Sales (Net of Deal)","Total Vendor Reimbursement","Total COGS During Campaign",'Flyer Margin Dollars',"Incremental_unit_lift_per_day","Incremental Units During Flight",
              "Incremental Units After Flight","Post Incremental_unit_lift_per_day","Incremental Margin After Reimbursement"]
    temp_ExecSmmryMetrics_by_Page_Number = temp_excel_data.pivot_table(index = ['Flyer Page'], columns = [], values = values, aggfunc = "sum").reset_index()
    temp_Summary_by_Vendor = temp_excel_data.pivot_table(index = ["Vendor Name"], columns = [], values = values, aggfunc = "sum").reset_index()
    #Handle for missing columns
    if 'Class Name' in temp_excel_data.columns:
      temp_Summary_by_Class = temp_excel_data.pivot_table(index = ['Class Name'], columns = [], values = values, aggfunc = "sum").reset_index()
    else:
      print(f"Skipping class summary due to missing columns for {file}")
      temp_excel_data["Class Name"] = "Missing Data"
    
    
    if 'Division' in temp_excel_data.columns:
      temp_excel_data["Division"] = temp_excel_data["Division"].fillna("Missing Data")
      temp_Summary_by_Division = temp_excel_data.pivot_table(index = ['Division'], columns = [], values = values, aggfunc = "sum").reset_index()
    else:
      print(f"Skipping division summary due to missing columns for {file}")
      temp_excel_data["Division"] = "Missing Data"
      temp_Summary_by_Division = temp_excel_data.pivot_table(index = ['Division'], columns = [], values = values, aggfunc = "sum").reset_index()
    #temp_excel_data.pivot_table(index = ["Fiscal Year","Fiscal Period"], columns = [], values = values, aggfunc = "sum").reset_index() #UNUSED
    
    
    #Iteratively Export each Pivot
    Fiscal_Year = int(file[-12:][:4])
    Fiscal_Period = int(file[-7:][:2])
    Time = f"{Fiscal_Year}-{Fiscal_Period}"
    export_list = [temp_ExecSmmryMetrics_by_Page_Number,temp_Summary_by_Vendor,]
    export_name_list = [f"PSI_{Time}_Executive Summary",f"PSI_{Time}_Summary_by_Vendor",]
    #handle for missing columns
    if 'Division' in temp_excel_data.columns:
      export_list.append(temp_Summary_by_Division)
      export_name_list.append(f"PSI_{Time}_Summary_by_Division")
    #handle for missing columns
    if 'Class Name' in temp_excel_data.columns:
      export_list.append(temp_Summary_by_Class)
      export_name_list.append(f"PSI_{Time}_Summary_by_Class")
      
    for pivot_table,file_name in zip(export_list,export_name_list):
      try:
        #Export with total rows
        total_row = pivot_table.sum()
        total_row.iloc[0] = "TOTAL"

        pivot_table.append(total_row,ignore_index = True).to_excel(f"C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Program Output/PSI/{file_name}.xlsx", index = False)
      except Exception as e:
        print(f"Could not export {file_name}, reason: ",e)
    
    #Iteratively append each file
    aggregate = aggregate.append(temp_excel_data,sort=True)

##To debug certain months
#    if file == 'PSI 2019 03.xlsx':
#      break
    
aggregate2 = aggregate#.dropna(axis=1,how = 'all',thresh = 0.4*len(aggregate)) #Drop columns if we don't have at least 50% of the data




PSI_aggregate = aggregate2



#PSI_aggregate = aggregate.dropna(axis=1,how = 'all',thresh = 0.5*len(aggregate)) #Drop columns if we don't have at least 50% of the data

#Put all the pivots here.
#See values list above...
PSI_ExecSmmryMetrics_by_Page_Number = PSI_aggregate.pivot_table(index = ["Fiscal Year","Fiscal Period"], columns = [], values = values, aggfunc = "sum").reset_index()
PSI_Summary_by_Vendor = PSI_aggregate.pivot_table(index = ["Vendor Name"], columns = [], values = values, aggfunc = "sum").reset_index()
PSI_Summary_by_Class = PSI_aggregate.pivot_table(index = ['Class Name'], columns = [], values = values, aggfunc = "sum").reset_index() 
PSI_Summary_by_Division = PSI_aggregate.pivot_table(index = ['Division'], columns = [], values = values, aggfunc = "sum",dropna= False).reset_index()
##MultiYear = PSI_aggregate.pivot_table(index = ["Fiscal Year","Fiscal Period"], columns = [], values = values, aggfunc = "sum").reset_index()


##Export Multi-Year Summaries
export_list = [PSI_ExecSmmryMetrics_by_Page_Number,PSI_Summary_by_Vendor,PSI_Summary_by_Division,PSI_Summary_by_Class]
export_name_list = [f"PSI_{Time}_Executive Summary",f"PSI_{Time}_Summary_by_Vendor",f"PSI_{Time}_Summary_by_Division",f"PSI_{Time}_Summary_by_Class"]
for pivot_table,file_name in zip(export_list,export_name_list):
      try:
        #Export with total rows
        total_row = pivot_table.sum()
        if "Executive Summary" in file_name:  
          total_row.iloc[0] = "TOTAL"
          total_row.iloc[1] = ""
        else:
          total_row.iloc[0] = "TOTAL"
        pivot_table.append(total_row,ignore_index = True).to_excel(f"C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Program Output/PSI/Multi-Year Summary/{file_name}.xlsx", index = False)
      except Exception as e:
        print(f"Could not export {file_name}, reason: ",e)




#PLOT EXECUTIVE SUMMARY
plot_data = PSI_ExecSmmryMetrics_by_Page_Number[['Fiscal Year', 'Fiscal Period', 'Total Campaign Sales (Net of Deal)', 
       'Total COGS During Campaign',
       'Flyer Margin Dollars',
       'Total Deal Amount', 
       'Total Vendor Reimbursement',
       'Total Incremental Sales',
       'Total Incremental Margin',
       ]]
plt.plot(plot_data.iloc[:,2:]) #Plot the graphs
plt.legend(plot_data.columns[2:],
           loc = 'lower left', #Location of legend box
           framealpha = 0, #Transparent legend box
           labelspacing = 0) #Space between legend items
#X-Axis custom tick markers
plt.xticks(range(len(plot_data)),[str(x) + "-" + str(z) for x,z in zip(plot_data.iloc[:,0],plot_data.iloc[:,1])])
plt.xticks(rotation=90) #Rotate x-axis 90 degrees
plt.xlabel("Date")
plt.ylabel("USD $")
plt.title("PSI Flyer Campaign")
plt.locator_params(axis='y', nbins=30) #Number of y-ticks
plt.savefig("C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Program Output/Graphic Summaries/PSI.png",dpi = 3000,bbox_inches='tight')


#Verify column name integrity. Everything looks clean so far.  #Look at the columns across files dataframe
PSI_Columns = pd.DataFrame()
i=0
for columns in PSI_column_list:
  try:
    print(i,columns[i])
    PSI_Columns = pd.concat([pd.DataFrame(columns),PSI_Columns],axis = 1,sort=True)
    i=i+1
  except:
    pass
  













#Compile PVI Data
os.chdir(Data_Directory_PVI)
PVI_aggregate = pd.DataFrame()
PVI_column_list = [] #to verify integrity across files
for file in os.listdir(Data_Directory_PVI):
  
  #Identifies if the folder is truly an excel file
  if file[-5:] == ".xlsx":
    print("Current File: ", f"{Data_Directory_PVI}/{file}")
    PVI_File_Fiscal_Year = int(file[-12:][:4])
    PVI_File_Fiscal_Month = int(file[-7:][:2])
    temp_excel_data = pd.read_excel(f"{Data_Directory_PVI}/{file}")
    temp_excel_data = drop_useless_columns(temp_excel_data) 
    #print(temp_excel_data)
    temp_excel_data["Fiscal Period"] = PVI_File_Fiscal_Month
    temp_excel_data["Fiscal Year"] = PVI_File_Fiscal_Year
    #temp_excel_data.columns = ["District","Store","Sales"]
    PVI_column_list.append(list(temp_excel_data.columns) + [f"{Data_Directory_PVI}/{file}"[-12:]])
    
    
    #Clean up old columns with new ones
    temp_excel_data["Division"] = temp_excel_data["PRB Division"]
    temp_excel_data["Class"] = temp_excel_data["PRB Class"]
    temp_excel_data["Sub Class"] = temp_excel_data["PRB Sub Class"]
    temp_excel_data["Department"] = temp_excel_data["PRB Department"]
    
    
    ##DEBUG COLUMNS
    #print("Number of columns = ", len(temp_excel_data.columns))
    #temp_excel_data.drop("Division",axis=1,inplace = True)
    #Clean up column names through replacement to allow them to merge
    temp_excel_data.columns = [x.replace("Non Member\nPromo Price","Sale Price") for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace('Vendor \nBill-Back\nPOS',"Vendor BB") for x in temp_excel_data.columns] #Replace column names
    temp_excel_data.columns = [x.replace('Post Quantity',"Post 6 Weeks Quantity") for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace('Flyer Margin Dollars','Flyer Margin Per Unit') for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace("Regular Margin Dollars","Regular Margin $ Per Unit") for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace('ItemDescription',"Item Description") for x in temp_excel_data.columns] #Replace column names
    temp_excel_data.columns = [x.replace("Comments","Offer Notes") for x in temp_excel_data.columns] #Replace column names
    temp_excel_data.columns = [x.replace('Flyer Page\nPet Valu',"Flyer Page") for x in temp_excel_data.columns] #Replace column names

    
    #Clean page column
    temp_excel_data["Flyer Page"] = clean_page_column(temp_excel_data["Flyer Page"])

       
    if 'Start Date' in temp_excel_data.columns:
      number_of_PVI_campaign_days = ((temp_excel_data["End Date"] - temp_excel_data["Start Date"]).dt.total_seconds() / (24* 60 * 60)).astype(int)+1
      print("Campaign Days = ", number_of_PVI_campaign_days[0])
    elif PVI_File_Fiscal_Year == 2018: #Pull dates from Mark's provided calendars
      number_of_PVI_campaign_days = int(PVI_2018_Flyer_Dates.loc[PVI_2018_Flyer_Dates["Fiscal Period"] == PVI_File_Fiscal_Month]["# of days"])           
      print("USING CALENDAR - Fiscal Period & Campaign Days = ", number_of_PVI_campaign_days)
    elif PVI_File_Fiscal_Year == 2019: #Pull dates from Mark's provided calendars
      try:
        number_of_PVI_campaign_days = int(PVI_2019_Flyer_Dates.loc[PVI_2019_Flyer_Dates["Fiscal Period"] == PVI_File_Fiscal_Month]["# of days"])           
        print("USING CALENDAR - Fiscal Period & Campaign Days = ", number_of_PVI_campaign_days)
      except:
        print("Calender missing Fiscal Period:",PVI_File_Fiscal_Month, "Defaulting to 13 days.")
        number_of_PVI_campaign_days = 13
    else: #Default for 2019
      number_of_PVI_campaign_days = 13
      print("Campaign Days = ", number_of_PVI_campaign_days)
    
    #Cleaning up/overwriting these columns, even thought they are already provided. 
    #Fill missing values
    temp_excel_data['Sale Price'] = pd.DataFrame({"A":temp_excel_data['Sale Price'], "B": temp_excel_data['Member\nPromo Price']}).max(axis=1)
    temp_excel_data['Vendor BB'] = temp_excel_data['Vendor BB'].fillna(0)
    cogs_per_unit = pd.DataFrame({"A":temp_excel_data['DirectUnitCost']/temp_excel_data["PurchUOM"], "B": temp_excel_data['LandedUSD']}).max(axis=1) 
    temp_excel_data['Current Cost'] = cogs_per_unit #Create the column to match PSI
    temp_excel_data['Flyer Margin Per Unit'] = temp_excel_data['Sale Price'] - cogs_per_unit #Overwrites to remove vendor BB from margin
    #Clean up total margin
    temp_excel_data['Flyer Margin Before Reimbursement'] = temp_excel_data['Flyer Margin Per Unit'] * temp_excel_data["Flyer Quantity"]
    #Clean up deal $ column
    temp_excel_data['Non \nMember  Offer $'] = temp_excel_data["Retail_ON"] - temp_excel_data['Sale Price']
    

    
    #Formulate new columns for analysis
    temp_excel_data["Post Margin Weekly"] = temp_excel_data["Post 6 Weeks Quantity"] * temp_excel_data["Regular Margin $ Per Unit"]
    temp_excel_data["Prior 13 Wks Weekly Units"] = temp_excel_data["Pre Quantity"] / 13 #Divide by 13 weeks.
    temp_excel_data["Flyer Weekly Units"] = temp_excel_data["Flyer Quantity"]/number_of_PVI_campaign_days*7
    temp_excel_data["Post Weekly Units"] = temp_excel_data["Post 6 Weeks Quantity"]/6 #divide by 6 weeks
    temp_excel_data["Vendor Funding"] = temp_excel_data['Vendor BB'] * temp_excel_data["Flyer Quantity"] 
    
    
    
    #VERIFY THESE
    temp_excel_data["Regular_units_per_day"] = temp_excel_data["Prior 13 Wks Weekly Units"] / 7
    temp_excel_data["Flyer Quantity Per Day"] = temp_excel_data["Flyer Weekly Units"]/7
    temp_excel_data["Post Quantity Per Day"] = temp_excel_data["Post Weekly Units"]/7
    temp_excel_data["Incremental_unit_lift_per_day"] = temp_excel_data["Flyer Quantity Per Day"] - temp_excel_data["Regular_units_per_day"]
    temp_excel_data["Post Incremental_unit_lift_per_day"] = temp_excel_data["Post Quantity Per Day"] - temp_excel_data["Regular_units_per_day"]
    temp_excel_data["% unit lift"] = temp_excel_data["Flyer Quantity Per Day"]/temp_excel_data["Regular_units_per_day"]-1
    temp_excel_data["Incremental_Margin_Per_Day"] = (temp_excel_data["Flyer Quantity Per Day"]*temp_excel_data['Flyer Margin Per Unit']) - (temp_excel_data["Regular Margin $ Per Unit"] * temp_excel_data["Regular_units_per_day"])
    temp_excel_data["Regular Margin per Day"] = temp_excel_data["Pre Quantity"] * temp_excel_data["Regular Margin $ Per Unit"]/ 91 
    temp_excel_data["Flyer Margin Per Day"] = temp_excel_data["Flyer Quantity"] * temp_excel_data['Flyer Margin Per Unit'] / number_of_PVI_campaign_days
    temp_excel_data["Incremental Margin Per Week"] = temp_excel_data["Incremental_Margin_Per_Day"]*7
    temp_excel_data["Total Incremental Margin"] = temp_excel_data["Incremental_Margin_Per_Day"]*number_of_PVI_campaign_days 
    temp_excel_data["Following 6 Weeks Units Per day"] = temp_excel_data["Post 6 Weeks Quantity"]/42
    temp_excel_data["Total Deal Amount"] = temp_excel_data["Flyer Quantity"] * temp_excel_data['Non \nMember  Offer $']
    temp_excel_data["Total Campaign Sales (Net of Deal)"] = temp_excel_data['Sale Price'] * temp_excel_data['Flyer Quantity']
    temp_excel_data["Total Vendor Reimbursement"] =  temp_excel_data['Vendor BB'] * temp_excel_data['Flyer Quantity']
    temp_excel_data["Incremental Sales per Day"] = temp_excel_data["Incremental_unit_lift_per_day"] * temp_excel_data['Sale Price']
    temp_excel_data["Total Incremental Sales"] = temp_excel_data["Incremental Sales per Day"] * number_of_PVI_campaign_days 
    temp_excel_data["Total COGS During Campaign"] = pd.DataFrame({"A":temp_excel_data['DirectUnitCost']/temp_excel_data["PurchUOM"], "B": temp_excel_data['LandedUSD']}).max(axis=1) * temp_excel_data['Flyer Quantity']
    temp_excel_data["Margin Per Unit Lift"] = temp_excel_data['Flyer Margin Per Unit'] - temp_excel_data["Regular Margin $ Per Unit"]
    temp_excel_data["Post Margin $ Per Day"] = temp_excel_data["Regular Margin $ Per Unit"] * temp_excel_data["Post 6 Weeks Quantity"] / (6*7)    
    temp_excel_data["Incremental Units During Flight"] = temp_excel_data["Incremental_unit_lift_per_day"] * number_of_PVI_campaign_days
    temp_excel_data["Incremental Units After Flight"] = temp_excel_data["Post Incremental_unit_lift_per_day"] * 42
    temp_excel_data["Incremental Margin After Reimbursement"] =  temp_excel_data["Total Incremental Margin"] + (temp_excel_data["Vendor BB"] * temp_excel_data["Flyer Quantity"])
    temp_excel_data["Campaign Days"] = number_of_PVI_campaign_days
    temp_excel_data["Flyer Margin Dollars"] = temp_excel_data["Total Campaign Sales (Net of Deal)"] - temp_excel_data["Total COGS During Campaign"] + temp_excel_data["Total Vendor Reimbursement"] #PVI & PVCI only
    temp_excel_data["Flyer Margin Dollars Before Reimbursement"] = temp_excel_data["Total Campaign Sales (Net of Deal)"] - temp_excel_data["Total COGS During Campaign"]
    temp_excel_data["Estimated_Regular_Margin_During_Campaign"] = temp_excel_data["Regular Margin $ Per Unit"] * temp_excel_data["Regular_units_per_day"] * temp_excel_data["Campaign Days"]
    temp_excel_data["Reimbursement for Incremental Units"] = temp_excel_data["Incremental Units During Flight"] * temp_excel_data['Vendor BB'] 
    
    #Put all the pivots here.
    values = ['Total Deal Amount', "Regular_units_per_day","Flyer Quantity Per Day",'Post Quantity Per Day',"Following 6 Weeks Units Per day","Total Incremental Sales","Total Incremental Margin",
              "Total Campaign Sales (Net of Deal)","Total Vendor Reimbursement","Total COGS During Campaign",'Flyer Margin Before Reimbursement',"Incremental_unit_lift_per_day","Incremental Units During Flight",
              "Incremental Units After Flight","Post Incremental_unit_lift_per_day","Incremental Margin After Reimbursement"]
    temp_ExecSmmryMetrics_by_Page_Number = temp_excel_data.pivot_table(index = ['Flyer Page'], columns = [], values = values, aggfunc = "sum").reset_index()
    temp_Summary_by_Vendor = temp_excel_data.pivot_table(index = ["VendorName"], columns = [], values = values, aggfunc = "sum").reset_index()
    
    if 'Class' in temp_excel_data.columns:
      temp_Summary_by_Class = temp_excel_data.pivot_table(index = ['Class'], columns = [], values = values, aggfunc = "sum").reset_index() ##Not available with current data
    else:
      print(f"Skipping class summary for {file} due to missing columns.")
      temp_excel_data["Class"] = "Missing Data"
    
    if 'Division' in temp_excel_data.columns:
      temp_Summary_by_Division = temp_excel_data.pivot_table(index = ['Division'], columns = [], values = values, aggfunc = "sum").reset_index()
    else:
       print(f"Skipping division summary for {file} due to missing columns.")
       temp_excel_data["Division"] = "Missing Data"
       temp_Summary_by_Division = temp_excel_data.pivot_table(index = ['Division'], columns = [], values = values, aggfunc = "sum").reset_index()
    
    
    #Iteratively Export each Pivot
    Fiscal_Year = int(file[-12:][:4])
    Fiscal_Period = int(file[-7:][:2])
    Time = f"{Fiscal_Year}-{Fiscal_Period}"
    if 'Division' in temp_excel_data.columns:
      export_list = [temp_ExecSmmryMetrics_by_Page_Number,temp_Summary_by_Vendor,temp_Summary_by_Class,temp_Summary_by_Division,temp_Summary_by_Class]
      export_name_list = [f"PVI_{Time}_Executive Summary",f"PVI_{Time}_Summary_by_Vendor",f"PVI_{Time}_Summary_by_Division",f"PVI_{Time}_Summary_by_Class"]
    else:
      export_list = [temp_ExecSmmryMetrics_by_Page_Number,temp_Summary_by_Vendor]
      export_name_list = [f"PVI_{Time}_Executive Summary",f"PVI_{Time}_Summary_by_Vendor"]
      
    for pivot_table,file_name in zip(export_list,export_name_list):
      try:
        #Export with total rows
        total_row = pivot_table.sum()
        if "Executive Summary" in file_name:  
          pass
        else:
          total_row.iloc[0] = "TOTAL"
        pivot_table.append(total_row,ignore_index = True).to_excel(f"C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Program Output/PVI/{file_name}.xlsx", index = False)
      except Exception as e:
        print(f"Could not export {file_name}, reason: ",e)

    temp_excel_data['Total Deal Amount'] = temp_excel_data['Total Deal Amount'].fillna(0)
    #temp_excel_data["Flyer Margin Dollars.1"] = temp_excel_data["Flyer Margin Dollars.1"].fillna(0)
    PVI_aggregate = PVI_aggregate.append(temp_excel_data,sort=True)  
    #print(PVI_aggregate["Flyer Margin Dollars.1"].isna().value_counts()) #Debug NA's
    


####To debug certain months
#    if file == 'PVI 2019 02.xlsx':
#      break


#Verify column name integrity. Everything looks clean so far.  #Look at the columns across files dataframe
PVI_Columns = pd.DataFrame()
i=0
for columns in PVI_column_list:
  try:
    print(i,columns[i])
    PVI_Columns_across_files = pd.concat([pd.DataFrame(columns),PVI_Columns],axis = 1,sort=True)
    i=i+1
  except:
    pass
  





#Put all the pivots here.
#See values list above...
PVI_ExecSmmryMetrics_by_Page_Number = PVI_aggregate.pivot_table(index = ["Fiscal Year","Fiscal Period"], columns = [], values = values, aggfunc = "sum").reset_index()
PVI_Summary_by_Vendor = PVI_aggregate.pivot_table(index = ["VendorName"], columns = [], values = values, aggfunc = "sum").reset_index()
PVI_Summary_by_Class = PVI_aggregate.pivot_table(index = ['Class'], columns = [], values = values, aggfunc = "sum").reset_index() #NOT AVAILABLE WITH CURRENT DATA
PVI_Summary_by_Division = PVI_aggregate.pivot_table(index = ['Division'], columns = [], values = values, aggfunc = "sum").reset_index()
#PVI_aggregate.pivot_table(index = ["Fiscal Year","Fiscal Period"], columns = [], values = values, aggfunc = "sum").reset_index() #UNUSED

Scotts_Request = PVI_aggregate.loc[(PVI_aggregate["Division"] == "Cat")]

#He wants this filtered for cat food
Scotts_Custom_Summary = Scotts_Request.pivot_table(index = ["Fiscal Year","Fiscal Period",'Flyer Page',"BRAND","Brand"],values = values, aggfunc = "sum").reset_index()
Scotts_Custom_Summary2 = Scotts_Request.pivot_table(index = ["Fiscal Year","Fiscal Period","Class"],values = values, aggfunc = "sum").reset_index()


##Export Multi-Year Summaries
export_list = [PVI_ExecSmmryMetrics_by_Page_Number,PVI_Summary_by_Vendor,PVI_Summary_by_Division,PVI_Summary_by_Class]
export_name_list = [f"PVI_{Time}_Executive Summary",f"PVI_{Time}_Summary_by_Vendor",f"PVI_{Time}_Summary_by_Division",f"PVI_{Time}_Summary_by_Class"]
for pivot_table,file_name in zip(export_list,export_name_list):
      try:
        #Export with total rows
        total_row = pivot_table.sum()
        if "Executive Summary" in file_name:  
          total_row.iloc[0] = "TOTAL"
          total_row.iloc[1] = ""
        else:
          total_row.iloc[0] = "TOTAL"
        pivot_table.append(total_row,ignore_index = True).to_excel(f"C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Program Output/PVI/Multi-Year Summary/{file_name}.xlsx", index = False)
      except Exception as e:
        print(f"Could not export {file_name}, reason: ",e)


#PLOT EXECUTIVE SUMMARY
plot_data = PVI_ExecSmmryMetrics_by_Page_Number[['Fiscal Year', 'Fiscal Period', 
       'Total Campaign Sales (Net of Deal)', 
       'Total COGS During Campaign',
       'Flyer Margin Before Reimbursement',
       'Total Incremental Sales',
       'Total Deal Amount',
       'Total Incremental Margin', 
       'Total Vendor Reimbursement']]
plt.plot(plot_data.iloc[:,2:]) #Plot the graphs
plt.legend(plot_data.columns[2:],
           loc = 'lower left', #Location of legend box
           framealpha = 0, #Transparent legend box
           labelspacing = 0) #Space between legend items
#X-Axis custom tick markers
plt.xticks(range(len(plot_data)),[str(x) + "-" + str(z) for x,z in zip(plot_data.iloc[:,0],plot_data.iloc[:,1])])
plt.xticks(rotation=90) #Rotate x-axis 90 degrees
plt.xlabel("Date")
plt.ylabel("USD $")
plt.title("PVI Flyer Campaign")
plt.locator_params(axis='y', nbins=20) #Number of y-ticks
plt.savefig("C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Program Output/Graphic Summaries/PVI.png",dpi = 3000,bbox_inches='tight')




















#Compile PVCI Data
os.chdir(Data_Directory_PVCI)
PVCI_aggregate = pd.DataFrame()
PVCI_column_list = [] #to verify integrity across files
for file in os.listdir(Data_Directory_PVCI):
  
  #Identifies if the folder is truly an excel file
  if file[-5:] == ".xlsx":
    print("Current File: ", f"{Data_Directory_PVCI}/{file}")
    PVCI_File_Fiscal_Year = int(file[-12:][:4])
    PVCI_File_Fiscal_Month = int(file[-7:][:2])
    temp_excel_data = pd.read_excel(f"{Data_Directory_PVCI}/{file}")
    temp_excel_data = drop_useless_columns(temp_excel_data)
    temp_excel_data["Fiscal Period"] = PVCI_File_Fiscal_Month
    temp_excel_data["Fiscal Year"] = PVCI_File_Fiscal_Year
    PVCI_column_list.append(list(temp_excel_data.columns) + [f"{Data_Directory_PVCI}/{file}"[-12:]])
    
    
    #Clean up old columns with new ones
    temp_excel_data["Division"] = temp_excel_data["PRB Division"]
    temp_excel_data["Class"] = temp_excel_data["PRB Class"]
    temp_excel_data["Sub Class"] = temp_excel_data["PRB Sub Class"]
    temp_excel_data["Department"] = temp_excel_data["PRB Department"]
    
    ##DEBUG COLUMNS
    #print("Number of columns = ", len(temp_excel_data.columns))
    #temp_excel_data.drop("Division",axis=1,inplace = True)
    #Clean up column names through replacement to allow them to merge
    temp_excel_data.columns = [x.replace("Non Member\nPromo Price","Sale Price") for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace('Vendor \nBill-Back\nPOS',"Vendor BB") for x in temp_excel_data.columns] #Replace column names
    temp_excel_data.columns = [x.replace('Post Quantity',"Post 6 Weeks Quantity") for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace('Flyer Margin Dollars','Flyer Margin Per Unit') for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace("Regular Margin Dollars","Regular Margin $ Per Unit") for x in temp_excel_data.columns]
    temp_excel_data.columns = [x.replace('ItemDescription',"Item Description") for x in temp_excel_data.columns] #Replace column names
    temp_excel_data.columns = [x.replace(' DirectUnitCost ','DirectUnitCost') for x in temp_excel_data.columns] #Replace column names
    temp_excel_data.columns = [x.replace(' LandedUSD ','LandedUSD') for x in temp_excel_data.columns] #Replace column names
    temp_excel_data.columns = [x.replace(" Retail_ON ","Retail_ON") for x in temp_excel_data.columns] #Replace column names
    temp_excel_data.columns = [x.replace("Comments","Offer Notes") for x in temp_excel_data.columns] #Replace column names
    #temp_excel_data.columns = [x.replace('SubClass',"Sub Class").replace('Subclass',"Sub Class") for x in temp_excel_data.columns] #Replace column names
    #temp_excel_data.columns = [x.replace('Division.1',"Division") for x in temp_excel_data.columns] #Replace column names
    temp_excel_data.columns = [x.replace('Flyer Page\nPet Valu',"Flyer Page") for x in temp_excel_data.columns] #Replace column names
    #temp_excel_data.columns = [x.replace('Franchise Flyer Qty',"Franchise Flyer QTY") for x in temp_excel_data.columns] #Replace column names
    
    
#    temp_excel_data.columns = [x.replace("ClassName","Class Name") for x in temp_excel_data.columns]
#    temp_excel_data.columns = [x.replace("Flyer Page","Page").replace('Page\n','Page') for x in temp_excel_data.columns]
#    temp_excel_data.columns = [x.replace("Prior 13 Flyer Margin Dollars","Prior 13 Norm Flyer Margin Dollars") for x in temp_excel_data.columns]
#    temp_excel_data.columns = [x.replace("Sub-Class","Sub Class").replace("SubClass","Sub Class") for x in temp_excel_data.columns]
#    temp_excel_data.columns = [x.replace("Regular Margin $","Regular Margin $ Per Unit") for x in temp_excel_data.columns]
#    temp_excel_data.columns = [x.replace("Flyer Margin $",'Flyer Margin $ Per Unit') for x in temp_excel_data.columns]
#    temp_excel_data.columns = [x.replace("Attribute 1","Consumable Type") for x in temp_excel_data.columns]
    
    temp_excel_data["Flyer Page"] = clean_page_column(temp_excel_data["Flyer Page"])



       
    if 'Start Date' in temp_excel_data.columns:
      number_of_PVCI_campaign_days = ((temp_excel_data["End Date"] - temp_excel_data["Start Date"]).dt.total_seconds() / (24* 60 * 60)).astype(int)+1
      print("Campaign Days = ", number_of_PVCI_campaign_days[0])
    elif PVCI_File_Fiscal_Year == 2018: #Pull dates from Mark's provided calendars
      number_of_PVCI_campaign_days = int(PVCI_2018_Flyer_Dates.loc[PVCI_2018_Flyer_Dates["Fiscal Period"] == PVCI_File_Fiscal_Month]["# of days"])           
      print("USING CALENDAR - Fiscal Period & Campaign Days = ", number_of_PVCI_campaign_days)
    elif PVCI_File_Fiscal_Year == 2019: #Pull dates from Mark's provided calendars
      number_of_PVCI_campaign_days = int(PVCI_2019_Flyer_Dates.loc[PVCI_2019_Flyer_Dates["Fiscal Period"] == PVCI_File_Fiscal_Month]["# of days"])           
      print("USING CALENDAR - Fiscal Period & Campaign Days = ", number_of_PVCI_campaign_days)                                 
    else: #Default for 2019
      number_of_PVCI_campaign_days = 11
      print("Campaign Days = ", number_of_PVCI_campaign_days)


    #Merges the Nonmember price per unit column with the member price per unit column
    temp_excel_data['Sale Price'] = pd.DataFrame({"A":temp_excel_data['Sale Price'], "B": temp_excel_data['Member\nPromo Price']}).max(axis=1)
    #Fill missing values
    temp_excel_data['Vendor BB'] = temp_excel_data['Vendor BB'].fillna(0)    
    #Cleaning up this column, even thought it is already provided. This is margin per unit
    
    
    #PVCI CURRENT COST SELECTION
    try:
      #cogs_per_unit = pd.DataFrame({"A":temp_excel_data['DirectUnitCost']/temp_excel_data["PurchUOM"], "B": temp_excel_data['LandedCAD']}).max(axis=1)  #FIXED FROM LANDE_USD 1/15/20
      cogs_df = pd.DataFrame({"DUC_PurchUOM":temp_excel_data['DirectUnitCost']/temp_excel_data["PurchUOM"], 
                              "CAD_COSTS"   :temp_excel_data['LandedCAD']})
      
      cogs_df['Selected Cost'] = cogs_df.apply(
                                    lambda row: row['DUC_PurchUOM'] if ( np.isnan(row['CAD_COSTS']) or row['CAD_COSTS'] == 0 ) else row['CAD_COSTS'], #if the canadian dollar column is missing or equal to zero, then use DirectUnitCost/PurchUOM
                                    axis=1)
      
      cogs_per_unit = cogs_df['Selected Cost'] #The logic here is that if the canadian dollar column is missing or equal to zero, then use DirectUnitCost/PurchUOM
      
    except:
      print(f"LandedCAD not present in {PVCI_File_Fiscal_Year}, Period {PVCI_File_Fiscal_Month}, using DirectUnitCost/PurchUOM")
      cogs_per_unit = temp_excel_data['DirectUnitCost']/temp_excel_data["PurchUOM"]
    temp_excel_data['Current Cost'] = cogs_per_unit #Create the column to match PSI
    
    
    temp_excel_data['Flyer Margin Per Unit'] = temp_excel_data['Sale Price'] - cogs_per_unit
    #Clean up total margin
    temp_excel_data['Flyer Margin Before Reimbursement'] = temp_excel_data['Flyer Margin Per Unit'] * temp_excel_data["Flyer Quantity"]
    

    
    #Clean up deal $ column
    temp_excel_data['Non \nMember  Offer $'] = temp_excel_data["Retail_ON"] - temp_excel_data['Sale Price']
    

    
    #Formulate new columns for analysis
    temp_excel_data["Post Margin Weekly"] = temp_excel_data["Post 6 Weeks Quantity"] * temp_excel_data["Regular Margin $ Per Unit"]
    temp_excel_data["Prior 13 Wks Weekly Units"] = temp_excel_data["Pre Quantity"] / 13 #Divide by 13 weeks.
    temp_excel_data["Flyer Weekly Units"] = temp_excel_data["Flyer Quantity"]/number_of_PVCI_campaign_days*7
    temp_excel_data["Post Weekly Units"] = temp_excel_data["Post 6 Weeks Quantity"]/6 #divide by 6 weeks
    temp_excel_data["Vendor Funding"] = temp_excel_data['Vendor BB'] * temp_excel_data["Flyer Quantity"] 
    
    
    
    #Everything is before reimbursement unless named otherwise
    temp_excel_data["Regular_units_per_day"] = temp_excel_data["Prior 13 Wks Weekly Units"] / 7
    temp_excel_data["Flyer Quantity Per Day"] = temp_excel_data["Flyer Weekly Units"]/7
    temp_excel_data["Post Quantity Per Day"] = temp_excel_data["Post Weekly Units"]/7
    temp_excel_data["Incremental_unit_lift_per_day"] = temp_excel_data["Flyer Quantity Per Day"] - temp_excel_data["Regular_units_per_day"]
    temp_excel_data["Post Incremental_unit_lift_per_day"] = temp_excel_data["Post Quantity Per Day"] - temp_excel_data["Regular_units_per_day"]
    temp_excel_data["% unit lift"] = temp_excel_data["Flyer Quantity Per Day"]/temp_excel_data["Regular_units_per_day"]-1
    temp_excel_data["Incremental_Margin_Per_Day"] = (temp_excel_data["Flyer Quantity Per Day"]*temp_excel_data['Flyer Margin Per Unit']) - (temp_excel_data["Regular Margin $ Per Unit"] * temp_excel_data["Regular_units_per_day"])
    temp_excel_data["Regular Margin per Day"] = temp_excel_data["Pre Quantity"] * temp_excel_data["Regular Margin $ Per Unit"]/ 91 
    temp_excel_data["Flyer Margin Per Day"] = temp_excel_data["Flyer Quantity"] * temp_excel_data['Flyer Margin Per Unit'] / number_of_PVCI_campaign_days
    temp_excel_data["Incremental Margin Per Week"] = temp_excel_data["Incremental_Margin_Per_Day"]*7
    temp_excel_data["Total Incremental Margin"] = temp_excel_data["Incremental_Margin_Per_Day"]*number_of_PVCI_campaign_days 
    temp_excel_data["Following 6 Weeks Units Per day"] = temp_excel_data["Post 6 Weeks Quantity"]/42 #42 days in 6 weeks
    temp_excel_data["Total Deal Amount"] = temp_excel_data["Flyer Quantity"] * temp_excel_data['Non \nMember  Offer $']
    temp_excel_data["Total Campaign Sales (Net of Deal)"] = temp_excel_data["Sale Price"] * temp_excel_data['Flyer Quantity']
    temp_excel_data["Total Vendor Reimbursement"] =  temp_excel_data['Vendor BB'] * temp_excel_data['Flyer Quantity']
    temp_excel_data["Incremental Sales per Day"] = temp_excel_data["Incremental_unit_lift_per_day"] * temp_excel_data['Sale Price']
    temp_excel_data["Total Incremental Sales"] = temp_excel_data["Incremental Sales per Day"] * number_of_PVCI_campaign_days 
    temp_excel_data["Total COGS During Campaign"] = pd.DataFrame({"A":temp_excel_data['DirectUnitCost']/temp_excel_data["PurchUOM"], "B": temp_excel_data['LandedUSD']}).max(axis=1) * temp_excel_data['Flyer Quantity']
    temp_excel_data["Margin Per Unit Lift"] = temp_excel_data['Flyer Margin Per Unit'] - temp_excel_data["Regular Margin $ Per Unit"]
    temp_excel_data["Post Margin $ Per Day"] = temp_excel_data["Regular Margin $ Per Unit"] * temp_excel_data["Post 6 Weeks Quantity"] / (6*7)    
    temp_excel_data["Incremental Units During Flight"] = temp_excel_data["Incremental_unit_lift_per_day"] * number_of_PVCI_campaign_days
    temp_excel_data["Incremental Units After Flight"] = temp_excel_data["Post Incremental_unit_lift_per_day"] * 42
    temp_excel_data["Incremental Margin After Reimbursement"] =  temp_excel_data["Total Incremental Margin"] + (temp_excel_data["Vendor BB"] * temp_excel_data["Flyer Quantity"])
    temp_excel_data["Campaign Days"] = number_of_PVCI_campaign_days
    temp_excel_data["Flyer Margin Dollars"] = temp_excel_data["Total Campaign Sales (Net of Deal)"] - temp_excel_data["Total COGS During Campaign"] + temp_excel_data["Total Vendor Reimbursement"] #PVI & PVCI only
    temp_excel_data["Flyer Margin Dollars Before Reimbursement"] = temp_excel_data["Total Campaign Sales (Net of Deal)"] - temp_excel_data["Total COGS During Campaign"]
    temp_excel_data["Estimated_Regular_Margin_During_Campaign"] = temp_excel_data["Regular Margin $ Per Unit"] * temp_excel_data["Regular_units_per_day"] * temp_excel_data["Campaign Days"]
    temp_excel_data["Reimbursement for Incremental Units"] = temp_excel_data["Incremental Units During Flight"] * temp_excel_data['Vendor BB']
    
        #Put all the pivots here.
    values = ['Total Deal Amount', "Regular_units_per_day","Flyer Quantity Per Day",'Post Quantity Per Day',"Following 6 Weeks Units Per day","Total Incremental Sales","Total Incremental Margin",
              "Total Campaign Sales (Net of Deal)","Total Vendor Reimbursement","Total COGS During Campaign",'Flyer Margin Before Reimbursement',"Incremental_unit_lift_per_day","Incremental Units During Flight",
              "Incremental Units After Flight","Post Incremental_unit_lift_per_day","Incremental Margin After Reimbursement",]
    temp_ExecSmmryMetrics_by_Page_Number = temp_excel_data.pivot_table(index = ['Flyer Page'], columns = [], values = values, aggfunc = "sum").reset_index()
    temp_Summary_by_Vendor = temp_excel_data.pivot_table(index = ["VendorName"], columns = [], values = values, aggfunc = "sum").reset_index()
    
    if 'Class' in temp_excel_data.columns:
      temp_Summary_by_Class = temp_excel_data.pivot_table(index = ['Class'], columns = [], values = values, aggfunc = "sum").reset_index() ##Not available with current data
    else:
      print(f"Skipping class summary for {file} due to missing columns.")
      temp_excel_data["Class"] = "Missing Data"
    
    if 'Division' in temp_excel_data.columns:
      temp_Summary_by_Division = temp_excel_data.pivot_table(index = ['Division'], columns = [], values = values, aggfunc = "sum").reset_index()
    else:
       print(f"Skipping division summary for {file} due to missing columns.")
       temp_excel_data["Division"] = "Missing Data"
       temp_Summary_by_Division = temp_excel_data.pivot_table(index = ['Division'], columns = [], values = values, aggfunc = "sum").reset_index()
    #temp_excel_data.pivot_table(index = ["Fiscal Year","Fiscal Period"], columns = [], values = values, aggfunc = "sum").reset_index() #UNUSED
    
    
    #Iteratively Export each Pivot
    Fiscal_Year = int(file[-12:][:4])
    Fiscal_Period = int(file[-7:][:2])
    Time = f"{Fiscal_Year}-{Fiscal_Period}"
    if 'Division' in temp_excel_data.columns:
      export_list = [temp_ExecSmmryMetrics_by_Page_Number,temp_Summary_by_Vendor,temp_Summary_by_Class,temp_Summary_by_Division,temp_Summary_by_Class]
      export_name_list = [f"PVCI_{Time}_Executive Summary",f"PVCI_{Time}_Summary_by_Vendor",f"PVCI_{Time}_Summary_by_Division",f"PVCI_{Time}_Summary_by_Class"]
    else:
      export_list = [temp_ExecSmmryMetrics_by_Page_Number,temp_Summary_by_Vendor]
      export_name_list = [f"PVCI_{Time}_Executive Summary",f"PVCI_{Time}_Summary_by_Vendor"]
      
    for pivot_table,file_name in zip(export_list,export_name_list):
      try:
        #Export with total rows
        total_row = pivot_table.sum()
        if "Executive Summary" in file_name:  
          pass
        else:
          total_row.iloc[0] = "TOTAL"
        pivot_table.append(total_row,ignore_index = True).to_excel(f"C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Program Output/PVCI/{file_name}.xlsx", index = False)
      except Exception as e:
        print(f"Could not export {file_name}, reason: ",e)

    temp_excel_data['Total Deal Amount'] = temp_excel_data['Total Deal Amount'].fillna(0)
    #temp_excel_data["Flyer Margin Dollars.1"] = temp_excel_data["Flyer Margin Dollars.1"].fillna(0)
    PVCI_aggregate = PVCI_aggregate.append(temp_excel_data,sort=True)  
    #print(PVCI_aggregate["Flyer Margin Dollars.1"].isna().value_counts()) #Debug NA's
#    if file == "PVCI 2019 04.xlsx":
#      break


####To debug certain months
#    if file == 'PVCI 2018 01.xlsx':
#      break


#Verify column name integrity. Everything looks clean so far. #Look at the columns across files dataframe
PVCI_Columns_across_files = pd.DataFrame()
i=0
for columns in PVCI_column_list:
  #print(columns)
  try:
    #print(i,columns)
    PVCI_Columns_across_files = pd.concat([pd.DataFrame(columns),PVCI_Columns_across_files],axis = 1,sort=True)
    i=i+1
  except:
    pass

  





#Put all the pivots here.
#See values list above...
PVCI_ExecSmmryMetrics_by_Page_Number = PVCI_aggregate.pivot_table(index = ["Fiscal Year","Fiscal Period"], columns = [], values = values, aggfunc = "sum").reset_index()
PVCI_Summary_by_Vendor = PVCI_aggregate.pivot_table(index = ["VendorName"], columns = [], values = values, aggfunc = "sum").reset_index()
PVCI_Summary_by_Class = PVCI_aggregate.pivot_table(index = ['Class'], columns = [], values = values, aggfunc = "sum").reset_index() #NOT AVAILABLE WITH CURRENT DATA
PVCI_Summary_by_Division = PVCI_aggregate.pivot_table(index = ['Division'], columns = [], values = values, aggfunc = "sum").reset_index()
#PVCI_aggregate.pivot_table(index = ["Fiscal Year","Fiscal Period"], columns = [], values = values, aggfunc = "sum").reset_index() #UNUSED

Scotts_Request = PVCI_aggregate.loc[(PVCI_aggregate["Division"] == "Cat")]

#He wants this filtered for cat food
Scotts_Custom_Summary = Scotts_Request.pivot_table(index = ["Fiscal Year","Fiscal Period",'Flyer Page',"BRAND","Brand"],values = values, aggfunc = "sum").reset_index()
Scotts_Custom_Summary2 = Scotts_Request.pivot_table(index = ["Fiscal Year","Fiscal Period","Class"],values = values, aggfunc = "sum").reset_index()


##Export Multi-Year Summaries
export_list = [PVCI_ExecSmmryMetrics_by_Page_Number,PVCI_Summary_by_Vendor,PVCI_Summary_by_Division,PVCI_Summary_by_Class]
export_name_list = [f"PVCI_{Time}_Executive Summary",f"PVCI_{Time}_Summary_by_Vendor",f"PVCI_{Time}_Summary_by_Division",f"PVCI_{Time}_Summary_by_Class"]
for pivot_table,file_name in zip(export_list,export_name_list):
      try:
        #Export with total rows
        total_row = pivot_table.sum()
        if "Executive Summary" in file_name:  
          total_row.iloc[0] = "TOTAL"
          total_row.iloc[1] = ""
        else:
          total_row.iloc[0] = "TOTAL"
        pivot_table.append(total_row,ignore_index = True).to_excel(f"C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Program Output/PVCI/Multi-Year Summary/{file_name}.xlsx", index = False)
      except Exception as e:
        print(f"Could not export {file_name}, reason: ",e)


#PLOT EXECUTIVE SUMMARY
plot_data = PVCI_ExecSmmryMetrics_by_Page_Number[['Fiscal Year', 'Fiscal Period', 
       'Total Campaign Sales (Net of Deal)', 
       'Total COGS During Campaign',
       'Flyer Margin Before Reimbursement',
       'Total Incremental Sales',
       'Total Deal Amount',
       'Total Incremental Margin', 
       'Total Vendor Reimbursement']]
plt.plot(plot_data.iloc[:,2:]) #Plot the graphs
plt.legend(plot_data.columns[2:],
           loc = 'lower left', #Location of legend box
           framealpha = 0, #Transparent legend box
           labelspacing = 0) #Space between legend items
#X-Axis custom tick markers
plt.xticks(range(len(plot_data)),[str(x) + "-" + str(z) for x,z in zip(plot_data.iloc[:,0],plot_data.iloc[:,1])])
plt.xticks(rotation=90) #Rotate x-axis 90 degrees
plt.xlabel("Date")
plt.ylabel("USD $")
plt.title("PVCI Flyer Campaign")
plt.ticklabel_format(style='plain', axis='y') #Suppress Scientific notation
plt.locator_params(axis='y', nbins=20) #Number of y-ticks
plt.savefig("C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Program Output/Graphic Summaries/PVCI.png",dpi = 3000,bbox_inches='tight')
    
    









############################## PRB AGGREGATION ##########################################################################
############################## PRB AGGREGATION ##########################################################################
############################## PRB AGGREGATION ##########################################################################


#Align columns
PSI_aggregate["Brand"].replace("Private Brand","PRIVATE LABEL", inplace = True)

#Throw away trash columns
try:
  PSI_aggregate = PSI_aggregate.drop("Sub-Class",axis = 1) #Dropping deprecated columns. Sub class is the true column
except:
  pass
try:
  PSI_aggregate = PSI_aggregate.drop("SubClass",axis = 1) #Dropping deprecated columns. Sub class is the true column
except:
  pass

#Align columns
PVI_aggregate.columns = [x.replace("Retail_ON","Reg. Price") for x in PVI_aggregate.columns] #Replace column names
PVCI_aggregate.columns = [x.replace("Retail_ON","Reg. Price") for x in PVCI_aggregate.columns] #Replace column names
PSI_aggregate.columns = [x.replace("SKU #","Item") for x in PSI_aggregate.columns] #Replace column names
                                   
                                   
PSI_aggregate["Banner"] = "PSI" #Add a banner column

#MERGE AGGREGATES INTO PRB AGGREGATE
PRB_aggregate = pd.concat([PVI_aggregate,PSI_aggregate,PVCI_aggregate],sort=True)

#Fix error in the source data. Private label consumables should not have reimbursements.
PRB_aggregate["Vendor BB"].loc[(PRB_aggregate["Division"] == "Consumables") & (PRB_aggregate["Brand"] == "PRIVATE LABEL")] = 0

#Clean up Brand Column
PRB_aggregate["Brand"] = PRB_aggregate["Brand"].replace("BRAND","Brand")
#Condense pages column values to 1,2, other
PRB_aggregate["Flyer Page"].loc[(PRB_aggregate["Flyer Page"] != 1) & (PRB_aggregate["Flyer Page"] != 2)] = "Other"

#Create new column to analyze reimbursement.
Reimbursement_Percentage = PRB_aggregate["Vendor BB"]/(PRB_aggregate["Reg. Price"] - PRB_aggregate["Sale Price"])
PRB_aggregate["Reimbursement Percentage"] = Reimbursement_Percentage

#Clean up Department column values
dirty_names = list(set(PRB_aggregate["Department"]))
for item in dirty_names:
  print(item)
  try:
    PRB_aggregate["Department"].replace(f"{item}",f"{item[0]}".upper() + f"{item[1:]}".lower(), inplace = True) #Cleans up all of the items that looks like this >>> PRB_aggregate["Sub Class"] = PRB_aggregate["Sub Class"].replace("TOYS","Toys")
  except:
    pass

#Clean up values to merge the buckets
PRB_aggregate["Division"].replace("CONSUMABLES","Consumables", inplace = True)
PRB_aggregate["Division"].replace("HARDLINES","Hardlines", inplace = True)
PRB_aggregate["Division"].replace("SPECIALTY","Specialty", inplace = True)
PRB_aggregate["Class"].replace("BASIC","Basic", inplace = True)
PRB_aggregate["Class"].replace("FEEDING","Feeding", inplace = True)
PRB_aggregate["Class"].replace("GENERAL MERCHANDISE","General Merchandise", inplace = True)
PRB_aggregate["Class"].replace("HABITAT","Habitat", inplace = True)
PRB_aggregate["Class"].replace("SCIENTIFIC","Scientific", inplace = True)
PRB_aggregate["Class"].replace("TREATS","Treats", inplace = True)
PRB_aggregate["Class"].replace("WILD BIRD","Wild Bird", inplace = True)
PRB_aggregate["Class"].replace("LITTER","Litter", inplace = True)
PRB_aggregate["Class"].replace("LIVE ANIMALS","Live Animals", inplace = True)
PRB_aggregate["Class"].replace("NATURAL","Natural", inplace = True)
PRB_aggregate["Class"].replace("SOLUTIONS","Solutions", inplace = True)

#Clean up Subclass column
PRB_aggregate["Sub Class"] = PRB_aggregate["Sub Class"].replace("AQUARIUM DECOR","Aquarium Decor").replace('Aquarium DÒcor',"Aquarium Decor").replace('AQUARIUM',"Aquarium").replace('AQUARIUM EQUIPMENT','Aquarium Equipment')

dirty_names = ['ACCESSORIES','ALTERNATIVE', 'APPAREL', 'AQUARIUM EQUIPMENT', 'Accessories', 'Alternative', 'Apparel', 'Aquarium', 'Aquarium Decor', 'Aquarium Equipment', 'BEDDING', 'BEDDING AND LITTER', 'BISCUITS', 'Bedding', 'Bedding/Litter', 'Biscuits', 'CAGES', 'CAGES AND FURNITURE', "CAT FOOD DISCO'D", 'CAT NUTRITIONAL & HEALTH', 'CLUMPING', 'COLLARS AND LEASHES', 'CONTAINMENT', 'CULINARY', 'Cages', 'Cages & Furniture', 'Clumping', 'Collars & Leashes', 'Containment', 'Conventional', 'Culinary', 'DENTAL', 'DOG COATS & JACKETS', 'DOG SWEATERS', 'DOG TEE & POLO SHIRTS', 'Dental', 'ENCLOSURES', 'ENHANCED', 'ENTERTAINMENT', 'EQUIPMENT', 'Enclosures', 'Enhanced', 'Entertainment', 'Equipment', 'FEEDERS', 'FEEDING', 'FLEA - TICK', 'FLEA AND TICK', 'FOOD', 'FURNITURE', 'Feeders', 'Feeding', 'Flea & Tick', 'Food', 'Furniture', 'GIFTABLES', 'GROOMING', 'Giftables', 'Grooming', 'HAY', 'HEALTH - WELLNESS', 'HEALTH AND WELLNESS', 'Hay', 'Health & Wellness', 'LITTER ACCESSORIES', 'Litter Accessories', 'MAINTENANCE', 'Maintenance', 'OCCUPANCY', 'Occupancy', 'Operating Supplies', 'PLANTS', 'POND', 'PRO PLAN ADULT', 'PRO PLAN CAT', 'Pond', 'REPTILE DECOR', 'REWARD', 'REWARD-TRAINING', 'Reptile Decor', 'Reptile DÒcor', 'Reward', 'Reward/Training', 'SPECIES', 'Seasonal', 'Species', 'TOYS', 'TRAINING AND ELECTRONICS', 'TRAVEL', 'TREATS', 'Toys', 'Training & Electronics', 'Travel', 'Treats', 'Undefined', 'WASTE MANAGEMENT', 'WATER CARE', 'WILD BIRD', 'Waste Management', 'Water Care','Wild Bird']
for item in dirty_names:
  print(item)
  PRB_aggregate["Sub Class"].replace(f"{item}",f"{item[0]}".upper() + f"{item[1:]}".lower(), inplace = True) #Cleans up all of the items that looks like this >>> PRB_aggregate["Sub Class"] = PRB_aggregate["Sub Class"].replace("TOYS","Toys")
#More cleaning of Subclass column  
PRB_aggregate["Sub Class"] = PRB_aggregate["Sub Class"].replace("Flea - tick","Flea & tick").replace("Flea and tick","Flea & tick")
PRB_aggregate["Sub Class"] = PRB_aggregate["Sub Class"].replace("Bedding and litter","Bedding/litter")
PRB_aggregate["Sub Class"] = PRB_aggregate["Sub Class"].replace("Cages and furniture","Cages & furniture")
PRB_aggregate["Sub Class"] = PRB_aggregate["Sub Class"].replace("Collars and leashes","Collars & leashes")
PRB_aggregate["Sub Class"] = PRB_aggregate["Sub Class"].replace("Health - wellness","Health & wellness").replace("Health and wellness","Health & wellness")
PRB_aggregate["Sub Class"] = PRB_aggregate["Sub Class"].replace("Reptile dòcor","Reptile decor")
PRB_aggregate["Sub Class"] = PRB_aggregate["Sub Class"].replace("Training and electronics","Training & electronics")
PRB_aggregate["Sub Class"] = PRB_aggregate["Sub Class"].replace("Reward-training","Reward/training")
#Per Mark Furry, remove extraneous 0 value from private label flag. Safe to assume it is BRAND.
PRB_aggregate["Brand"].replace(0,"Brand",inplace = True)
#Per Mark Furry, we can assume empty values are part of BRAND.
PRB_aggregate["Brand"].fillna("Brand",inplace = True)






def clean_up_banners(PRB_aggregate):
  #Clean up banner column
  PRB_aggregate["Banner"] = PRB_aggregate["Banner"].replace("PAULMAC's","PAULMAC'S")
  #Aggregate banners
  PRB_aggregate["Banner"] = PRB_aggregate["Banner"].replace('TISOL','PVCI-BC').replace("BOSLEY'S",'PVCI-BC').replace('TOTAL PET','PVCI-BC')
  
  #Aggregate below into PVCI BC
  #Bosleys 
  #Tisol
  #Total Pet
  #PVCI-BC
  
  #Condense the Banners into PVI,PSI,PVCI
  PRB_aggregate["Banner"].replace("PET VALU","PV Canada (PVCI)",inplace = True) #I believe this is the banner that was moved from PVI.
  PRB_aggregate["Banner"].replace("PAULMAC'S","PV Canada (PVCI)",inplace = True)
  PRB_aggregate["Banner"].replace("PVI - EAST","PV USA (PVI)",inplace = True)
  PRB_aggregate["Banner"].replace("PVI - MIDWEST","PV USA (PVI)",inplace = True)
  PRB_aggregate["Banner"].replace("PVCI-BC","PV Canada (PVCI)",inplace = True)
  
  return PRB_aggregate["Banner"]
PRB_aggregate["Banner"] = clean_up_banners(PRB_aggregate)







#PRB_2019_aggregate = pd.concat([PVI_aggregate,PSI_aggregate,PVCI_aggregate]).loc[PRB_aggregate["Fiscal Year"] == 2019]
Clean_PRB_aggregate = PRB_aggregate.dropna(axis=1,how = 'all',thresh = .90*len(PRB_aggregate)) #Drop columns if we don't have at least 50% of the data







Unit_Trend = []
i=0
#for pre_quantity, flyer_quantity, post_quantity in zip(PRB_aggregate["Pre Quantity"],PRB_aggregate["Flyer Quantity"],PRB_aggregate["Post 6 Weeks Quantity"]):
for pre_quantity, flyer_quantity, post_quantity in zip(PRB_aggregate["Regular_units_per_day"],PRB_aggregate["Flyer Quantity Per Day"],PRB_aggregate["Post Quantity Per Day"]):
  #print(pre_quantity,flyer_quantity,post_quantity)
  if (pre_quantity < .8*flyer_quantity) and (flyer_quantity < .8*post_quantity):
    trend = "Continually Rising Units" #pre is less than 80% of flyer
    Unit_Trend.append(trend)
#    print(trend)
#    print(pre_quantity,flyer_quantity,post_quantity)
  elif (pre_quantity *.8 > flyer_quantity) and (flyer_quantity * .8 > post_quantity): #verify later
    trend = "Continually Falling Units"
    Unit_Trend.append(trend)
#    print(trend)
#    print(pre_quantity,flyer_quantity,post_quantity)
  elif (pre_quantity < flyer_quantity * .8) and (pre_quantity * .8 > post_quantity ):
    trend = "Customer Hoarding"
    Unit_Trend.append(trend)
    #print(trend)
#  elif (pre_quantity > flyer_quantity) and (flyer_quantity < post_quantity) and (pre_quantity < post_quantity):
#    trend = "Cannibalized During Flight"
#    Unit_Trend.append(trend)
    #print(trend)    
#  elif (pre_quantity < flyer_quantity) and (flyer_quantity > post_quantity) and (pre_quantity == post_quantity):
#    trend = "Temporary Lift"
#    Unit_Trend.append(trend)
#    #print(trend)
  elif (pre_quantity < flyer_quantity) and (flyer_quantity > post_quantity) and (pre_quantity < post_quantity):
    trend = "Resetting the Bar"
    Unit_Trend.append(trend)
    #print(trend)
#  elif (pre_quantity > flyer_quantity) and (flyer_quantity < post_quantity) and (pre_quantity > post_quantity):
#    trend = "Cannibalized And/Or Lost Customers"
#    Unit_Trend.append(trend)
#    print(trend)
  else:
    trend = "No Trend"
    Unit_Trend.append(trend)
    #print(trend)
  #print(i)
  i=i+1

PRB_aggregate["Units Sold Trend"] = Unit_Trend




#Create Margin Type Classification Column
Margin_Type = []
#for pre_margin,flyer_margin in zip((PRB_aggregate['Flyer Margin Per Unit'] * PRB_aggregate["Flyer Quantity"]),(PRB_aggregate["Regular Margin $ Per Unit"] * PRB_aggregate["Pre Quantity"])):
flyer_margin = PRB_aggregate["Flyer Quantity Per Day"]*PRB_aggregate['Flyer Margin Per Unit'] * PRB_aggregate["Campaign Days"]
pre_margin = PRB_aggregate["Regular Margin $ Per Unit"] * PRB_aggregate["Regular_units_per_day"] * PRB_aggregate["Campaign Days"]
for pre_margin,flyer_margin in zip(pre_margin,flyer_margin):
  #print(pre_margin,flyer_margin)
  if flyer_margin > 1.2 * pre_margin:
    try:
      print((flyer_margin / pre_margin - 1)*100, "%")
      print(pre_margin,flyer_margin)
    except:
      pass
    Margin_Type.append("Strongly Positive Incremental Margin")
  elif (flyer_margin < 1.2 * pre_margin) and (flyer_margin - pre_margin > 0):
    Margin_Type.append("Weak Positive Incremental Margin")
  elif (flyer_margin - pre_margin) < 0:
    Margin_Type.append("Negative Incremental Margin")
#  elif (flyer_margin == pre_margin):  
#    Margin_Type.append("Flat Margin")
  else:
    #print("Error")
    Margin_Type.append("No Trend")
  
PRB_aggregate["Margin_Type"] = Margin_Type


PRB_aggregate.drop_duplicates(inplace=True)

#PRB_aggregate.fillna("Error")

#PRB PIVOT
test = PRB_aggregate.pivot_table(index = ["Units Sold Trend","Banner"],values = values,aggfunc = "sum").reset_index()
test2 = PRB_aggregate.pivot_table(index = ["Units Sold Trend"],values = values,aggfunc = "sum").reset_index()

frequency_table = PRB_aggregate.pivot_table(index = ["Banner","Item Description",], values = ["Fiscal Period"],columns = ["Units Sold Trend","Margin_Type"], aggfunc = "count")
frequency_table_by_page_itemDescription = PRB_aggregate.pivot_table(index = ["Flyer Page","Item Description",], values = ["Fiscal Period"],columns = ["Units Sold Trend","Margin_Type"], aggfunc = "count")
frequency_table_by_subclass_itemdescription = PRB_aggregate.pivot_table(index = ["Sub Class","Item Description",], values = ["Fiscal Period"],columns = ["Units Sold Trend","Margin_Type"], aggfunc = "count")
frequency_table_by_banner_class_subclass = PRB_aggregate.pivot_table(index = ["Banner","Class","Sub Class"], values = ["Fiscal Period"],columns = ["Units Sold Trend","Margin_Type"], aggfunc = "count")
frequency_table_by_banner_page = PRB_aggregate.pivot_table(index = ["Banner","Flyer Page"], values = ["Fiscal Period"],columns = ["Units Sold Trend","Margin_Type"], aggfunc = "count")
frequency_table_by_banner_page_division = PRB_aggregate.pivot_table(index = ["Banner","Flyer Page","Division"], values = ["Fiscal Period"],columns = ["Units Sold Trend","Margin_Type"], aggfunc = "count")


##Tables for presentation
PRB_Flyer_Categ_Division = PRB_aggregate.pivot_table(index = ["Flyer Page","Category","Division"], values = ["Fiscal Period"],columns = ["Units Sold Trend","Margin_Type"], aggfunc = "count")
PRB_Category_Division = PRB_aggregate.pivot_table(index = ["Category","Division"], values = values,columns = ["Units Sold Trend","Margin_Type"], aggfunc = "sum")
PRB_Page_MarginType = PRB_aggregate.pivot_table(index = ["Flyer Page","Margin_Type"], values = ["Fiscal Period"],columns = ["Units Sold Trend"], aggfunc = "count")
#Skype Pivot Request
Brand_Page_Time_Division = PRB_aggregate.pivot_table(index = ["Brand","Flyer Page","Fiscal Year","Division"], values = ["Fiscal Period"],columns = ["Units Sold Trend","Margin_Type"], aggfunc = "count")

#Banner/Page/Item Category
Distribution_by_Banner_Page_Item_Category = PRB_aggregate.pivot_table(index = ["Banner","Flyer Page","Sub Class"], values = ["Fiscal Period"],columns = ["Units Sold Trend","Margin_Type"], aggfunc = "count")








#PRB SUMMARY PIVOT TABLES
values = ["Placement Fee",'Total Deal Amount', "Regular_units_per_day","Flyer Quantity Per Day",'Post Quantity Per Day',"Following 6 Weeks Units Per day","Total Incremental Sales","Total Incremental Margin",
              "Total Campaign Sales (Net of Deal)","Total Vendor Reimbursement","Total COGS During Campaign",'Flyer Margin Dollars',"Incremental_unit_lift_per_day","Incremental Units During Flight",
              "Incremental Units After Flight","Post Incremental_unit_lift_per_day","Incremental Margin After Reimbursement","Estimated_Regular_Margin_During_Campaign"]

#Executive Summary Tables
PRB_ExecSmmryMetrics_by_Banner = PRB_aggregate.pivot_table(index = ["Fiscal Year","Fiscal Period","Banner"], columns = [], values = values,  aggfunc = "sum").reset_index()
PRB_ExecSmmryMetrics_by_Fisc_Period = PRB_aggregate.pivot_table(index = ["Fiscal Year","Fiscal Period"], columns = [], values = values, aggfunc = "sum").reset_index()
PRB_Summary_by_Vendor = PRB_aggregate.pivot_table(index = ["VendorName"], columns = [], values = values, aggfunc = "sum").reset_index()
PRB_Summary_by_Class = PRB_aggregate.pivot_table(index = ['Class'], columns = [], values = values, aggfunc = "sum").reset_index()
PRB_Summary_by_Division = PRB_aggregate.pivot_table(index = ['Division'], columns = [], values = values, aggfunc = "sum").reset_index()
PRB_Summary_by_Department = PRB_aggregate.pivot_table(index = ['Department'], columns = [], values = values, aggfunc = "sum").reset_index()
PRB_ExecSmmryMetrics_by_Fiscal_PD_Brand = PRB_aggregate.pivot_table(index = ["Fiscal Year","Fiscal Period","Brand"], columns = [], values = values, aggfunc = "sum").reset_index()
PRB_ExecSmmryMetrics_by_Margin_Type = PRB_aggregate.pivot_table(index = ["Fiscal Year","Fiscal Period"], columns = "Margin_Type", values = values,  aggfunc = "sum").reset_index()
PRB_ExecSmmryMetrics_by_Margin_Type_Brand_Division_Time = PRB_aggregate.pivot_table(index = ["Fiscal Year","Fiscal Period","Margin_Type","Brand","Division"], columns = [], values = values,  aggfunc = "sum").reset_index()

#Other Tables
PRB_Consumables_by_Fiscal_PD_Brand = PRB_aggregate.loc[PRB_aggregate["Division"] == "Consumables"].pivot_table(index = ["Fiscal Year","Fiscal Period","Brand"], columns = [], values = values, aggfunc = "sum").reset_index()
PRB_Consumables_by_Fiscal_PD_Brand_Banner = PRB_aggregate.loc[PRB_aggregate["Division"] == "Consumables"].pivot_table(index = ["Fiscal Year","Fiscal Period","Brand","Banner"], columns = [], values = values, aggfunc = "sum").reset_index()
PRB_Summary_by_Class_Fiscal_PD_Brand = PRB_aggregate.pivot_table(index = ["Fiscal Year","Fiscal Period",'Class'], columns = [], values = values, aggfunc = "sum").reset_index()

PRB_Summary_by_Division_Trend = PRB_aggregate.pivot_table(index = ["Margin_Type","Units Sold Trend",'Division'], columns = [], values = values, aggfunc = "sum").reset_index()
PRB_Summary_by_Department_Trend = PRB_aggregate.pivot_table(index = ["Margin_Type","Units Sold Trend",'Department'], columns = [], values = values, aggfunc = "sum").reset_index()
PRB_Summary_by_Class_Trend = PRB_aggregate.pivot_table(index = ["Margin_Type","Units Sold Trend",'Class'], columns = [], values = values, aggfunc = "sum").reset_index()
PRB_Summary_by_Subclass_Trend = PRB_aggregate.pivot_table(index = ["Margin_Type","Units Sold Trend",'Sub Class'], columns = [], values = values, aggfunc = "sum").reset_index()
PRB_Summary_by_Item_Trend = PRB_aggregate.pivot_table(index = ["Margin_Type","Units Sold Trend",'Item'], columns = [], values = values, aggfunc = "sum").reset_index()
#Requested by Scott
#PRB_Summary_by_Discount = PRB_aggregate.pivot_table(index = ["Offer Notes","Division","Class","Sub Class","BRAND"], columns = [], values = values, aggfunc = "sum").reset_index()
PRB_Summary_by_Discount = PRB_aggregate.pivot_table(index = ["Offer Notes","Fiscal Year","Fiscal Period","Banner","Flyer Page","Department","Division","Class","Sub Class","Brand","Margin_Type",'Sale Price','Vendor BB', 'Reg. Price'], columns = [], values = values, aggfunc = "sum").reset_index()


#EXPORT All TABLES
def export_PRB_pivot_tables():
  f_table_directory = "C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Program Output/PRB"
  #Append new tables here
  table_name_list = ["frequency_table","frequency_table_by_page_itemDescription",
                     "frequency_table_by_subclass_itemdescription","frequency_table_by_banner_class_subclass",
                     "Distribution_by_Banner_Page_Item_Category","frequency_table_by_banner_page_division",
                     "PRB_Flyer_Categ_Division","PRB_Category_Division","PRB_Page_MarginType","Brand_Page_Time_Division",
                     "PRB_ExecSmmryMetrics_by_Banner","PRB_ExecSmmryMetrics_by_Fisc_Period","PRB_Summary_by_Vendor","PRB_Summary_by_Class",
                     "PRB_Summary_by_Division","PRB_ExecSmmryMetrics_by_Fiscal_PD_Brand","PRB_ExecSmmryMetrics_by_Margin_Type",
                     "PRB_Consumables_by_Fiscal_PD_Brand","PRB_Summary_by_Class_Fiscal_PD_Brand","PRB_Summary_by_Division_Trend",
                     "PRB_Summary_by_Department_Trend","PRB_Summary_by_Department_Trend","PRB_Summary_by_Class_Trend",
                     "PRB_Summary_by_Subclass_Trend","PRB_Summary_by_Item_Trend","PRB_Summary_by_Department","PRB_Summary_by_Discount",
                     "PRB_ExecSmmryMetrics_by_Margin_Type_Brand_Division_Time"]
  
  table_list = [eval(x) for x in table_name_list]
  for table,name in zip(table_list,table_name_list):
    table.to_excel(f"{f_table_directory}/{name}.xlsx")

#export_PRB_pivot_tables()            #REACTIVATE THIS TO EXPORT TABLES


#DROP THESE ROWS
#PRB_aggregate.loc[PRB_aggregate["Division"] == "DOG"]




#ENTIRE HEIRARCHY
#Division/Department/Class/Subclass/item



#Keep everything except fiscal period 7 and beyond. This will help us backtest period 7.
PRB_aggregate_date_filtered = PRB_aggregate.loc[~((PRB_aggregate["Fiscal Period"] >= 7) & (PRB_aggregate["Fiscal Year"] == 2019))]
#Create margin_type history count pivot table by item
Margin_Type_History = PRB_aggregate_date_filtered.groupby("Item")['Margin_Type'].value_counts().unstack(fill_value=0)
#Margin_Type_History.columns = ["Historical_Count " + x for x in Margin_Type_History.columns]
Margin_Type_History.to_csv("C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Promotion Level Rollup/Promotion Line History Count.csv")


#Append each item's margin history to the aggregate data table. Probably a stupid idea since you can't pivot the table now.
#PRB_aggregate = pd.merge(PRB_aggregate,Margin_Type_History, on = "Item")


#Exports For discussion/quick pivots/graphing
PRB_aggregate.to_csv("C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Program Output/All Data.csv")
  
Export_Columns = ["Fiscal Period", "Fiscal Year","Division","Department","Class","Sub Class","Placement Fee",'Total Deal Amount', "Regular_units_per_day","Flyer Quantity Per Day","Following 6 Weeks Units Per day","Total Incremental Sales","Total Incremental Margin",
              "Total Campaign Sales (Net of Deal)","Total Vendor Reimbursement","Total COGS During Campaign",'Flyer Margin Dollars',"Incremental_unit_lift_per_day","Incremental Units During Flight",
              "Incremental Units After Flight","Post Incremental_unit_lift_per_day","Incremental Margin After Reimbursement","Brand","Units Sold Trend",
              "Margin_Type","Banner","Reg. Price","Sale Price","Vendor BB",
              "Item Description","Flyer Page","Offer Notes","Item","Campaign Days","Reimbursement for Incremental Units","Flyer Margin Dollars Before Reimbursement"]
            #Removed Percentage reimbursement due to bugginess
            
export = PRB_aggregate[Export_Columns]
export["Item Placements"] = len(PRB_aggregate)
export.to_csv("C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Program Output/Light Data Dump for Pivoting.csv")   #For loading up elsewhere in python
export.to_excel("C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Program Output/Light Data Dump for Pivoting.xlsx") #For tableau

















#Price buckets based price before discount during preperiod and divide by 10 and take the integer.

#Pivot by:
#Banner (do later), page, Division, department, class, subclass



#Next steps
#Goal is to decide what we're going to continue puting on the flyer.   

#Filter empty items to reduce size?

#Subclass is support data for the appendix
#We need to understand what items are driving desired behavior & which are not. The questions that will be asked:
#Does cat food & dogfood drive incrementality?
#Cat trees/bowl drive incrementalty?
#Department = Dog/Cat 


#Consumables, cat/dog, treat or food, natural/enhanced (subclass)/basic (class)
#Go down to subclass for consumables on distribution table. 
#For the rest, class
  
#Powerpoint class



#Non-Linear regression such as "Step wise" regression or "logistic regression"
#See this library: https://pypi.org/project/stepwise-regression/
#We want to figure out the DRIVER variables of incremental margin. Which variables drive it the most/least? Change in sales price? Etc. Our dependent variable is incremental margin


#VISUALS SECTION
#
#
#
#from mpl_toolkits.mplot3d import Axes3D
#os.chdir("C:/Users/mfrangos/Desktop/Marketing Analytics 2/Product Analysis with Mark Furry/Flyer Data/PSI/Visuals")
#for sku in aggregate["SKU #"]:
#    data = aggregate.loc[aggregate["SKU #"] == sku]
#    
#    fig = plt.figure()
#    ax = fig.add_subplot(111, projection='3d')   
#                      
#    x = "Fiscal Period"
#    y = "Follow 6 $"
#    z =  "Flyer $"
#     
#    ax.scatter(xs = data[f"{x}"], 
#                ys = data[f"{y}"],
#                zs = data[f"{z}"],
#                #c = data["Division"],
#                #s = (Merged_Pet_Spa_Data.loc[Merged_Pet_Spa_Data["ACCOUNT NAME"]==account]["GL AMOUNT"])*.10  ,    #SIZE
#                alpha = 1 #Transparency                                     
#              ) 
#    plt.title(f"{sku}")
#    plt.xlabel(f"{x}")
#    plt.ylabel(f"{y}")
#    ax.set_zlabel(f"{z}")
#    plt.savefig(f"_{sku}.png")
#    plt.close('all')    
#
#
#
#  
#
#    
#####VISUALIZE CORRELATIONS``
#Corr_Data0 = aggregate.fillna(0)
##For account groups
##Corr_Data = Corr_Data0.pivot_table(index = ["GL YEAR", "PERIOD"], columns = ["Account Group","ACCT #","ACCOUNT NAME"], values = "GL AMOUNT", aggfunc = "sum").fillna(0)
##For store sales correlations
#Corr_Data = Corr_Data0.pivot_table(index = ["GL YEAR", "PERIOD"], columns = ["DEPT #"], values = "GL AMOUNT", aggfunc = "sum")#.fillna(0)                                  
#                                   
#Correlation_Table=Corr_Data0.corr()
#Correlation_Table.to_excel("Correlation Table.xlsx")
#len(Corr_Data.corr())
#len(Corr_Data.corr().columns)
#
##VISUALIZE
#
#import numpy as np
#def Visualize():
#  Corr_Data1 = Correlation_Table.values
#  fig1 = plt.figure()
#  ax1 = fig1.add_subplot(111)    
#  heatmap1 = ax1.pcolor(Corr_Data1, cmap=plt.cm.RdYlGn)
#  fig1.colorbar(heatmap1)
#  ax1.set_xticks(np.arange(Corr_Data1.shape[1]) + 0.5, minor=False)
#  ax1.set_yticks(np.arange(Corr_Data1.shape[0]) + 0.5, minor=False)
#  #Flip y axis
#  ax1.invert_yaxis()
#  ax1.xaxis.tick_top()
#  column_labels = Correlation_Table.columns
#  row_labels = Correlation_Table.index
#  ax1.set_xticklabels(column_labels)
#  ax1.set_yticklabels(row_labels)
#  ax1.tick_params(axis="x", labelsize=5)
#  ax1.tick_params(axis="y", labelsize=5)
#  plt.xticks(rotation=90)
#  heatmap1.set_clim(-1,1)
#  plt.tight_layout()
#  plt.savefig(f"Correlations", dpi = (1600))
#  plt.show()
#
##Visualize()
##Correlation_Table.to_excel("Correlation.xlsx")



  

