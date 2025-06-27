# -*- coding: utf-8 -*-
"""
Created on Thu Jun 12 11:31:38 2025

@author: MID80

Converts raw eLabNext download into FLAME study requested format. For Marsland Lab. 
"""

import pandas as pd
from datetime import datetime
import os 

def get_report(name, sheetName = 0): 
    if not os.path.exists(name) and os.path.isfile(name):
        print( "File does not exist. Check file name and path are correct.\n")
        try_again()
        # sys.exit()
    else :
        file = pd.ExcelFile(name)
        
    if (sheetName == 0) or sheetName in file.sheet_names : 
        df = pd.read_excel(name, sheet_name = sheetName)
    else :
        print( "The sheet name you entered does not exist." + 
              f"The following sheetnames are valid names for your file:\n{file.sheet_names}\n")
        try_again()
        # sys.exit()
        
    if len(df.columns) == 0  : 
        print( "File has 0 columns. Check file name and path are correct.") 
        try_again()
    return df
    
 
def save_report(df, name):
    
    # lname = name.rsplit(sep = "/", maxsplit = 1)[1]
    oname = name.rsplit(sep = "/", maxsplit = 1)[0]
    files = os.listdir(oname)
    newname = name.replace(".xlsx", f"_clean_{datetime.today().strftime('%Y_%m_%d')}.xlsx")
    lname = newname.rsplit(sep ="/", maxsplit = 1)[1]
    if lname in files : 
        newname = name.replace(".xlsx", f"_clean_{datetime.today().strftime('%Y_%m_%d_%H%M')}.xlsx")
        
    df.to_excel(newname, #name.replace(".xlsx", f"_clean_{datetime.today().strftime('%Y_%m_%d')}.xlsx"), 
                sheet_name = "FLAME Report", index = False)
    print("Saved report as " + newname)
    return 

def clean_report(df): 
    df2 = df.loc[:, ["Name", "Sample type", "Location path", "Location",  "Date"]]
    df2 = df2.rename(columns = {"Location" : "Box Name"})
    df2['Freezer Name'] = df2['Location path'].str.split(" - ", n = 0, expand = True)[0]
    
    df2["Name"] = [x if isinstance(x, str) and 'F' in x else f"F{x}" for x in df2["Name"].tolist()]
    df2["Name"] = [a.strip() if a == 1 else a.strip() for a in df2["Name"].tolist()]
    df2["Name"] = df2["Name"].str.split(n = 1, expand = True)[0]
    ogsite = {"F1" : "Pitt", "F2" : "KUMC",  "F3" : "NEU"}
    df2['Original Site'] =  df2['Name'].str[:2].map(ogsite) 
    cursite = { "NEU" : "NEU", "KUMC" : 'KUMC', 'Kirk' : 'Pitt', 'Chest' : 'Pitt', 'In' : "In Transit"}
    df2['Current Site'] = df2["Freezer Name"].str.split(n = 1, expand = True)[0].map(cursite)
    
    df3 = df2.drop("Location path", axis = 1)
    df4gp = df3.groupby(["Name", "Sample type"], as_index = False)
    df4 = df4gp.nth(0)
    df4ct = df4gp.agg("count").iloc[:, :3]
    df4ct = df4ct.rename(columns = {df4ct.columns[2]:"count"})
    df4 = pd.merge(df4, df4ct, how = "inner", on = ["Name", "Sample type"] )
    df4['Sample type'] = [f"{df4['Sample type'][x]} ({df4['count'][x]})" if df4['count'][x] > 1 else df4["Sample type"][x] 
                            for x in range(0, len(df4['count'])) ]
    df4 = df4.drop('count', axis = 1)
    return df4

def check_cols(df) :
    cn = df.columns
    cn2 = cn.copy().to_list()
    for i in range(0, len(cn) -1): 
        cn2[i] = str.lower(cn2[i]) 
        cn2[i] = cn2[i].replace(" ", "")
    tf = ["Name", "Sample type", "Location path", "Location", "Date"] 
    altf = {"Name" : ["ID", "Sample ID"], 
            "Sample type" : [], "Location path" : ["Path"], "Location" : ["Box Name", "Box"], "Date" : []}
    print("Your column names are\t")
    print( cn.to_list() ) 
    for i in tf :
        if not i in cn :
            if not str.lower(i).replace(" ","") in cn2 :
                othername = False 
                if len(altf[i]) != 0 : 
                    for j in range(0, len(altf[i]) - 1) : 
                        if altf[i][j] in cn : 
                            df = df.rename(columns = {altf[i][j] : i })
                            othername = True 
                            break  
                        elif str.lower(altf[i][j]) in cn : 
                            df = df.rename(columns = {str.lower(altf[i][j]) : i})
                            othername = True 
                            break
                                
                if othername == False : 
                    print( '\nColumn "' + i + '" expected but not included. ')
                    an = input ("If you know this column is actually named something else," + 
                                "\n A: type the actual column name and hit Enter to continue, OR" + 
                                "\n B: press Enter to stop the program to update names. \n")
                    if an == "" : 
                        print("PLEASE UPDATE NAMES AND RESTART" + 
                              ''' REQUIRED COLUMNS ARE: \n\t "Name", "Sample type", "Location path", "Location", "Date" ''') 
                        try_again() # end task here 
                    elif not an in cn : 
                        print( f'''The name you entered ("{an}") does not exist in your data. Please check data.''' + 
                              ''' \n REQUIRED COLUMNS ARE: \n\t "Name", "Sample type", "Location path", "Location", "Date" ''')
                        try_again()
                        # end task here 
                    else: 
                        df = df.rename(columns = {an : i})
            else: 
                ind = cn2.index(str.lower(i).replace(" ",""))
                df = df.rename(columns = {df.columns[ind] : i})
    return df
    

def try_again():
    input("\nPress Enter to restart task, or Ctrl+c to exit program.\n")
    main()

def main(): 
    print( "This program generates eLabNext report for the FLAME study from downloaded data.")
    print( 'Input the .xlsx file that includes required columns "Name", "Sample type", "Location path", "Location", "Date"')
    name = input("Paste file path here (including <fileName>.xlsx): ")
    #  r"C:\Users\MID80\OneDrive - University of Pittsburgh\Desktop\Code Genreal\eLabNext\elabinventory_RAW" 
    name = name.replace("\'", "")
    name = name.replace("\"", "")
    name = name.replace("\\", "/")
    name = name.strip()
    if not name[(len(name) - 5 ):] == '.xlsx' :
        name += '.xlsx'

        
    sheetName = input("If multiple sheets, input Sheet Name with results to clean here\n or press enter: ")
    if sheetName == "" :
        sheetName = 0
    print("Importing data...")
    df1 = get_report(name, sheetName)
    print("Done.")
    
    print("Checking for required columns...")
    df1 = check_cols(df1)
    print("Done.")

    print("Cleaning Data...")
    df_clean = clean_report(df1)
    print("Done.")

    print("Saving Data...")   
    save_report(df_clean, name) 
    
    try_again()
    return 

  
if __name__ == '__main__':
    main()


