#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import math
import numpy as np


# In[2]:


level_col = 0
item_col = 1
desc_col = 2
unit_col = 6
total_col = 7


# In[3]:


#Read in Excel File to get sheet names
#excel = pd.ExcelFile("Tender Schedule (NEW Branding).xlsm")
excel = pd.ExcelFile("Tender_Schedule_test.xlsm")
sheets = excel.sheet_names


# In[4]:


df_forms = pd.DataFrame(columns = {"Form ID", "Form Name"})


# In[5]:


#Read pages and convert to Data Frames
#Data Frames stored in DF_List
df_list = list()
inc = 0
for x in sheets :
    df_list.append(pd.read_excel("Tender_Schedule_test.xlsm", sheet_name=x))
    form = [[inc,x]]
    df_forms = df_forms.append(pd.DataFrame(form,columns = ["Form ID", "Form Name"]),ignore_index = True)
    inc += 1


# In[6]:


print(df_forms)


# In[7]:


df_col = ["Form ID", "Level 1", "Level 2", "Level 3", "Level 4", "Level 5", "Item ID","Description","Units"]
df = pd.DataFrame(columns = df_col)
print(df)


# In[8]:


for x in range(0,len(sheets)-1):
    inside = False
    for index, rows in df_list[x].iterrows():
        if inside :
            #Check if at end of table
            if rows[total_col] == "Total" :
                inside = False
            else :
                hold = ""
                df_hold = ["","0","0","0","0","0","","",""]
                full_count = 0
                rows[item_col] = str(rows[item_col])
                if rows[item_col] != "nan":
                    count = 0
                    length = len(rows[item_col])
                    for i in str(rows[item_col]):
                        count += 1
                        if i != ".":
                            hold = hold + i
                            if count == length:
                                df_hold[full_count + 1] = hold
                        else:
                            df_hold[full_count + 1] = hold
                            full_count += 1
                            hold = ""
                    df_hold[7] = rows[desc_col]
                    df_hold[0] = df_forms['Form ID'][x]
                    for y in range(0,6):
                        df_hold[6] = str(df_hold[6]) + str(df_hold[y]) + "."
                    df_hold[6] = df_hold[6][:-1]
                    df_hold[8] = rows[unit_col]
                    #PUSH TO SQL HERE
                    s_row = pd.Series(df_hold, index=df.columns)
                    df = df.append(s_row,ignore_index=True)
                    print(df_hold)       

        if rows[item_col] == "Item" :
            inside = True


# In[ ]:





# In[ ]:





# In[9]:


df.to_csv("test.csv")


# In[10]:


df_forms.to_csv("Form_IDs.csv")


# In[ ]:




