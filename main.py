#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import os


# In[11]:


data=pd.read_csv("movies.csv")


# In[12]:


data.mpaa_rating.value_counts()


# In[4]:


#For our output, let's make an excel tab for each mpaa rating.
#where each tab will have the respective top 10 rated movies.


# In[13]:


data.columns


# In[16]:


columns_needed=['title', 'mpaa_rating','release_date','genre', 'rating', 'summary']
data=data[columns_needed]


# In[27]:


pg_movies=data[data['mpaa_rating']=="PG"].sort_values(by='rating',ascending=False).head(10)
r_movies=data[data['mpaa_rating']=="R"].sort_values(by='rating',ascending=False).head(10)
pg13_movies=data[data['mpaa_rating']=="PG-13"].sort_values(by='rating',ascending=False).head(10)
g_movies=data[data['mpaa_rating']=="G"].sort_values(by='rating',ascending=False).head(10)


# In[28]:


def format_excel_tab(df,tab_name,colour="green"):
    """

    :param df: name of dataframe
    :param tab_name: tab in df to be formatted
    :param colour: colour for tab header + column header
    :return: selected tabs is formatted, inplace
    """
    worksheet = writer.sheets[tab_name]
    worksheet.set_tab_color(colour)
    auto_width(df,tab_name)
    header_format(df,tab_name,colour)

def auto_width(df,tab):
    """


    :param df: name of dataframe
    :param tab: tab in df with improper width
    :return: selected tabs' columns' width is auto set, inplace
    """
    for column in df:
        column_width = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets[f'{tab}'].set_column(col_idx, col_idx, column_width)

def header_format(df,tab,colour="green"):
    """

    :param df: name of dataframe
    :param tab: tab in df with improper width
    :return:selected tabs' columns' headers are coloured, inplace
    """
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': colour,
        'border': 1,
        'align':'center'})

    for col_num, value in enumerate(df.columns.values):
        writer.sheets[f'{tab}'].write(0, col_num, value, header_format)


# In[29]:


saved_filename="Movies_Per_Category.xlsx"

with pd.ExcelWriter(saved_filename) as writer:

    workbook  = writer.book

    g_movies.to_excel(writer, sheet_name='G',index=False)
    format_excel_tab(g_movies,"G","orange")

    r_movies.to_excel(writer, sheet_name='R',index=False)
    format_excel_tab(pg_movies,"R","red")

    pg_movies.to_excel(writer, sheet_name='PG',index=False)
    format_excel_tab(pg_movies,"PG") #default green

    pg13_movies.to_excel(writer, sheet_name='PG-13',index=False)
    format_excel_tab(pg13_movies,"PG-13")


# In[30]:


os.startfile("Movies_Per_Category.xlsx", 'open')

