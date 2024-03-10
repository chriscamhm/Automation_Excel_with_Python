#!/usr/bin/env python
# coding: utf-8

# <font size="6"><h1><center> **Upload Sharepoint list modifications to Excel**</center></h1>
# <font size="3">This workbook is used to upload Sharepoint list modifications to Excel. The only parameters that need to be changed in the book, when modifications are loaded, are the locations of the two files used, the Portfolio of Social Investment file and the Excel file that is connected to Sharepoint.
#  

# In[11]:


import pandas as pd
import numpy as np
pd.set_option("display.max_columns", None) #Para desplegar todas las columnas


# In[12]:


port_act=pd.read_excel('Portafolio Desarrollo Sostenible 23 de noviembre 2023 V0.xlsx', sheet_name='Portafolio')
#Delete the nulls or those last rows that are sometimes in the portfolio that do not correspond to any project.
#Use the Consecutive column as a reference to delete that
port_act.dropna(how='all',inplace=True, subset=['Consecutivo'])
#This line deletes the last row that corresponds to the row that adds the totals of the excel table
port_act = port_act.iloc[:-1, :]


# **Excel file connected to sharepoint:** This file is called on my personal computer <em>"Portfolio List"</em> and it is an excel file that is connected to Sharepoint. Before using the Python code, that file must be updated in Data, Update Data.

# In[13]:


port_new=pd.read_excel('Portafolio Lista_Cristian.xlsx', sheet_name='Inclusión Proyectos-Iniciativas')
port_new.drop(['Editor.Title'], axis=1, inplace=True)


# The next line extracts the names of the columns in a list and then adds an "(a)" to another list and then places the name of the columns with that new list in the Portfolio List file. This is so that when you merge with the Portfolio, you can differentiate between the Portfolio columns and those of the Portfolio List file that is connected to Sharepoint.

# In[14]:


columnas= port_new.columns.to_list() #Takes the column names and passes them to a list
new_list = [x+"(a)" if x!='ID' else x for x in columnas] #With a new list It gives them one (a) except for the ID one
port_new.columns=new_list #rename the Sharepoint portfolio columns with the new names
port_new


# In Sharepoint it turns out that when they put "Not Applicable" in the form, the column is not actually filled with "Not Applicable" but with a date, January 1, 2000, so I identify that those are "Not Applicable". This is done because in Sharepoint you cannot place text in a date column. Now, in the next line of code I replace those with January 1, 2000 with blank values so that they do not appear in the portfolio and start giving negative comments to the update.

# In[15]:


port_new.replace("2000-01-01", np.nan, inplace=True)
port_new.replace("2001-12-31", np.nan, inplace=True)
port_new


# <font size="4"><h1><left> **Merge Sharepoint and Normal Portfolio**</left></h1> <br>
# <font size="3"> Now we merge between Portfolio and Sharepoint Portfolio. We added two new columns that correspond to the observations that indicate the changes from how it was in Normal Portfolio and after the update.

# In[16]:


#Merge left to cross those with ID_Sharepoint in Portfolio
df=port_act.merge(port_new, left_on=['ID_Sharepoint'],right_on=['ID'], how='left')
df['Obs']='' #Observations of changes in other columns that are not dates
df['Obs_fec']='' #Observations of changes in date columns
df


# In the next line we remove the columns that do not go into the portfolio, such as modifications, created by, ID-Calculated and PPF Observations. Then, we separate the date columns since these are the ones that update the most and require looking at them separately.

# In[17]:


not_columns=[
 'Modified',
 'Created', 'Observaciones PPF',
            'ID_Calculado']

fechas=['Fecha Inicio de Proyecto',
 'Fecha Fin de Proyecto',
 'Fecha Plan Comité Proyectos VDS',
 'Fecha Real Comité Proyectos VDS',
 'Fecha Plan Comité Convenios / Contratos VDS',
 'Fecha Real Comité Convenios / Contratos VDS',
 'Fecha plan comité GCF / Radicado ABA',
 'Fecha real comité GCF / real Radicado ABA',
 'Fecha de suscripción plan',
 'Fecha de suscripción real',
 'Fecha Plan de firma Acta de Inicio ',
 'Estado del proyecto']
for i in not_columns:
    columnas.remove(i)
for i in fechas:
    columnas.remove(i)


# The following line of code is the most important and whose logic also applies to the codes to enter accrual in Portfolio. Basically the logic is the following:
# * We return to the list **Columns** which contains the names of the columns.
# - With this list we iterate on each of its elements except for some which should not change or whose change is made by talking to Iván, example Grouping level 4.
# + The first part of the iteration corresponds to the change observations.
# * these observations remain blank if the status of the project is Canceled or its ID is null, that is, it is not in Sharepoint, or the Normal Portfolio and Sharepoint Portfolio information is the same indicating that there is no change, or that either of the two columns be null or blank. If it does not meet these conditions, then column e will be filled, taking the information from the normal Portfolio, indicating the name of the column and saying that it will change to a value x which will be the Sharepoint Portfolio.
# * The second part of the iteration corresponds to the update in the column and has the same conditions as in Observations, only this time the column will change to the Sharepoint values if it meets the conditions.
# * The same process is done with dates, the only difference is in some values.

# In[18]:


for i in columnas:
    if i != 'ID' and i != 'Consecutivo' and i!='Línea de Inversión' and i != 'Agrupación Nivel 4' and i != 'Agrupación Nivel 1' and i != 'Agrupación Nivel 2' and i != 'Agrupación Nivel 3' and i != 'Tipo Inversión' and i != 'Eje de Inversión':
        df['Obs']= df.apply(lambda x: x['Obs'] if ((x['Estado del proyecto'] == 'Cancelado')|
                                                                           (pd.isnull(x['ID']))|
                                                                            ((x['{}'.format(i)])==(x['{}(a)'.format(i)]))|
                                                   (str(x['{}'.format(i)])==str(x['{}(a)'.format(i)]))|
                                                    (pd.isnull(x['{}(a)'.format(i)]) and x['{}'.format(i)]=='No aplica')|
                                                          (pd.isnull(x['{}(a)'.format(i)]) and x['{}'.format(i)]=='Por definir')|
                                                  (pd.isnull(x['{}'.format(i)]) and pd.isnull(x['{}(a)'.format(i)]))|
                                                 ((x['{}'.format(i)])==0 and pd.isnull(x['{}(a)'.format(i)])))
                                           else (x['Obs']+","+str('{}'.format(i))+": "+str(x['{}'.format(i)])+" cambió a "+
                                                str('{}(a)'.format(i))+": "+str(x['{}(a)'.format(i)]))
                                                 , axis=1)
        df['{}'.format(i)]=df.apply(lambda x: x['{}'.format(i)] if (
                                                            (x['Estado del proyecto'] == 'Cancelado')|
                                                                           (pd.isnull(x['ID']))|
                                                                            ((x['{}'.format(i)])==x['{}(a)'.format(i)])|
            (str(x['{}'.format(i)])==str(x['{}(a)'.format(i)]))|
             (pd.isnull(x['{}(a)'.format(i)]) and x['{}'.format(i)]=='No aplica')|
                                                          (pd.isnull(x['{}(a)'.format(i)]) and x['{}'.format(i)]=='Por definir')|
        (pd.isnull(x['{}'.format(i)]) and pd.isnull(x['{}(a)'.format(i)]))|
                                                 ((x['{}'.format(i)])==0 and pd.isnull(x['{}(a)'.format(i)])))
                                           else x['{}(a)'.format(i)], axis=1)
        
for i in fechas:
    if i != 'ID' and i != 'Consecutivo':
        df['Obs_fec']=df.apply(lambda x: x['Obs_fec'] if ((x['Estado del proyecto'] == 'Cancelado')|
                                                                           (pd.isnull(x['ID']))|
                                                                            ((x['{}'.format(i)])==x['{}(a)'.format(i)])|
                                                                            (pd.isnull(x['{}(a)'.format(i)]) and x['{}'.format(i)]=='No aplica')|
                                                          (pd.isnull(x['{}(a)'.format(i)]) and x['{}'.format(i)]=='Por definir')|
                                                          (pd.isnull(x['{}'.format(i)]) and pd.isnull(x['{}(a)'.format(i)]))|
                                                 ((x['{}'.format(i)])==0 and pd.isnull(x['{}(a)'.format(i)])))
                                           else (x['Obs_fec']+","+'{}'.format(i)+": "+str(x['{}'.format(i)])+" cambió a "+
                                                '{}(a)'.format(i)+": "+str(x['{}(a)'.format(i)]))
                                                 , axis=1)
        df['{}'.format(i)]=df.apply(lambda x: x['{}'.format(i)] if (
                                                            (x['Estado del proyecto'] == 'Cancelado')|
                                                                           (pd.isnull(x['ID']))|
                                                                            ((x['{}'.format(i)])==x['{}(a)'.format(i)])|
                                                                            (pd.isnull(x['{}(a)'.format(i)]) and x['{}'.format(i)]=='No aplica')|
                                                                            (pd.isnull(x['{}(a)'.format(i)]) and x['{}'.format(i)]=='Por definir')|
                                                                            (x['{}(a)'.format(i)]=='2000-01-01')|
            (pd.isnull(x['{}'.format(i)]) and pd.isnull(x['{}(a)'.format(i)]))|
                                                 ((x['{}'.format(i)])==0 and pd.isnull(x['{}(a)'.format(i)])))
                                           else x['{}(a)'.format(i)], axis=1)


# In[19]:


#Reattach the date column names to the column list
for i in fechas:
    columnas.append(i)
# to the list of columns we add the observations columns
for i in ('Obs','Obs_fec'):
    columnas.append(i)
#The final database will only have the columns that are in the column list
#This way when running the macro it will not have a long list of columns but only the necessary ones
df = df[columnas]
df


# In[20]:


#Export to Excel that new Portfolio with which the macro will be used to paste into the Portfolio excel.
df.to_excel('Modificados.xlsx', sheet_name='Python', index = False)


# In[ ]:





# In[ ]:





# In[ ]:




