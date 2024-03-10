#!/usr/bin/env python
# coding: utf-8

# In[89]:


import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
get_ipython().run_line_magic('matplotlib', 'inline')


# In[90]:


port_act=pd.read_excel('Portafolio Desarrollo Sostenible 06 de enero 2023.xlsx', sheet_name='Portafolio')
port_act.dropna(how='all',inplace=True, subset=['Consecutivo']) #Quita nulos que haya en la columna de Consecutivo
port_act = port_act.iloc[:-1, :] #Quita la última fila que corresponde a totales en Portafolio
columnas= port_act.columns.to_list() #nombres de las columnas en una lista
port_act


# In[91]:


# Import module for k-protoype cluster. K prototype permite hacer el cluster con valores categóricos y númericos
from kmodes.kprototypes import KPrototypes
# Ignore warnings
import warnings
warnings.filterwarnings('ignore', category = FutureWarning)
# Format scientific notation from Pandas
pd.set_option('display.float_format', lambda x: '%.3f' % x)


# In[92]:


#Normalizo el valor total del proyectos para que valores más grandes o muy pequeños no afecten
min_params = port_act['Valor Total Proyecto'].min()
max_params = port_act['Valor Total Proyecto'].max()
port_act['Valor Normalizado']= (port_act['Valor Total Proyecto'] - min_params) / (max_params- min_params)
port_act


# In[93]:


#esto es para jugar con las variables y quitar las que no usaría en el clustering
df=port_act.drop(['Consecutivo', 'Nombre corto del proyecto','Valor Total Proyecto', "Activo 1", "Departamento","Tipo Inversión", "Agrupación Nivel 1"], axis = 1)
df


# In[94]:


# Get the position of categorical columns
catColumnsPos = [df.columns.get_loc(col) for col in list(df.select_dtypes('object').columns)]
print('Categorical columns           : {}'.format(list(df.select_dtypes('object').columns)))
print('Categorical columns position  : {}'.format(catColumnsPos))


# In[95]:


# Convert dataframe to matrix
dfMatrix = df.to_numpy()
dfMatrix


# In[96]:


#print length of each array
print(len(dfMatrix))


# In[104]:


# Choose optimal K using Elbow method. Método para elegir el número óptimo de cluster
cost = []
for cluster in range(1, 20):
    try:
        kprototype = KPrototypes(n_jobs = -1, n_clusters = cluster, init = 'Huang', random_state = 0)
        kprototype.fit_predict(dfMatrix, categorical = catColumnsPos)
        cost.append(kprototype.cost_)
        print('Cluster initiation: {}'.format(cluster))
    except:
        break
# Converting the results into a dataframe and plotting them
df_cost = pd.DataFrame({'Cluster':range(1, 20), 'Cost':cost})
# Data viz
plt.figure(figsize=(16,8))
plt.plot(df_cost['Cluster'], df_cost['Cost'], 'bx-')
plt.xlabel('k')
plt.ylabel('Distortion')
plt.title('The Elbow Method showing the optimal k')
plt.show()


# In[105]:


# Fit the cluster
kprototype = KPrototypes(n_jobs = -1, n_clusters = 8, init = 'Huang', random_state = 0)
kprototype.fit_predict(dfMatrix, categorical = catColumnsPos)


# In[106]:


# Cluster centroid. Los centros de cada cluster
kprototype.cluster_centroids_


# In[107]:


# Check the iteration of the clusters created
kprototype.n_iter_


# In[108]:


# Check the cost of the clusters created
kprototype.cost_


# In[109]:


# Add the cluster to the dataframe. Nombrar cada cluster para identificarlos
df['Cluster Labels'] = kprototype.labels_
df['Segment'] = df['Cluster Labels'].map({0:'First', 1:'Second', 2:'Third', 3:'Fourth', 4:'Fifth', 5:'Sixth', 6:'Seventh',
                                          7:'Eighth'})
# Order the cluster
df['Segment'] = df['Segment'].astype('category')
df['Segment'] = df['Segment'].cat.reorder_categories(['First','Second','Third','Fourth','Fifth'
                                                      ,'Sixth','Seventh','Eighth'])


# In[110]:


df


# In[80]:


df['Segment'].value_counts()


# In[81]:


#Unir esos nombres con los proyectos
df=pd.merge(df, port_act[['Consecutivo','Nombre del Proyecto', 'Valor Total Proyecto']], left_index=True, right_index=True)


# In[ ]:


#Exportar a excel
df.to_excel('Clustering.xlsx', sheet_name='Python', index = False)


# In[ ]:




