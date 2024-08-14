#!/usr/bin/env python
# coding: utf-8

# In[17]:


#BIBLIOTECAS


import json

import pandas as pd
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer
from sklearn.tree import DecisionTreeClassifier
from sklearn.tree import plot_tree
import matplotlib.pyplot as plt

from sklearn.model_selection import train_test_split
from sklearn.impute import SimpleImputer
from sklearn.metrics import mean_squared_error
import numpy as np
from sklearn.metrics import mean_squared_error
from sklearn.model_selection import train_test_split


from sklearn.ensemble import RandomForestRegressor


# In[20]:


#LER O DATAFRAME NA FONTE
import pandas as pd
df = pd.read_excel(r'C:\Users\B28166\Desktop\EXCEL RELAORIOS DE AVALIAÇÃO\Excel_corretos.xlsx')


# In[34]:


df


# In[36]:


# Supondo que df é o seu DataFrame

# Número total de documentos únicos
total_docs = df['document'].nunique()

# Contar a frequência de cada campo em documentos únicos
campo_counts = df.groupby('key')['document'].nunique()

# Calcular a percentagem
campo_percentages = (campo_counts / total_docs) * 100

# Adicionar a percentagem ao DataFrame original
df['percentagem_chave'] = df['key'].map(campo_percentages)


# In[37]:


df


# In[45]:


# Número total de documentos únicos
total_docs = df['document'].nunique()

# Contar a frequência de cada key em documentos únicos
key_counts = df.groupby('key')['document'].nunique()

# Calcular a percentagem
key_percentages = (key_counts / total_docs) * 100

# Calcular a mediana de 'acc' e 'confidence' por 'key'
median_df = df.groupby('key')[['confidence', 'acc']].median().reset_index()

# Adicionar a percentagem ao DataFrame de medianas
median_df['percentagem_chave'] = median_df['key'].map(key_percentages)

# Renomear as colunas para uma leitura mais clara
median_df.rename(columns={'acc': 'acc_median', 'confidence': 'confidence_median'}, inplace=True)

# Ordenar o DataFrame resultante pela chave
median_df.sort_values(by='key', inplace=True)


# In[46]:


median_df


# In[48]:


# Defina o caminho do arquivo Excel
output_excel_path = r"C:\Users\B28166\Desktop\EXCEL RELAORIOS DE AVALIAÇÃO\AVALIACAO1.xlsx"

# Exportar o DataFrame para um arquivo Excel
median_df.to_excel(output_excel_path, index=False, engine='openpyxl')

print(f"DataFrame exportado para {output_excel_path}")


# In[ ]:




