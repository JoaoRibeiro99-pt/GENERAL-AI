#!/usr/bin/env python
# coding: utf-8

# #INSTALAÇÕES,BIBLITECAS E CONSTANTES
# 
# 

# In[27]:


get_ipython().system('pip install unidecode')
get_ipython().system('pip install azure-ai-formrecognizer')
get_ipython().system('pip install pymupdf unidecode')
get_ipython().system('pip install openpyxl')

from unidecode import unidecode
# Azure AI
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential

from azure.ai.formrecognizer import (
    DocumentModelAdministrationClient,
    ModelBuildMode,
)
from azure.core.credentials import AzureKeyCredential

# search regex
import re

# read pdfs
import fitz

# list files in directory
from os import listdir
from os.path import isfile, join
import shutil
from pathlib import Path

import json
import numpy as np
import pandas as pd
import time
from datetime import datetime
from typing import Tuple, Dict


import os
import pandas as pd
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential

ENDPOINT = "https://di-ne-docintel001.cognitiveservices.azure.com/"
KEY = "6e0fb08c07ae4557b40d9aab940edd73"
MODEL_ID = "prebuilt-document"
input_path = r"C:\Users\B28166\Desktop\Relatorios_avaliacao"
output_excel_path = r"C:\Users\B28166\Desktop\EXCEL RELAORIOS DE AVALIAÇÃO\CAMPOS NECESSARIOS4.xlsx"


# #ITERAR FICHEIRO POR PASTA

# In[12]:


# Iterar sobre cada item na pasta
for item in os.listdir(input_path):
    item_path = os.path.join(input_path, item)
    
    # Verificar se o item é um ficheiro
    if os.path.isfile(item_path):
        with open(item_path, 'rb') as file:
            content = file.read()
        if content:
            print(f"File {item} read successfully.")
        else:
            print(f"Failed to read file {item}.")


# PROCESSAMENTO DO DOCUMENTO 

# In[13]:


# Crie o cliente de análise de documentos
document_analysis_client = DocumentAnalysisClient(endpoint=ENDPOINT, credential=AzureKeyCredential(KEY))

# Inicialize uma lista para armazenar os dados
data = []

# Iterar sobre cada item na pasta
for item in os.listdir(input_path):
    item_path = os.path.join(input_path, item)
    
    # Verificar se o item é um ficheiro
    if os.path.isfile(item_path):
        # Leia o conteúdo do arquivo
        with open(item_path, "rb") as file:
            file_content = file.read()

        # Inicie a análise do documento
        poller = document_analysis_client.begin_analyze_document(MODEL_ID, file_content)
        result = poller.result()

        # Adicione os pares chave-valor ao DataFrame
        for kv_pair in result.key_value_pairs:
            key = kv_pair.key.content if kv_pair.key else None
            value = kv_pair.value.content if kv_pair.value else None
            confidence = kv_pair.confidence if kv_pair.confidence else None
            data.append({'document': item, 'key': key, 'value': value, 'confidence': confidence})


# In[14]:


# Crie um DataFrame a partir da lista de dados
df = pd.DataFrame(data)
# Converta todos os valores do DataFrame para maiúsculas
df = df.applymap(lambda x: x.upper() if isinstance(x, str) else x)


# In[15]:


# Converta todos os valores do DataFrame para maiúsculas
df = df.applymap(lambda x: x.upper() if isinstance(x, str) else x)


# EXCEL COM OS CAMPOS TODOS
# 

# In[16]:


# Defina o caminho do arquivo Excel
output_excel_path = r"C:\Users\B28166\Desktop\EXCEL RELAORIOS DE AVALIAÇÃO\TODOS OS CAMPOS.xlsx"

# Exportar o DataFrame para um arquivo Excel
df.to_excel(output_excel_path, index=False, engine='openpyxl')

print(f"DataFrame exportado para {output_excel_path}")


# DEFINIÇÃO DOS VALORES DESEJADOS E DT COM OS CAMPOS DESEJADOS
# 

# In[17]:


valores_desejados = [
    'REFª:',
    'Nº DE PROCESSO	',
    'DATA ALTERACAO ESTADO',
    'NIF',
    'COORDENADAS GPS:',
    'RUA',
    'ARTIGO',
    'CÓDIGO POSTAL',
    'LOCALIDADE',
    'DISTRITO',
    'FREGUESIA',
    'CONCELHO',
    'ANO DE CONSTRUÇÃO',
    'AREA HABITACAO',
    'ACIMA DO SOLO',
    'Nº DE PISOS:',
    'PERSIANAS',
    'CAIXILHARIA',
    'ESTRUTURA',
    'REVESTIMENTO EXTERIOR',
    'COBERTURA',
    'DESCRIÇÂO DA HABITAÇÃO',
    'TIPOLOGIA',
    'RATING',
    'FRAÇÃO',
    'ESTADO CONSERVAÇÃO EDIFÍCIO',
    'ESTADO CONSERVAÇÃO CANALIZAÇÃO',
    'ESTADO CONSERVAÇÃO ELÉTRICA',
    'TIPO EDIFICIO',
    'ESTADO DE CONSERVAÇÃO',
    'PAVIMENTOS ZONAS SECAS',
    'TETOS',
    'EQUIPAMENTOS',
    'NÍVEL DE ACABAMENTOS',
    'PAREDES ZONAS HÚMIDAS',
    'PAREDES ZONAS SECAS',
    'PAVIMENTOS ZONAS HÚMIDAS',
    'ÁREA DO PAVIMENTO',
    'VALOR COMERCIAL',
    'Nº APÓLICE:',
    'ANO DE CONSTRUÇÃO'
    'VALOR ATUAL',
    ]

# Filtrar o DataFrame para manter apenas as linhas onde o valor da coluna 'KEY' está na lista de valores desejados
df_filtrado = df[df['key'].isin(valores_desejados)]

# Se você tem uma coluna que identifica documentos únicos, substitua 'documento' por esse nome
# Caso contrário, ajuste conforme a sua estrutura
coluna_documento = 'document'  # Substitua 'documento' pelo nome da sua coluna

if coluna_documento in df_filtrado.columns:
    # Remover duplicatas mantendo a primeira ocorrência dentro de cada grupo de documento e KEY
    df_filtrado_unico = df_filtrado.drop_duplicates(subset=[coluna_documento, 'key'])
else:
    # Se não houver coluna 'documento', apenas remova duplicatas baseadas na coluna 'KEY'
    df_filtrado_unico = df_filtrado.drop_duplicates(subset=['key'])


# In[18]:


df_filtrado_unico

