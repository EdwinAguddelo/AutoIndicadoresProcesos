from datetime import datetime
global DOD_FilePath,OC_FilePath
import pandas as pd
import numpy as np
import calendar
import os

def pathConstructor(path):
    global DOD_FilePath,OC_FilePath
    DOD_FilePath = path.DOD_FilePath
    OC_FilePath = path.OC_FilePath

def startproccess(path):    
    OC_DataSet = pd.read_excel(OC_FilePath)
    DOD_DataSet = pd.read_excel(DOD_FilePath)
    dataSetFinal = proccess(OC_DataSet,DOD_DataSet)
    mes = datetime.now().month
    anteriormes = calendar.month_name[mes - 1]
    fileAnteriorMes = 'DOD {}.xlsx'.format(anteriormes)
    rutaPath = os.path.join(path,fileAnteriorMes)
    ExportarExcel(dataSetFinal,rutaPath)

def proccess(OC_DataSet,DOD_DataSet):
    OC_DataSet = renombrarColumnaID(OC_DataSet)
    OC_DataSet['ID2']=pd.to_numeric(OC_DataSet['ID2'], errors='coerce')
    OC_DataSet['ID2'] = OC_DataSet['ID2'].fillna(0).astype(np.int64)
    DOD_Final = DOD_DataSet.merge(OC_DataSet, left_on='ID', right_on='ID2', how='inner')
    dataSetFinal = DOD_Final[['ID','Carta de certificación','carta sin pruebas','Carta de certificación condicionada','GIT',
          'Fecha de vencimiento carta de certificación','Aprueba el Paso a Producción','Dueño de producto que aprueba  paso a producción',
          'Orden de  Cambio','Codigo Cierre','Dirección']]
    return dataSetFinal

def renombrarColumnaID(OC_DataSet):
    OC_DataSet = OC_DataSet.rename(columns = {'Código Definición de Terminado':'ID2'})
    return OC_DataSet

def ExportarExcel(dataSetFinal,rutaPath):
    print(rutaPath)
    dataSetFinal.to_excel(rutaPath,index=False)
    print('exportado a {}'.format(rutaPath))
