import pandas as pd
import numpy as np
from array import *
from sklearn.feature_extraction.text import CountVectorizer



def Proc_NPL(detye):
    roles = ['Analista administrativo', 'Agile Coach', 'Analista Financiero', 'App Dynamics', 'Applications Support', 'Arquitecto Software', 'Desarrollador Backend', 'Asistente Administrativo ', 'Business Analyst', 'Costumer Experience', 'Ingeniero de datos', 'DataBase Administrator', 'Coordinador de desarrollo', 'Desarrollador FullStack', 'IT Admin ', 'Tech leader', 'MDM', 'Project Manager', 'Quality assurance', 'Scrum Master', 'Product Owner', 'Ingeniero de soporte', 'Tester Analyst','Administrador web application','Desarrollador Frontend', 'Mobile developer', 'Cloud engineer', 'Desarollador Software', 'Desarollador Software IV']
    entrada = [detye]
    rolesE = CountVectorizer(binary=True, ngram_range=(1,1), analyzer='word')
    vector_rolesE = rolesE.fit_transform(roles)
    a = rolesE.get_feature_names()
    #Estandarización entrada
    entradaE = CountVectorizer(binary=True, ngram_range=(1,1), analyzer='word')
    vector_entradaE = entradaE.fit_transform(entrada)
    b = entradaE.get_feature_names()
    #print(b)
    #print(vector_entradaE.toarray())
    aa = list(a)
    bb = list(b)
    #print (aa)
    #print(bb)
    resultante = set(aa).intersection(set(bb))
    #print(resultante)
    resultanteE = CountVectorizer(binary=True, ngram_range=(1,1), analyzer='word')
    vector_normalizado = resultanteE.fit_transform(resultante)
    trak=(" ".join(map(str, resultante)))
    return  trak
    #print(vector_normalizado.toarray())     
    #print(a)
    #print(vector_rolesE.toarray())


gda=Proc_NPL(input())
print(gda)

