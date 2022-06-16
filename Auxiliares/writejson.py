import json

from numpy import size
import Functions_Bot as FB
import random
from array import *
from sklearn.feature_extraction.text import CountVectorizer
import json
import pandas as pd

# function to add to JSON

def Proc_NPL(detye, conjuntoDatos):
    roles = conjuntoDatos
    entrada = [detye]
    rolesE = CountVectorizer(binary=True, ngram_range=(1, 1), analyzer='word')
    vector_rolesE = rolesE.fit_transform(roles)
    a = rolesE.get_feature_names()
    entradaE = CountVectorizer(
        binary=True, ngram_range=(1, 1), analyzer='word')
    vector_entradaE = entradaE.fit_transform(entrada)
    b = entradaE.get_feature_names()
    aa = list(a)
    bb = list(b)
    resultante = set(aa).intersection(set(bb))
    resultanteE = CountVectorizer(
        binary=True, ngram_range=(1, 1), analyzer='word')
    vector_normalizado = resultanteE.fit_transform(resultante)
    trak = (" ".join(map(str, resultante)))
    print("resultante:  ",resultante)
    
    lista = list(resultante)
    """print(type(lista))
    for i in range (size(lista)):
        print(i,"   :  ",lista[i])"""
    print (trak)
    return trak, lista



def write_json(new_data, filename='Data.json'):
    with open(filename,'w') as file:
        json.dump(new_data,file,indent=4)
        
    # python object to be appended

def JsonAddMas(alte):
    
    """with open('Data.json') as f:
        datos = json.load(f)"""

    df = pd.read_json('Data.json')
    tam = size(df)
    d = []
    for i in range(tam):
        a=df["roles"][i]["cargo"]
        d.append(a)
    print(d)
    rta=Proc_NPL(detye=alte,conjuntoDatos=d)
    print(rta)

    with open("Data.json") as json_file:
        dre = json.load(json_file)
        temp = dre["roles"]
        y ={"inp":alte,
        "cargo":rta[0]
        } 
        temp.append(y)

    write_json(dre)
    return rta[1]

kta=JsonAddMas("arquitecto de software")
print(kta)