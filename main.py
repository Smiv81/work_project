from navec import Navec
import numpy as np
import pymorphy3
import re
import pandas as pd
from scipy import spatial
morf = pymorphy3.MorphAnalyzer()

navec = Navec.load('navec_hudlit_v1_12B_500K_300d_100q.tar')


"""Открывает строки из XLSX убирает символы и приводит все к нижнему регистру"""

def open_without_sub_lower(string):
    text_lst = np.array(pd.read_excel(string).dropna())
    return  [re.sub('\W+',' ', text[1].lower()) for text in (text_lst)]



"""Создает токены"""

def morf(data : list) -> list:
    morf = pymorphy3.MorphAnalyzer()
    total_lst = []
    for lst in data:
        lst_good = []
        for word in lst.split():
            if len(word) > 2:
                lst_good.append(morf.parse(word)[0].normal_form)
        total_lst.append(lst_good)
    return total_lst

"""Переводит текст в вектор"""

def text_to_vector(word : str) -> list:
    global vector
    f = []
    try:
        vector = navec[word]
    except KeyError:
        pass
    return vector


"""Считает расстояние между векторами"""

def vector_analiz(data_pok, data_rez, count_pok, count_rez, value):

    coin = 0
    total_cos = 0
    for a in data_pok:
        lst_total = []
        for b in data_rez:
            cos = 1 - spatial.distance.cosine(a, b)
            lst_total.append(cos)
        total_cos += max(lst_total)
        coin += 1
    if total_cos / coin > float(value):
            return ['Рез' , count_pok, 'Показате', count_rez, round(total_cos / coin, 3)]


def jaccard(list1, list2):
    total_K = 0
    count = 0
    for i in range(len(list1)):
        jackar_max = []
        count += 1
        for j in range(len(list2)):
            A = list1[i]
            B = list2[j]
            C = len(set(A).intersection(set(B)))
            D = len(set(A).union(set(B)))
            jackar =  C / D
            jackar_max.append(jackar)
        total_jakar = max(jackar_max)
        total_K += total_jakar
    return [total_K / count]

"""Подготовка данных показателей (переводит в список [id, значение])"""

def pok_id_val(data):
    morf = pymorphy3.MorphAnalyzer()
    pok = [' '.join(re.sub('\W+', ' ', text.lower()) for text in data.split() if len(text) > 2)]
    lst_pok = [i.split() for i in pok][0]
    pok_id = [lst_pok[0]]
    pok_value = lst_pok[1:]
    pok_value_normal = [morf.parse(w)[0].normal_form for w in pok_value]
    return [pok_id, pok_value_normal]

"""Подготовка данных результатов (переводит в список [id, значение])"""

def rez_vec_id_val(data):
    morf = pymorphy3.MorphAnalyzer()
    list_rez_id = []
    list_rez_val = []
    list_rez_vec = []
    for i in range(len(data)):
        if i % 2 == 0:
            rez = [' '.join(re.sub('\W+', ' ', text.lower()) for text in data[i].split())]
            cur_id = [rez[0].split()[0]]
            cur_val = rez[0].split()[1:]
            cur_val_norm = [morf.parse(w)[0].normal_form for w in cur_val]
            list_rez_id.append(cur_id)
            list_rez_val.extend([cur_val_norm])
        else:
            cur_vec = [data[i]]
            list_rez_vec.extend([cur_vec])

    return [list_rez_id, list_rez_val, list_rez_vec]

def open_file(res):
    if isinstance(res, str):
        data_file_1 = pd.read_excel(res[0])




