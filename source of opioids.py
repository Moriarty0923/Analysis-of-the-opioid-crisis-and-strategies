import xlrd
import pandas as pd
import numpy as np
from math import *
import xlwt
import operator
import copy
from functools import reduce
import matplotlib.pyplot as plt
import matplotlib
import networkx as nx
from itertools import combinations
from igraph import *
import seaborn as sns
import pylab
import random


def init(data):
    sheet1 = dasheet1 = data.sheets()[0] # Date based on Zone
    sheet2 = data.sheets()[1] # ZoneCounty Correspondence Table
    sheet3 = data.sheets()[2] # County Position
    # map substanceName to Drug number
    drug2num = {}
    num2drug = []

    rank = 0
    dataslice = [[], [], [], [], [], [], [], []]
    for row in range(1, sheet1.nrows):
        entry = sheet1.row_values(row)
        YYYY = entry[0]
        substanceName = entry[2]
        if substanceName not in drug2num.keys():
            drug2num[substanceName] = rank
            num2drug.append(substanceName)
            rank += 1

        # seperate this data into 8 parts by year
        if YYYY == 2010:
            dataslice[0].append(entry)
        if YYYY == 2011:
            dataslice[1].append(entry)
        if YYYY == 2012:
            dataslice[2].append(entry)
        if YYYY == 2013:
            dataslice[3].append(entry)
        if YYYY == 2014:
            dataslice[4].append(entry)
        if YYYY == 2015:
            dataslice[5].append(entry)
        if YYYY == 2016:
            dataslice[6].append(entry)
        if YYYY == 2017:
            dataslice[7].append(entry)

    # dict to store each zone’s position
    center = {}
    for row in range(1, sheet2.nrows):
        entry = sheet2.row_values(row)
        zone = int(entry[0])
        if zone not in center.keys():
            lat = entry[1]
            lon = entry[2]
        center[zone] = (lon, lat)

    return drug2num, num2drug, dataslice, center


def make_distance_xls(center):
    excel = xlwt.Workbook()
    sheet1 = excel.add_sheet("distance")
    for i in range(0,100):
        for j in range(0,100):
            sheet1.write(i,j,haversine(center[i],center[j]))
    excel.save("distance.xls")


def make_simi_xls(data,center,alpha):
    excel = xlwt.Workbook()
    sheet1 = excel.add_sheet("sim")
    for i in range(0,100):
        dict={}
        ad = cal_simi(i,data,center,alpha)
        for it in ad:
            dict[it[0]]=it[1]
        for j in range(0,100):
            if i==j:
                sheet1.write(i, j, 1)
            else:
                sheet1.write(i,j,dict.get(j,0))
    excel.save("similarity.xls")



def cal_simi(zone, data, center, alpha):
    adjacentZones = []
    for i in range(100):
        if i != zone and haversine(center[i], center[zone]) < alpha:
            Sum = np.sum(np.sum(np.array(data[i]))) + np.sum(np.sum(np.array(data[zone])))
            corr = 0
            for j in range(69):
                tendency_i = data[i][j]
                tendency_zone = data[zone][j]
                sum_i = np.sum(np.array(tendency_i))
                sum_zone = np.sum(np.array(tendency_zone))
                if sum_zone == 0 or sum_i == 0:
                    continue
                por = sum_i + sum_zone
                coef = np.corrcoef(tendency_i, tendency_zone)[0, 1]
                if isnan(coef) or coef == 0:
                    pass
                else:
                    corr += coef * 1.0 * por / Sum
            adjacentZones.append([i, corr])
    return adjacentZones

def sparseMatrix(data, drug2num, num2drug):
    """
    Matrix has 100 zones, each forms a row
    each row represents the num of drug report in that zone for each drugs
    initially, we set it all to zeros
    """
    # len(num2drug) = 69
    sparseMatrix = np.zeros([100, 69])
    for row in range(len(data)):
        zone = int(data[row][1])
        drugName = data[row][2]
        drugNum = drug2num[drugName]
        drugReport = data[row][3]
        sparseMatrix[zone][drugNum] += drugReport
    return sparseMatrix


def haversine(pos1, pos2):
    """
    Calculate the great circle distance between two points
    on the earth (specified in decimal degrees)
    """
    # convert decimal into arc system
    lon1, lat1 = pos1
    lon2, lat2 = pos2
    lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])
    # haversine formula
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * asin(sqrt(a))
    r = 6371 # radius of earth, unit as kilometer
    return c * r * 1000


def similarity(zone, data, center, alpha):
    adjacentZones = cal_simi(zone,data,center,alpha)
    if len(adjacentZones) > 10:
        adjacentZones.sort(key=lambda x: x[1])
        return list(map(lambda x: x[0], adjacentZones[-10:]))
    else:
        return list(map(lambda x: x[0], adjacentZones))


def adjMatrix(medical_name, time, distance_threshold):
    G=nx.DiGraph()
    H=nx.DiGraph()
    M=nx.DiGraph()
    excel_path = 'kmeans.xls'
    klocation = pd.read_excel(excel_path, sheet_name='ZoneCounty Correspondence Table')
    print(klocation)
    kdata = pd.read_excel(excel_path, sheet_name='Date based on Zone')
    distance = pd.read_excel('distance.xls', sheet_name='distance',header=None).values
    simi = pd.read_excel('similarity.xls',sheet_name = 'sim',header=None).values

    edges = []
    medical_data = kdata[(kdata['substanceName'] == medical_name) & (kdata['YYYY'] == time)]
    combins = [list(c) for c in combinations(set(medical_data['Zone'].values.tolist()), 2)]
    #print(medical_data)
    #print(combins)
    weights = []

    for i in range(len(combins)):
        x = combins[i][0]
        y = combins[i][1]
        if distance[x, y] < distance_threshold:
            #print(medical_data[medical_data[’Zone’] == x])
            #print(medical_data[medical_data[’Zone’] == y])
            if (sum(medical_data[medical_data['Zone'] == x]['DrugReportZone'].values)
                    >=sum(medical_data[medical_data['Zone'] == y]['DrugReportZone'].values)):
                edges.append(combins[i])
                G.add_edges_from([combins[i]], weight = round(simi[combins[i][0],combins[i][1]],1))
                weights.append(round(simi[combins[i][0],combins[i][1]],1))
            else:
                combins[i].reverse()
                edges.append(combins[i])
                #print(combins[i].reverse())
                G.add_edges_from([combins[i]], weight = round(simi[combins[i][0], combins[i][1]],1))
                weights.append(round(simi[combins[i][0], combins[i][1]],1))
    print(edges)
    degree = [[x, 0] for x in range(100)]
    for i in range(len(edges)):
        degree[edges[i][0]][1] += 1
    degree.sort(key=lambda x: x[1], reverse=True)
    ra = random.randint(0, 3)
    origin = degree[ra][0]

    print(degree)
    # Draw the picture
    # G.add_edges_from(edges)
    ind = []
    posi = []
    for i in range(100):
        ind.append(i)
        posi.append((klocation['Longitude'].iloc[i], klocation['Latitude'].iloc[i]))
    position = dict(zip(ind, posi))
    vals = []
    for i in range(100):
        vals.append(int(sum(medical_data[medical_data['Zone'] == i]['DrugReportZone'].values)))
    values_dic = dict(zip(ind, vals))
    values = [values_dic.get(node, 1.0) for node in G.nodes()]
    for i in range(len(values)):
        values[i] = values[i] - 10
        if (values[i] > 210):
            values[i] = 210
    edge_labels = dict([((u, v,), d['weight'])for u, v, d in G.edges(data=True)])

    plt.figure()
    nx.draw_networkx_edge_labels(G,pos = position ,edge_labels=edge_labels)
    #nx.draw_networkx_edges(G, pos = position, width = 2, alpha = 0.5, arrows = True, arrowstyle=’->’,edge_color=weights,with_labels=True)
    nx.draw(G, pos = position, node_color = "skyblue", node_size = 2000,with_labels=True)
    fig = matplotlib.pyplot.gcf()
    fig.set_size_inches(40, 20)
    plt.savefig(str(medical_name) + '_' + str(time) + "_" + "0.png") # save as png

if __name__ == '__main__':
    data = xlrd.open_workbook('kmeans.xls')
    drug2num, num2drug, dataslice, center = init(data)
    #cal_distance(center)
    trainingMatrix = np.zeros((100,69,7),dtype = np.int)
    #print(len(num2drug))
    for year in range(7):
        tmp = sparseMatrix(dataslice[year], drug2num, num2drug)
        for i in range(100):
            for j in range(len(num2drug)):
                trainingMatrix[i][j][year] += tmp[i][j]

    held_outMatrix = sparseMatrix(dataslice[7], drug2num, num2drug)

    #print(trainingMatrix)
    make_simi_xls(trainingMatrix,center,300000)
    #print(similarity(2, trainingMatrix, center, 300000))
    adjMatrix('Heroin', 2014, 200000)
    # #get_all_origin(2010, 200000)