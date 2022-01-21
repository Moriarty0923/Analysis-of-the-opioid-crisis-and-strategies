import xlrd
import pandas as pd
import numpy as np
from math import radians, cos, sin, asin, sqrt
import xlwt
import operator
import copy
from functools import reduce


def init(table):
    nrows = table.nrows
    ncols = table.ncols
    # use FIPS_combined as key to record info
    county = {}
    print(table.row_values(0)[5])
    print(table.row_values(0)[10])
    print(table.row_values(0)[11])
    for i in range(1, nrows):
        FIPS_Combined = table.row_values(i)[5]
        if FIPS_Combined not in county.keys():
            lat = float(table.row_values(i)[10])
            lon = float(table.row_values(i)[11])
            county[FIPS_Combined] = (lon, lat)
    return county

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

def renew_cluster(centers, pos, cities):
    label = dict().fromkeys([i for i in range(len(centers))], [])
    for city in cities:
        dist = []
        for center in centers:
            #print(center)
            dist.append(haversine(pos[city], center))
        index = dist.index(min(dist))
        cluster = copy.deepcopy(label[index])
        cluster.append(city)

        label[index] = cluster
        new_centers = []

    for i in range(len(centers)):
        positions = list(map(lambda x: pos[x], label[i]))
        #print(positions)
        new_centers.append(np.mean(positions, axis=0))
        #print(new_centers)
        #print(label)
    return new_centers, label

# convert positions into k clusters
def k_means(k, pos):
    cities = list(pos.keys())
    if k > len(cities):
        print("K too large")
        return False
    else:
        # initial centers
        old_centers = [list(pos[cities[i]]) for i in range(k)]
        # itartion until convergence
        while True:
            new_centers, cluster = renew_cluster(old_centers, pos, cities)
            # print(new_centers)

            if (np.array(new_centers) == np.array(old_centers)).all():
                break
            else:
                old_centers = new_centers
    zone = {}
    for key in cluster:
        for city in cluster[key]:
            zone[city] = key
    return new_centers, cluster, zone

def output(centers, cluster, zone, data):
    # new_file = xlwt.Workbook(encoding = ’utf-8’)
    # new_table = file.add_sheet(’data’)
    table = []
    writen = []
    DrugReportZone=[[{} for i in range(100)] for i in range(0,8)]
    TotalDrugReportZone=[[0 for i in range(100)]for i in range(0,8)]
    for i in range(1,data.nrows):
        YYYY = int(data.row_values(i)[0])
        FIPS_Combined = data.row_values(i)[5]
        substanceName = data.row_values(i)[6]
        DrugReport = data.row_values(i)[7]
        TotalDrugReport = data.row_values(i)[8]
        label = zone[FIPS_Combined]
        DrugReportZone[YYYY-2010][label][substanceName] = DrugReportZone[YYYY-2010][label].get(substanceName,0)+DrugReport
        if (YYYY,FIPS_Combined) not in writen:
            TotalDrugReportZone[YYYY-2010][label] += TotalDrugReport
            writen.append((YYYY,FIPS_Combined))
        entry = [YYYY, label, substanceName, DrugReport, TotalDrugReport]
        table.append(entry)
    for entry in table:
        entry[3] = DrugReportZone[entry[0]-2010][entry[1]][entry[2]]
        entry[4] = TotalDrugReportZone[entry[0]-2010][entry[1]]
    #print(table)
    return table
# def f(x, y):
# if (np.array(x[:3]) == np.array(y[:3])).all():
# return x[:3] + [x[3]+y[3]] + [x[4]+y[4]]
# reduce(lambda x, y: f , table)



if __name__ == '__main__':
    data = xlrd.open_workbook("Data_GEO.xls").sheets()[0]
    county = init(data)
    centers, cluster, zone = k_means(100, county)
    table = output(centers, cluster, zone, data)
    #write data into excel
    excel = xlwt.Workbook()
    sheet1 = excel.add_sheet("Date based on Zone")
    sheet2 = excel.add_sheet("ZoneCounty Correspondence Table")
    sheet3 = excel.add_sheet("County Position")
    sheet1.write(0, 0, "YYYY")
    sheet1.write(0, 1, "Zone")
    sheet1.write(0, 2, "substanceName")
    sheet1.write(0, 3, "DrugReportZone")
    sheet1.write(0, 4, "TotalDrugReportZone")
    for row, entry in enumerate(table):
        for col in range(len(entry)):
            sheet1.write(row+1, col, entry[col])
    sheet2.write(0, 0, "Zone")
    sheet2.write(0, 1, "Latitude")
    sheet2.write(0, 2, "Longitude")

    for row, key in enumerate(list(cluster.keys())):
        sheet2.write(row+1, 0, row)
        sheet2.write(row+1, 1, centers[row][1])
        sheet2.write(row+1, 2, centers[row][0])
        for col in range(len(cluster[key])):
            sheet2.write(row+1, col+3, cluster[key][col])
    sheet3.write(0, 0, "County")
    sheet3.write(0, 1, "FIS_Combined")
    sheet3.write(0, 2, "Latitude")
    sheet3.write(0, 3, "Longitude")
    written = []
    row_num = 0
    for row in range(1, data.nrows):
        entry = data.row_values(row)
        #print(entry)
        if entry[2] not in written:
            row_num += 1
            written.append(entry[2])
            info = [entry[2], entry[3]+'+'+entry[4], entry[10], entry[11]]
            for col in range(4):
                sheet3.write(row_num, col, info[col])
    excel.save("kmeans.xls")