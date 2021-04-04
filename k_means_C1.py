#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Apr  4 20:14:40 2021

@author: lea
"""

"""
Created on Sat Apr  3 20:23:42 2021

@author: lea

Porgramme pour déterminer le nombre de classe et pour  réaliser le clustering 
Ecriture des résultatsq dans un excel 
"""

import numpy as np 
import pandas as pd 
import matplotlib.pyplot as plt 
from scipy.cluster.hierarchy import dendrogram, linkage, fcluster
from sklearn.cluster import KMeans
from sklearn import metrics, cluster
from xlwt import Workbook
import xlrd

# ----- PREPARATION DE L'EXCEL FINAL
classeur1 = Workbook() 
feuille1 = classeur1.add_sheet("Feuil1",cell_overwrite_ok=True)

classeur = Workbook() 
feuille = classeur.add_sheet("Feuil1",cell_overwrite_ok=True)

feuille.write( 0, 0, "Compteur")
feuille.write( 0, 1, "Données brutes Classe K-means 4 ")
feuille.write( 0, 2, "Données brutes Classe CAH 4")

feuille.write( 0, 3, "Données brutes Classe K-means 5 ")
feuille.write( 0, 4, "Données brutes Classe CAH 5")

feuille.write( 0, 5, "Données écart Classe K-means 4 ")
feuille.write( 0, 6, "Données écart Classe CAH 4")

feuille.write( 0, 7, "Données écart Classe K-means 5 ")
feuille.write( 0, 8, "Données écart Classe CAH 5")

workbook_ecriture = xlrd.open_workbook("/Users/lea/Desktop/dossier sans titre/1_recap.xlsm")
SheetNameList = workbook_ecriture.sheet_names() 
worksheet_ecriture = workbook_ecriture.sheet_by_name(SheetNameList[0])
num_rows_e = worksheet_ecriture.nrows 
num_cells_e = worksheet_ecriture.ncols 
                        

for i in range (1,num_rows_e): 
    
    compteur = str(worksheet_ecriture.cell_value(i, 0))
    feuille.write(i, 0, str(compteur))

# ----- FICHIER 1 :DONNEES DE LA MOYENNE ANNUELLE DE CONSOMMATION AU PAS DE 10 MIN

data1 = pd.read_excel ("/Users/lea/Desktop/dossier sans titre/1_recap.xlsm", index_col = 0)
X1 = data1.iloc[:,1:-1].values

#METHODE DU COUDE 
# wcss1=[]
# for i in range(2,10):
#     kmeans1 =KMeans(n_clusters=i, init ="k-means++", random_state=0)
#     kmeans1.fit(X1)
#     wcss1.append(kmeans1.inertia_)
# plt.plot(range(2,10),wcss1)
# plt.title("methode du coude pour trouver k données ")
# plt.xlabel("nombre de clusters")
# plt.ylabel("WCSS")
# plt.show()


# #Attribution des classes 
k_means1A = KMeans(n_clusters=6, init = 'k-means++', random_state=0)
y_kmeans1A=k_means1A.fit_predict(X1)


for i in range(len(y_kmeans1A)): 
    feuille.write(i+1, 1, int(y_kmeans1A[i]))
    


    
#   # #utilisation de la métrique "silhouette" #faire varier le nombre de clusters de 2 à 10 
# res1 = np.arange(10,dtype="double")
# for k in np.arange(10):
#     km1 = cluster.KMeans(n_clusters=k+2)
#     km1.fit(data1) 
#     res1[k] = metrics.silhouette_score(data1,km1.labels_)
# print(res1)

# #graphique
# plt.title("Silhouette ") 
# plt.xlabel("Nombre de clusters") 
# plt.plot(np.arange(2,12,1),res1) 
# plt.show()

# # # #DENDROGRAM
Z1 = linkage(X1, 'ward')
# fig = plt.figure(figsize=(25, 10))
# dn = dendrogram(Z1)
# plt.title("Dendrogramme classe 1")

# # # # #Classement 
groupes_cah1 = fcluster(Z1,6,criterion='maxclust') 
# #print(groupes_cah1)
# ##index triés des groupes
# idg1 = np.argsort(groupes_cah1)
# ##affichage des observations et leurs groupes
# #print(pd.DataFrame(data1.index[idg1],groupes_cah1[idg1]))

for i in range(len(groupes_cah1)): 
    feuille.write(i+1, 2, int(groupes_cah1[i]))


# groupes_cah1 = fcluster(Z1,5,criterion='maxclust') 
# for i in range(len(groupes_cah1)): 
#     feuille.write(i+1, 4, int(groupes_cah1[i]))
    
#



#---------------FICHIER 2 : ECART RELATIF PAS DE 10 MIN PAR RAPPORT A LA MOYENNE ANNUELLE TOTALE
    
data2 = pd.read_excel ("/Users/lea/Desktop/dossier sans titre/1_ecart.xlsm", index_col = 0)
X2 = data2.iloc[:,1:-1].values

# #METHODE DU COUDE 
wcss2=[]
for i in range(2,10):
    kmeans2 =KMeans(n_clusters=i, init ="k-means++", random_state=0)
    kmeans2.fit(X2)
    wcss2.append(kmeans2.inertia_)
    
plt.plot(range(2,10),wcss2)
plt.title("methode du coude pour trouver k ")
plt.xlabel("nombre de clusters")
plt.ylabel("WCSS")
plt.show()

# #utilisation de la métrique "silhouette" #faire varier le nombre de clusters de 2 à 10 
res2 = np.arange(10,dtype="double")
for k in np.arange(10):
    km2 = cluster.KMeans(n_clusters=k+2)
    km2.fit(data2) 
    res2[k] = metrics.silhouette_score(data2,km2.labels_)

#graphique
plt.title("Silhouette") 
plt.xlabel("Nombre de clusters") 
plt.plot(np.arange(2,12,1),res2) 
plt.show()


# #Attribution des classes 
k_means2A = KMeans(n_clusters=3, init = 'k-means++', random_state=0)
y_kmeans2A=k_means2A.fit_predict(X2)
for i in range(len(y_kmeans2A)): 
    feuille.write(i+1, 5, int(y_kmeans2A[i]))
    

# #DENDROGRAM
Z2 = linkage(X2, 'ward')
# fig = plt.figure(figsize=(25, 10))
# dn = dendrogram(Z2)
# plt.title("Dendrogramme classe 1 ")



# #Classement 
groupes_cah2 = fcluster(Z2,3,criterion='maxclust') 
#print(groupes_cah2)
#index triés des groupes
idg2 = np.argsort(groupes_cah2)
#affichage des observations et leurs groupes
#print(pd.DataFrame(data2.index[idg2],groupes_cah2[idg2]))
for i in range(len(groupes_cah2)): 
    feuille.write(i+1, 6, int(groupes_cah2[i]))
    



    
classeur.save("/Users/lea/Desktop/classement_comparaison1AA.xls")  