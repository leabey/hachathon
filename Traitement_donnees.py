#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Mar 30 14:17:44 2021

@author: lea


Ce programme permet de traiter les données 
Nous avons décidé de créer pour chaque dossier, un excel
1 feuille = 1 compteur 
1 ligne = 1 date 
1 colonne = 10 minutes 
"""
import os 

import xlrd
 
from xlwt import Workbook, Formula
 
chemin ="/Users/lea/Desktop/Hackhaton/données_conso_importants"
for element in os.listdir(chemin):
   classeur = Workbook()
   if os.path.exists(chemin + "/"+element):
        if os.path.isdir(chemin + "/"+element):
            
            for fichier in os.listdir(chemin + "/"+element):
                if fichier != ".DS_Store" : 
                    
  #si fichier xlsx
                    if fichier == "P10 CULTURE.xlsx":
                        
                        workbook = xlrd.open_workbook("/Users/lea/Desktop/Hackhaton/données_conso_importants/"+element + "/" + fichier)
                        SheetNameList = workbook.sheet_names() 
                        worksheet = workbook.sheet_by_name(SheetNameList[0])
                        num_rows = worksheet.nrows 
                        num_cells = worksheet.ncols 
                        
                        for k in range (1, num_cells): 
                            nom = worksheet.cell_value(0,k)
                            #le nom a une taille max 
                            if len(nom)> 30:
                                nom = nom[:30]
                            feuille = classeur.add_sheet(nom,cell_overwrite_ok=True)
                            
                            #création de l'entete de l'excel
                            feuille.write(0,0,"date")
                            colonne = 1
                            for i in range(0, 24,1) : 
                                for j in range(0, 60, 10): 
                                     heure = str(i) + ":"+ str(j)
                                     feuille.write(0,colonne, heure )
                                     colonne +=1
                            ligne = 0
                            date_precedente = "" 
                            
                            #premier colonne  1 ligne = 1 date 
                            for m in range(1, num_rows): 
                                date = worksheet.cell_value(m, 0)
                                y, mm, d, h, i, s = xlrd.xldate_as_tuple(date,0)
                                d = str(d)
                                mm =str(mm)
                                if len(d) == 1: 
                                    d= "0" + d
                                if len(mm)==1:
                                    mm="0"+mm
                                    
                                dd = d+ "/" + mm+"/" + str(y)

                                if date != date_precedente :
                                    ligne +=1
                                    date_precedente = date
                                    feuille.write(ligne, 0, dd)
                                    colonne = 1
                                    
                                try : 
                                    var = int(worksheet.cell_value(m, k)) 
                                    feuille.write(ligne, colonne, var)
                                except (RuntimeError, TypeError, NameError, ValueError):
                                     feuille.write(ligne, colonne, worksheet.cell_value(m, k))
                                
                                colonne +=1
                                
                                
                    elif fichier == "30000112124579.xlsx":
                        workbook = xlrd.open_workbook("/Users/lea/Desktop/Hackhaton/données_conso_importants/"+element + "/" + fichier)
                        SheetNameList = workbook.sheet_names() 
                        worksheet = workbook.sheet_by_name(SheetNameList[0])
                        num_rows = worksheet.nrows 
                        num_cells = worksheet.ncols 
                        nom_compteur = ""
                        
                        feuille = classeur.add_sheet(worksheet.cell_value(1,0),cell_overwrite_ok=True)
                        feuille.write(0,0,"date")
                        colonne = 1
                        for i in range(0, 24,1) : 
                            for j in range(0, 60, 10): 
                                 heure = str(i) + ":"+ str(j)
                                 feuille.write(0,colonne, heure )
                                 colonne +=1
                        ligne = 0
                           
                        date_precedente = "" 
                        
                        for m in range(1, num_rows): 
                            date = worksheet.cell_value(m, 1)
                            y, mm, d, h, i, s = xlrd.xldate_as_tuple(date,0)
                            d = str(d)
                            mm =str(mm)
                            if len(d) == 1: 
                                d= "0" + d
                            if len(mm)==1:
                                mm="0"+mm
                                
                            dd = d+ "/" + mm+"/" + str(y)
                            
                            if date != date_precedente :
                                ligne +=1
                                date_precedente = date
                                feuille.write(ligne, 0, dd)
                                colonne = 1
                                if m == 1 : 
                                    colonne =7
                            for k in range(3, num_cells):  
                                try : 
                                    var = int(worksheet.cell_value(m, k)) 
                                    feuille.write(ligne, colonne, var)
                                except (RuntimeError, TypeError, NameError, ValueError):
                                     feuille.write(ligne, colonne, worksheet.cell_value(m, k))
                                
                                colonne +=1
                        #pour les autres types de fichier    
                    else : 

                        ff= fichier
                        if len(ff) > 30:
                            ff = ff[:30]
                        
                        feuille = classeur.add_sheet(ff[:-3],cell_overwrite_ok=True)
                        feuille.write(0,0,"date")
                        colonne = 1
                        for i in range(0, 24,1) : 
                            for j in range(0, 60, 10): 
                                 heure = str(i) + ":"+ str(j)
                                 feuille.write(0,colonne, heure )
                                 colonne +=1
                        ligne = 0
                        date_precedente = ""    
        
                        if fichier != ".DS_Store" : 
                            if fichier.endswith(".txt") or fichier.endswith(".TXT"):
                                nn="/Users/lea/Desktop/Hackhaton/données_conso_importants/"+ element + "/"+fichier
                              
                                f = open(nn,"r")
                                t = f.readlines() 
                                for k in range(len(t)): 
                                    date = t[k][0:10]
                                    if date != date_precedente :
                                        ligne +=1
                                        date_precedente = date
                                        feuille.write(ligne, 0, date)
                                        colonne = 1
                                    var = ""
                                    
                                    debut = 10 
                                    
                                    while (t[k][debut])!=":" :
                                        debut+=1
                                    
                                    if  t[k][debut+3]=='\t':
                                        cas = 2
                                        for p in range(debut+3, len(t[k])):
                                        
                                            if t[k][p] != '\t': 
                                                var +=  t[k][p]
                                                if p == (len(t[k])-1): 
                                                    try : 
                                                        var = int(var)
                                                        feuille.write(ligne, colonne, var)
                                                    except (RuntimeError, TypeError, NameError, ValueError):
                                                        feuille.write(ligne, colonne, var)
                                                    var = ""
                                                    colonne +=1
                                            elif (t[k][p]=='\t' and var != ""):
                                                try : 
                                                    var = int(var)
                                                    feuille.write(ligne, colonne, var)
                                                except (RuntimeError, TypeError, NameError, ValueError):
                                                    feuille.write(ligne, colonne, var)
                                                var = ""
                                                colonne +=1
                                    else :
                                        
                                        for p in range(debut+3, len(t[k])):
                        
                                            if t[k][p]!=" " or t[k][p]!=' ' : 
                                                var +=  t[k][p]
                                                if p == (len(t[k])-1):
                                                    try : 
                                                        var = int(var)
                                                        feuille.write(ligne, colonne, var)
                                                    except (RuntimeError, TypeError, NameError, ValueError):
                                                        feuille.write(ligne, colonne, var)
                                                    
                                                    var = ""
                                                    colonne +=1
                                            elif (t[k][p]==" " or t[k][p]=="\t" )and var != "":
                                                try : 
                                                    var = int(var)
                                                    feuille.write(ligne, colonne, var)
                                                except (RuntimeError, TypeError, NameError, ValueError):
                                                    feuille.write(ligne, colonne, var)
                                                
                                                var = ""
                                                colonne +=1
                                               

                            elif fichier.endswith(".csv"): 
                                f = open("/Users/lea/Desktop/Hackhaton/données_conso_importants/"+ element + "/"+fichier,"r")
                                t = f.readlines() 
                                for k in range(len(t)): 
                                    date = t[k][:10]
                                    if date != date_precedente :
                                        ligne +=1
                                        date_precedente = date
                                        feuille.write(ligne, 0, date)
                                        colonne = 1
                                    var = ""
                                    for p in range(16, len(t[k])):
                                        if t[k][p]!=";": 
                                            var +=  t[k][p]
                                            if p == (len(t[k])-1): 
                                                try : 
                                                    var = int(var)
                                                    feuille.write(ligne, colonne, var)
                                                except (RuntimeError, TypeError, NameError, ValueError):
                                                    feuille.write(ligne, colonne, var)
                                                var = ""
                                                colonne +=1
                                        elif t[k][p]==";"and var != "":
                                            try : 
                                                var = int(var)
                                                feuille.write(ligne, colonne, var)
                                            except (RuntimeError, TypeError, NameError,ValueError):
                                                feuille.write(ligne, colonne, var)
                                            var = ""
                                            colonne +=1
                                     

                         
            classeur.save("/Users/lea/Desktop/Hackhaton/excel/"+element + ".xls")      
            #sauvegarde 

    
    
