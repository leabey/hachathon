#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Apr  2 14:23:42 2021

@author: lea


Trouver l'adresse des compteurs 
"""

import xlrd 
import os 

from xlwt import Workbook
classeur = Workbook() 
chemin_fichier_compteur ="/Users/lea/Desktop/les_compteurs"
chemin_fichier_ecriture = "/Users/lea/Desktop/classeV1.xls"


workbook_ecriture = xlrd.open_workbook(chemin_fichier_ecriture )
SheetNameList = workbook_ecriture.sheet_names() 
worksheet_ecriture = workbook_ecriture.sheet_by_name(SheetNameList[0])
num_rows_e = worksheet_ecriture.nrows 
num_cells_e = worksheet_ecriture.ncols 
                        
feuille = classeur.add_sheet("Feuil1",cell_overwrite_ok=True)

feuille.write( 0, 0, "Compteur")
feuille.write( 0, 1, "Classe")
feuille.write( 0, 2, "Ministère")
feuille.write( 0, 3, "Lieux")
feuille.write( 0, 4, "Adresse")
feuille.write( 0, 5, "Code Postal")
feuille.write( 0, 6, "Ville")
for i in range (1,num_rows_e): 
    
    valeur1=0
    compteur = str(worksheet_ecriture.cell_value(i, 0))
    feuille.write(i, 0, str(compteur))
    feuille.write(i, 1, int(worksheet_ecriture.cell_value(i, 1)))
    
   
    for fichier in os.listdir(chemin_fichier_compteur ):
        if fichier != ".DS_Store":
            
            
            workbook_l = xlrd.open_workbook(chemin_fichier_compteur +"/" + fichier)
            SheetNameList_l = workbook_l.sheet_names() 
            worksheet_l = workbook_l.sheet_by_name(SheetNameList_l[0])
            num_rows_l = worksheet_l.nrows 
            num_cells_l = worksheet_l.ncols 
            for j in range (8,num_rows_l): 
                if str(worksheet_l.cell_value(j, 13))== compteur : 
                    feuille.write(i, 2, worksheet_l.cell_value(j, 2))
                    feuille.write(i, 3, worksheet_l.cell_value(j, 8))
                    feuille.write(i, 4, worksheet_l.cell_value(j, 9))
                    feuille.write(i, 5, worksheet_l.cell_value(j, 10))
                    feuille.write(i, 6, worksheet_l.cell_value(j, 11))
                    valeur1=1
                
                if valeur1 ==1 : 
                    break
            
            if valeur1==1 : 
                break 
                    
                    

                
        
classeur.save("/Users/lea/Desktop/compteur_classe_adresse.xls")                