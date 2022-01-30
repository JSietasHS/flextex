# -*- coding: utf-8 -*-
"""
Created on Wed Sep 29 15:49:51 2021

@author: SIETAS
"""

import streamlit as st
import pandas as pd
import numpy as np
#import altair as alt
import plotly.express as px
import os

import tkinter as tk
from tkinter import filedialog
from tkinter.filedialog import askopenfilename

from Content import Text, Diagram, Table, Multirow, Multicol
from FlexTexDocument import FlexTexDocument
from Flextexdataset import Flextexdataset, Csv, MariaDB

import streamlit as st

from streamlit_quill import st_quill

from enums import Types

from pylatex import Section, Subsection, Command, Package, \
    MultiRow, Tabular, MultiColumn, NewLine
from pylatex.utils import NoEscape
from pylatex.section import Chapter

from pylatex.section import Chapter





#from altair_saver import save



# Save tuple in Datase (File Name, Dataset)
#datasets = []

def resetTableSessionstates():
    st.session_state.fontweightcoloumnlist = []
    st.session_state.fontweightrowlist = []
    st.session_state.coloumncolorlist = []
    st.session_state.rowcolorlist = []
    st.session_state.alignlist = []
    st.session_state.multirow =[]
    st.session_state.multicol = []
    


def main():
    st.title("FlExTeX")
    
    menu = ["Document","Text", "Diagram", "Table", "Dataset"]
    boollistfalse = [False, True]
    boollisttrue = [True, False]

    choice = st.sidebar.selectbox("Menu", menu)
    
    #Save Datasets
#    if "datasets" not in st.session_state:
#        st.session_state.datasets = [];
    
    #Save Displayed Content
    #if "content" not in st.session_state:
    #    st.session_state.content = [];
    
    if  choice == 'Text':
        st.subheader("Text")
        
        resetTableSessionstates()
        
        #st.success("Full Layout")
        
        if "document" not in st.session_state:
            st.warning('Es muss ein Dokument angelegt werden, bevor ein Text erstellt werden kann')
        else:
            
            #Two Columns
            col1,col2 = st.columns(2)
            
            # Spawn a new Quill editor
            text = st_quill()
            
            # Display editor's content as you type
            
            #Create a button
            appendtext = col1.button("Append Text")
            
            #Create Dataset = st.button("Append Text")   
    
            if appendtext :        
                st.write(text)
                textcontent = Text(text)
                #st.session_state.content.append(textcontent)
                st.session_state.document.appendContent(textcontent)
        
            
            
            
    elif choice == "Dataset":
        
        resetTableSessionstates()
        
        if "document" not in st.session_state:
            st.warning('Es muss ein Dokument angelegt werden, bevor ein Datensatz erstellt werden kann')
        else:
            
            st.subheader("Dataset anlegen")
            
            datasetconnectionlist = ["CSV","MariaDB"]
            datasetconnection = st.selectbox("Wähle eine Verbindungsart", datasetconnectionlist)
            
            if datasetconnection == "CSV":
            
                seperator = [";",",","."," "]
                sep = st.selectbox("Seperator", seperator)
                #data_file = st.file_uploader("Upload CSV", type = ["csv"])
                #if data_file is not None:
                #    file_details = { "filename" : data_file.name,
                #                    "filetype":data_file.type, 
                #                    "filesize":data_file.size}
                    
                #    print(data_file)
                #    st.write(file_details)
                
                csvpicker = st.button('Wähle eine CSV Datei')
                
                # Set up tkinter
                root = tk.Tk()
                root.withdraw()
                
                # Make folder picker dialog appear on top of other windows
                root.wm_attributes('-topmost', 1)
                
                data_file = None
                
                if csvpicker:
                    data_file = askopenfilename(master=root,filetypes = (("CSV Files","*.csv"),))
                    st.write(r'Ausgewählter ordner:', data_file)
                    st.session_state.data_file = data_file
                    
                        
                # Append to Datasets
                loadcsv = st.button('Datensatz aus CSV Laden')
                if (loadcsv):
                    if "data_file" in st.session_state :
                        dataset = Csv(st.session_state.data_file, sep)
                        
                        st.dataframe(dataset.getDataset())        
                        
                        #st.session_state.datasets.append((st.session_state.data_file ,dataset.getDataset()))
                        st.session_state.document.appendDataset(dataset)
                        st.session_state.data_file = None
                    
            
            elif datasetconnection == "MariaDB":
                
                user = st.text_input('User')
                password = st.text_input('Passwort')
                hostname = st.text_input('hostname')
                #port = st.text_input('port')
                database = st.text_input('database')
                query = st.text_input('query')
                datasetname = st.text_input('Datensatzname')
                
                createconnection = st.button('Verbindung herstellen')
                
                if createconnection:
                    dataset = MariaDB(hostname, database, user, password, query, datasetname)
                    st.session_state.document.appendDataset(dataset)
                    
                    #st.session_state.datasets.append(('mdbconn',dataset.getDataset()))
                    
                    st.dataframe(dataset.getDataset())
    
    
            if ( ("document" in st.session_state) and ( st.session_state.document.getDatasetnames() != [])):
                
                st.subheader("Dataset updaten")
                
                datasets = st.selectbox("dataset", st.session_state.document.getDatasetnames())
                
                updatedataset = st.button('Ausgewähltes Dataset Updaten')
                updatedatasets = st.button('Alle Datasets updaten')
                
                if updatedataset:
                    st.session_state.document.updateDatasetByIndex(st.session_state.document.getDatasetnames().index(datasets))
                    
                if updatedatasets:
                      st.session_state.document.updateAllDatasets()
                
                st.dataframe(st.session_state.document.getDatasetByIndex(st.session_state.document.getDatasetnames().index(datasets)))
                
                
                
        
    elif choice == "Diagram":
        st.subheader("Diagram")
        
        resetTableSessionstates()
        
 #       datasetsnames = []
        diagramlist = ["bar graph", "line chart"]
#        textselectboxcontent = []
#        types = [Types.O.name, Types.Q.name, Types.N.name]
        
        if "document" not in st.session_state:
            st.warning('Es muss ein Dokument angelegt und ein Datensatz eingefügt werden, bevor ein Diagramm erstellt werden kann')
        else:
            if st.session_state.document.getDatasetnames() != []:
                datasets = st.selectbox("dataset", st.session_state.document.getDatasetnames())
                fd = st.session_state.document.getFlextexdatasetByIndex(st.session_state.document.getDatasetnames().index(datasets))
                
                dt = st.selectbox("Diagramtyp", diagramlist)
                
                da= fd.getDataset()
                x = st.selectbox("x", da.columns)
                y = st.selectbox("y", da.columns)
                if dt == "bar graph":
        
#                    xtype = st.selectbox("xtype", types)
#                    ytype = st.selectbox("ytype", types)
#                    xlabel = st.radio("xlabel", [True, False])
#                    ylabel = st.radio("ylabel", [True, False])
                    title = st.text_input("title")
#                    xtitle = st.text_input("xtitle")
#                    ytitle = st.text_input("ytitle")
                    xaxisangle = st.number_input("Winkel der X Achsen Beschiftung", -360, 360, format = '%f')
                    xscalemin =  st.number_input("xscalemin", format = '%f', value = 0)
                    xscalemax = st.number_input("xscalemax", format = '%f', value = 0)
                    yscalemin = st.number_input("yscalemin", format = '%f', value = 0)
                    yscalemax = st.number_input("yscalemax", format = '%f', value = 0)
                    text = st.selectbox("text", da.columns)
                    #text = st.radio("text", [True, False])
#                    align = st.radio("align", ["center"])
#                    baseline = st.radio("baseline", ["bottom"])
#                    dx = 0
#                    dy = -10
                    
         #           encodetext
                    
                    width = st.number_input("width", format = '%f', value = 0)
                    height= st.number_input("height", format = '%f', value = 0)
                    colorred = st.number_input("R", 0, 255, format = '%f', value = 79)
                    colorgreen = st.number_input("G", 0, 255, format = '%f', value = 129)
                    colorblue = st.number_input("B", 0, 255, format = '%f', value = 189)
                    coloropacity = st.number_input("opacity",0.00, 1.00, format = '%.2f', value = 1.00)
                    
                    diagramcaption = st.text_input("LatexCaption")
                    diagramlabel = st.text_input("LatexLabel")
#                    size = 20
#                    textlabel = y # orx
#                    caption = "nocaption"
                

            
                range_x = None  
                range_y = None
                if ((xscalemin != 0) and (xscalemax != 0)):
                    range_x = [xscalemin, xscalemax]
                else:
                    range_x = None  
                    
                if ((yscalemin != 0) and (yscalemax != 0)):
                    range_y = [yscalemin, yscalemax]
                else:
                    range_y = None 
                    
                if width == 0:
                    width = None
                if height == 0:
                    height = None
                
#                labels=None
                
#                if ((xtitle == "") and (ytitle != "")):
#                    labels = dict(y=ytitle)   
#                elif ((xtitle != "") and (ytitle == "")):
#                    labels = dict(x=xtitle)   
#                elif((xtitle != "") and (ytitle != "")):
#                    labels = dict(y=ytitle, x = xtitle)   


 
                
                
                    
                colorstring = "rgba("+str(colorred)+","+str(colorgreen)+","+str(colorblue)+","+str(coloropacity)+")"
                
                diagramcontent = Diagram(fd,x=x,y=y,title=title,color_discrete_sequence=[colorstring], text=text, opacity = coloropacity, range_x = range_x , range_y = range_y , width = width, height = height, caption = diagramcaption, label = diagramlabel)
                
                

                fig = px.bar(da, x=x , y=y, title=title, color_discrete_sequence=[colorstring],text = text, opacity = coloropacity, range_x = range_x , range_y = range_y , width = width, height = height)
                    
                        
                fig.update_layout(xaxis_tickangle= xaxisangle)
                fig.update_traces(texttemplate='%{text:.2s}', textposition='outside')
                st.plotly_chart(fig, use_container_width=True)
            
                appenddiagram = st.button("Append Diagram")
                
                if appenddiagram:
                    #st.session_state.content.append(diagramcontent)
                    st.session_state.document.appendContent(diagramcontent)
        
            else: st.warning('Es muss ein Datensatz hinzugefügt werden, bevor ein Diagram erstellt werden kann')
                    
    elif choice == "Table":
        st.subheader("Table")
        
        #datasetsnames = []
        if "document" not in st.session_state:
            st.warning('Es muss ein Dokument angelegt und ein datensatz eingefügt werden, bevor eine Tabelle erstellt werden kann')
        else:
                        
            
            if st.session_state.document.getDatasetnames() != []:
                
                colorlist= ["black", "blue", "brown", "cyan", "darkgray", "gray", "green", "lightgray", "lime", "magenta", "olive", "orange", "pink", "purple", "red", "teal", "violet", "white", "yellow"]
                
                latexfontsizelist = ["normalsize", "Huge", "huge", "LARGE", "Large", "large", "small", "footnotesize", "scriptsize", "tiny"]
                
                aligns = ["l", "c", "r"]
                
                datasets = st.selectbox("dataset", st.session_state.document.getDatasetnames())
                        
                addcoloumnnames = st.selectbox("Füge Die erste Spalte als Überschrift hinzu", boollisttrue)    
                
                fontsize = st.selectbox("Schriftgröße", latexfontsizelist)  
                framecolor = st.selectbox("Rahmenfarbe", colorlist) 
                framewidth = st.number_input("framewidth", 0, 10, format = '%f', value = 2)
                rowstretch = st.number_input("rowstretch", 0.0, 5.0, format = '%.1f', value = 1.5)
                
                # fette Zeilen
                df = st.session_state.document.getDatasetByIndex(st.session_state.document.getDatasetnames().index(datasets)) 
               
                allrows = []
                rownumbers = []
                rownumbers2 = []
                
                rcelist = []
                 
                 #get all coloumns in sesseionstate alignlist              
                if st.session_state.rowcolorlist is not []:
                    for rce in st.session_state.rowcolorlist:
                                rcelist.append(rce[0])
                
                for i in range (0,df.shape[0]):
                    if st.session_state.fontweightrowlist is not []:
                        if not(i in st.session_state.fontweightrowlist):
                           rownumbers.append(i)
                    else:
                        rownumbers.append(i)
                        
                    if rcelist is not []:
                        if not(i in rcelist):
                            rownumbers2.append(i)
                    else:
                        rownumbers2.append(i)
                    allrows.append(i)

                    
                allcoloumns = []
                coloumnnumbers = []
                coloumnumbers2 = []
                cclist = []
                alignlist2 = []
                alignlistcoloumns = []
                
                 
                #get all coloumns in sesseionstate rowcolorlist
                rcelist = []
                if st.session_state.alignlist is not []:
                    for a in st.session_state.alignlist:
                                alignlist2.append(a[0])
                
                 #get all coloumns in sesseionstate rowcolorlist
                if st.session_state.coloumncolorlist is not []:
                         for ccl in st.session_state.coloumncolorlist:
                                cclist.append(ccl[0])
                
                for j in range (0,df.shape[1]):
                    if st.session_state.fontweightcoloumnlist is not []:
                        if not(j in st.session_state.fontweightcoloumnlist):
                            coloumnnumbers.append(j)
                    else:
                        coloumnnumbers.append(j)
                    
                    if cclist is not []:
                        if not(j in cclist):
                            coloumnumbers2.append(j)
                    else:
                        coloumnumbers2.append(j)
                        
                    if alignlist2 is not []:
                        if not(j in alignlist2):
                            alignlistcoloumns.append(j)
                    else:
                        alignlistcoloumns.append(j) 
                    allcoloumns.append(j)
                                     
                
                boldrow =  st.selectbox("Bold Row", rownumbers) 
                appendboldrow = st.button("Append Row to Bold Rows")
                if appendboldrow:
                    st.session_state.fontweightrowlist.append(boldrow)

                #fette Spalten
                boldcoloumn =  st.selectbox("Bold coloumn", coloumnnumbers)
                appendboldcoloumn = st.button("Append Coloumn to Bold Coloumn")
                if appendboldcoloumn:
                    st.session_state.fontweightcoloumnlist.append(boldcoloumn)

                
                #farbige Spalten
                colorforcoloumncolorlist = st.selectbox("Spaltenfarbe", colorlist) 
                coloumnforcoloumncolorlis = st.selectbox("Spalten die eingefärbt werden sollen", coloumnumbers2)
                appendcoloumncolor = st.button("Färbe die ausgewählte Spalte ein")
                if appendcoloumncolor:
                    st.session_state.coloumncolorlist.append((coloumnforcoloumncolorlis,colorforcoloumncolorlist))
                
                #farbige Zeilen
                colorforrowcolorlist = st.selectbox("Zeilenfarbe", colorlist) 
                coloumnforrowcolorlist = st.selectbox("Zeilen die eingefärbt werden soll", rownumbers2) 
                appendrowcolor = st.button("Färbe die ausgewählte Zeile ein")
                if appendrowcolor:
                    st.session_state.rowcolorlist.append((coloumnforrowcolorlist,colorforrowcolorlist))
                
                coloumnalign = st.selectbox("Textausrichtung", aligns) 
                coloumnforalign = st.selectbox( "Spalte, die eine Textausrichtung erhalten soll", alignlistcoloumns)
                appenalign = st.button("Füge eine Textausrichtung hinzu")
                if appenalign:
                    st.session_state.alignlist.append((coloumnforalign,coloumnalign))
                
                #Zellen innerhalb 
                startrow = st.selectbox("Wähle eine Zeile, in der Zellen verbunden werden soollen", allrows)
                startcoloumn = st.selectbox("Wähle die erste Spalte, in der die Zellen verbunden werden sollen", allcoloumns)
                endcoloumn = st.selectbox("Wähle die letzte Spalte, in der die Zellen verbunden werden sollen", allcoloumns)
                appendmulticol = st.button("Füge Multicoloumn ein")
                
                if appendmulticol:
                    if startcoloumn > endcoloumn:
                        st.error("Die erste Spalte muss kleiner sein als die letzte")
                    else:
                        st.session_state.multicol.append(Multicol(startrow, startcoloumn, endcoloumn))
                #Zellen innerhab einer einer Zeile verbindenSpalte verbinden
                startcoloumn = st.selectbox("Wähle eine Spalte, in der Zellen verbunden werden soollen", allcoloumns)
                startrow =  st.selectbox("Wähle die erste Zele, in der die Zellen verbunden werden sollen", allrows)
                endrow =  st.selectbox("Wähle die letzte Zeile, in der Zellen verbunden werden soollen", allrows)
                appendmultirow = st.button("Füge Multirow ein")
                if appendmultirow:
                    if startrow > endrow:
                        st.error("Die erste Reihe muss kleiner sein als die letzte")
                    else:
                        st.session_state.multirow.append(Multirow(startcoloumn, startrow, endrow))
                
                latexlabel = st.text_input("LatexLabel")

            
                        
                        
                appendtable = st.button("append table to Tex document")
                
                    
                if appendtable:
                    
                    if (st.session_state.multicol == []):
                         st.session_state.multicol == None
                         
                    if (st.session_state.multirow == []):
                         st.session_state.multirow == None
                         
                    if (st.session_state.coloumncolorlist == []):
                        st.session_state.coloumncolorlist == None
                        
                    if (st.session_state.rowcolorlist == []):
                        st.session_state.rowcolorlist == None 
                        
                    if (st.session_state.alignlist == []):
                        st.session_state.alignlist == None 
                    
                    da = st.session_state.document.getFlextexdatasetByIndex(st.session_state.document.getDatasetnames().index(datasets))
               
                    tablecontent = Table(da, addcoloumnnames, multcolumn= st.session_state.multicol, multrow =  st.session_state.multirow, columncolor = st.session_state.coloumncolorlist,  rowcolor = st.session_state.rowcolorlist, fontsize = fontsize,  framecolor = framecolor, framewidth = framewidth, rowstretch = rowstretch, align = st.session_state.alignlist, fontweightcoloumnlist = st.session_state.fontweightcoloumnlist, fontweightrowlist = st.session_state.fontweightrowlist, label = latexlabel)
                    #st.session_state.content.append(tablecontent)
                    st.session_state.document.appendContent(tablecontent)
               
                     
            else: st.warning('Es muss ein Datensatz hinzugefügt werden, bevor eine Tabelle erstellt werden kann')
                
    elif choice == "Document":
        st.subheader("Document")
        
        resetTableSessionstates()
        
        fontsizelist = ["10","10.5", "11", "11.5", "12"]
        documentclasslist = ["article", "report", "book" ,"letter", "scrartcl", "scrreprt", "scrbook", "scrlttr2"]
        papersizeformatlist = ["a4paper ","letterpaper", "a5paper", "b5paper", "executivepaper", "legalpaper"]
        singleordoublesidelist = ["twoside ","oneside"]

    
        
        if "document" not in st.session_state:
            
            
            # Set up tkinter
            root = tk.Tk()
            root.withdraw()
            
            # Make folder picker dialog appear on top of other windows
            root.wm_attributes('-topmost', 1)
            
            # Folder picker button
            st.write('Bitte wähle einen Speicherordner aus:')
            folderpicker = st.button('Ordner wählen')
            if folderpicker:
                dirnametmp = st.text_input(r'Ausgewählter ordner:', filedialog.askdirectory(master=root))
                if "dirname" not in st.session_state:
                    st.session_state.dirname = dirnametmp;
            filename = st.text_input ('Bitte wähle einen Dateinamen')
            
            st.write('Erweiterte Einstellungen')
        
            fontsize = st.selectbox("Schriftgröße", fontsizelist)
            documentclass = st.selectbox("Dokumentenklasse", documentclasslist)
            papersizeformat = st.selectbox("Papierformat", papersizeformatlist)
            singleordoubleside = st.selectbox("Seitentyp", singleordoublesidelist)
            lmodern = st.selectbox("lmoder (Schriftart)", boollistfalse)
            textcomp = st.selectbox("textcomp", boollistfalse)
            pagenumbers = st.selectbox("page_numbers", boollistfalse)
            listofentrypraefix = st.selectbox("listofentrypräfix", boollisttrue)
                
                    
            createdoc = st.button('Lege Dokument an')
            
            if createdoc: 
                    st.session_state.document = FlexTexDocument(st.session_state.dirname, filename, documentclass, fontsize, papersizeformat, singleordoubleside, listofentrypraefix ,lmodern, textcomp, pagenumbers)
                    st.success("Das Dokument wurde angelegt. Sie können jetzt Text und Datensätze einfügen.")

        else:
        
            createtex = st.button("Erzeuge Dokument")
        
            if createtex:
                st.session_state.document.generateFlexTex()
                #doc = Document(r'C:\Temp\newdoc', documentclass="scrreprt",document_options="a4paper, 10.5pt, twoside, listof=entryprefix", lmodern=False, textcomp=False,page_numbers=False )
                #for c in st.session_state.content :
                #    doc = c.toTex(doc)
                #doc.generate_tex() 
                
            st.subheader("Bearbeite das Dokument")
                
            folderpicker = st.button('Neuen Ordner wählen')
            
            # Set up tkinter
            root = tk.Tk()
            root.withdraw()
            
            # Make folder picker dialog appear on top of other windows
            root.wm_attributes('-topmost', 1)
            
            if folderpicker:
                dirnametmp = st.text_input(r'Ausgewählter ordner:', filedialog.askdirectory(master=root))
                st.session_state.dirname = dirnametmp;
            else:
                dirnametmp = st.write(r'Ausgewählter ordner:' + st.session_state.document.getDirname())

            
            filename = st.text_input('Dateiname:',st.session_state.document.getFilename())
            
            st.write('Erweiterte Einstellungen')
        
            fontsize = st.selectbox("Schriftgröße", fontsizelist, index=fontsizelist.index(st.session_state.document.getFontsize()) )
            documentclass = st.selectbox("Dokumentenklasse", documentclasslist, index=documentclasslist.index(st.session_state.document.getDocumentclass()) )
            papersizeformat = st.selectbox("Papierformat", papersizeformatlist, index=papersizeformatlist.index(st.session_state.document.getPapersizeformat()) )
            singleordoubleside = st.selectbox("Seitentyp", singleordoublesidelist, index=singleordoublesidelist.index(st.session_state.document.getSingleordoubleside()) )
            lmodern = st.selectbox("lmoder (Schriftart)", boollistfalse, index=boollistfalse.index(st.session_state.document.getLmodern()) )
            textcomp = st.selectbox("textcomp", boollistfalse, index=boollistfalse.index(st.session_state.document.getTextcomp()) )
            pagenumbers = st.selectbox("page_numbers", boollistfalse, index=boollistfalse.index(st.session_state.document.getPagenumbers()) )
            listofentrypraefix = st.selectbox("listofentrypräfix", boollisttrue, index=boollisttrue.index(st.session_state.document.getListofentrypraefix()) )
                
            changedoc = st.button('Dokumenteneinstellungen ändern')
            
            if changedoc:
                st.session_state.document.setDirname( st.session_state.dirname )
                st.session_state.document.setFilename( filename )
                st.session_state.document.setFontsize( fontsize)
                st.session_state.document.setDocumentclass( documentclass)
                st.session_state.document.setPapersizeformat( papersizeformat)
                st.session_state.document.setSingleordoubleside( singleordoubleside)
                st.session_state.document.setLmodern( lmodern)
                st.session_state.document.setTextcomp( textcomp)
                st.session_state.document.setPagenumbers( pagenumbers)
                st.session_state.document.setListofentrypraefix( listofentrypraefix)
                
                
            
            
if __name__ == '__main__':
    main()
    



