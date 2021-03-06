from pylatex import Document, Section, Subsection, Command, Package, \
    MultiRow, Tabular, MultiColumn, NewLine
from pylatex.utils import NoEscape
from pylatex.section import Chapter
import  altair  as  alt 
import numpy as np
import pandas as pd
import os

#from altair_saver import save

from selenium import webdriver


#row - The row in which the multi col start 
#start coloumn - the start coloumn to connect the row
#end coloumn - the end coloumn to connect the row
class Multicol():
    """A class that saves Values to define a Multicoloumn."""
    
    row = 0
    startcolumn = 0
    endcolumn = 0
    data = None
    noescapedata = True

    def __init__(self, row, startcolumn, endcolumn, noescapedata = True):
        self.row = row
        self.startcolumn = startcolumn
        self.endcolumn = endcolumn
        self.noescapedata = noescapedata
        
    def get_row(self):
        return self.row
    
    def get_startcolumn(self):
        return self.startcolumn
    
    def get_endcolumn(self):
        return self.endcolumn
    
    def get_length(self):
        return self.endcolumn +1 - self.startcolumn
    
    def get_data(self):
        return self.data
    
    def set_noescapedatafalse(self):
        self.noescapedata = False
    
    def set_data(self, data):
        if self.noescapedata:
            self.data = MultiColumn(self.get_length(), align='c|', data=NoEscape(data))
        else:
            self.data = MultiColumn(self.get_length(), align='c|', data=data)
        
        
    def checkMulticolInDF(self, row,column):
        return ((row == self.get_row()) and (column == self.get_startcolumn()))
    

#col - The coloumn in which the multi row start 
#startrow - the start coloumn to connect the coloumns
#endrow - the end coloumn to connect the coloumns
class Multirow():
    """A class that saves Values to define a Multicoloumn."""
    
    column = 0
    startrow = 0
    endrow = 0
    data = None

    def __init__(self, column, startrow, endrow):
        self.column = column
        self.startrow = startrow
        self.endrow = endrow
        
    def get_col(self):
        return self.column
    
    def get_startrow(self):
        return self.startrow
    
    def get_endrow(self):
        return self.endrow
    
    def get_data(self):
        return self.data
    
    def set_data(self, data):
        self.data = MultiRow(NoEscape(int(self.get_length()*(-1))), data=NoEscape(data))

    def get_length(self):
        return self.endrow +1 - self.startrow
 
    def checkMultirowInDF(self, row,column):
        return ((column == self.get_col()) and (row == self.get_startrow()))
    
def getCommandstringFromList(valuelist, checkvalue,latexcommand):
    res = ""
    if valuelist is not None:
           for value in valuelist:
               if value[0] == checkvalue and value[1] is not None :
                   res = '\\'+latexcommand+" {"+ str(value[1]) +"}"
    return res

#Helper Class to Print Horizontal Line in different colors    
class Hhlinevalue():
    
    color = "white"
    size = 0
    bordersamecolor = False
    
    def __init__(self, size, color, bordersamecolor = False):
        self.size = size
        self.color = color   
        self.bordersamecolor = bordersamecolor
    
    def set_color(self, color):
        self.color = color
    
    def set_size(self, size):
        self.size = size
        
    def set_bordersamecolor(self, bordersamecolor):
        self.bordersamecolor = bordersamecolor
        
    def get_size(self):
        return self.size
    
    def get_color(self):
        return self.color
    
    def get_bordersamecolor(self):
        return self.bordersamecolor
        
    def set_size_minusone(self):
        self.size = self.size -1

def getColorfromCommandString(command):
    res = None
    if ("{" in command ) and ("}" in command):
        res = command[command.index('{')+1:command.index('}')]
    return res

#Create a Table out of a pandas Dataframe 
#input  document
#       dataframe: a pandas dataframe
#       coloumncolor: List (index, color)
#       fontsize: Latexfontsize
#       coloumwithcolor list of integer #Paint the number of the coloumn in the 
#                               coloumncolor default value all coloumns
#       multirows # List of Triple [Rownumber,start coloumn, end coloumn]
#                   rownumber start at 0 the row of the table
#                   coloumnnumber start at 0 the start coloumn to connect coloumns
#       multicoloumns # # List of MultiCol elements to add multicoloumns     
#       framecolor: color of te frame as string   
#       framewidth: width of the frame in pt
#       rowheight: 
#       align: tuple (index, [clr])
def createtable(doc, df, addcoloumnnames = True, multcolumn=None, multrow=None, columncolor = None, rowcolor=None, fontsize = None, framecolor = "black", framewidth = 2, rowstretch = None, align = None, fontweightcoloumnlist = [], fontweightrowlist = [], label = None):
    dfcopy = df.copy(deep=True) #copy of the dataframe
    #multcolumn = [Multicol(0,0,1)] #test
    #multrow = [Multirow(0,0,1)] #test
    #columncolor = [(1,"gray",0.8),(3,"gray",0.8)]
    #rowcolor = [(1,"gray"),(3,"gray")]
    rowcellcolor = ""
    columncellcolor = ""
    data = ""
    addmcol = None # add multicoloumn
    addmrow = None # add multirow
    values = [] #add values to table
    i=0 #counter
    h=0 #counterhhline
    c='|' #coloumns
    headercolor = ""
    setalign = False
    fontweightrow = False
    fontweightcolumn = False
    hhline = [] ### contains a tuple of size and color 
    hhlinelatexcommand = r"\hhline{|" #set hhline command 
    cellcolorstring = "white" #color
    
    all_columns = list(dfcopy) # Creates list of all column headers
    dfcopy[all_columns] = dfcopy[all_columns].astype(str) #Set all coloumnsto String
    
    doc.append(NoEscape("\\setlength{\\arrayrulewidth}{"+ str(framewidth) +"pt}"))
    #create table coloumns
    if framecolor is not None:
        doc.append(NoEscape(r'\arrayrulecolor{'+framecolor+'}'))
    if fontsize is not None:
        doc.append(NoEscape(r'\begin{'+fontsize+'}'))
    if rowstretch is not None:
        doc.append(NoEscape(r'\renewcommand{\arraystretch}{'+str(rowstretch)+'}'))
        
    #align
    while i in range(0,len(df.columns)): 
        if align is not None:
            setalign = False
            for a in align:
                if ((a[0] is not None) and (a[1] is not None) and (a[0] == i)):
                    c = c + str(a[1]) + '|'
                    setalign = True
        if not setalign:
           c = c + 'c|' 
        i = i+1 
               
    table3 = Tabular(c)
    table3.add_hline()
    #add header coloumn of pd.dataframe
    if addcoloumnnames:
        coloumnnames = df.columns.values.tolist()
        if (fontweightrowlist is not None) and (-1 in fontweightrowlist):
            for i in range(0,len(df.columns)):
                coloumnnames[i] = r"\textbf{"+coloumnnames[i]+"}"
        if fontweightcoloumnlist is not None:
             for i in range(0,len(df.columns)):
                 if i in fontweightcoloumnlist:
                     coloumnnames[i] = r"\textbf{"+coloumnnames[i]+"}"
        if rowcolor is not None:
            headercolor = ""
            headercolor = getCommandstringFromList(columncolor,0,"cellcolor")
            for i in range(0,len(df.columns)):
                coloumnnames[i] = NoEscape(headercolor + " " + coloumnnames[i])
        if columncolor is not None:
            for i in range(0,len(df.columns)):
                headercolor = ""
                headercolor = getCommandstringFromList(columncolor,i,"cellcolor")
                coloumnnames[i] = NoEscape(headercolor + " " + coloumnnames[i])
        table3.add_row(coloumnnames)
        table3.add_hline()
        
    #crete hhline
    for cid, c in enumerate(df.columns):
        cellcolor= ""
        columncellcolor = ""
        columncellcolor = getCommandstringFromList(columncolor,cid,"cellcolor")    
        #columncolor before rowcolor     
        if columncellcolor == "" :
            cellcolor = rowcellcolor
        else:
            cellcolor = columncellcolor
        if cellcolor != "":
            cellcolorstring = getColorfromCommandString(cellcolor)
        else:
            cellcolorstring = "white"          
        #create List for hhline
        hhline.append(Hhlinevalue(0,cellcolorstring))
            
    
    #iterate through rows
    for row in dfcopy.index:
        
        print(str(row))
        
        #add font weight
        fontweightrow= False
        if row in fontweightrowlist:
            fontweightrow = True
        
        rowcellcolor = ""
        #print('row:' + str(row))
        values = []
        #iterate through coloumns
        rowcellcolor = getCommandstringFromList(rowcolor,row,"cellcolor")
        for columnind, column in enumerate(df.columns):
            #add fontweight
            fontweightcolumn = False
            if i in fontweightcoloumnlist:
                fontweightcolumn = True
                
            if fontweightrow or fontweightcolumn:
                dfcopy.iat[row,columnind] = "\\textbf{" + str(dfcopy.iat[row,columnind]) + "}"
            
            addmcol = None
            addmrow = None
            cellcolor= ""
            columncellcolor = ""
            columncellcolor = getCommandstringFromList(columncolor,columnind,"cellcolor")
            
            #columncolor before rowcolor
           
            if columncellcolor == "" :
                cellcolor = rowcellcolor
            else:
                cellcolor = columncellcolor
                
              
            if cellcolor != "":
                cellcolorstring = getColorfromCommandString(cellcolor)
            else:
                cellcolorstring = "white"
                          
            
            #check if multicoloumn is in cell
            if multcolumn is not None:   
                for mcol in multcolumn:
                    data = ''
                    if mcol.checkMulticolInDF(row,columnind):
                        addmcol = mcol
            #check multirow is in cell
            if multrow is not None: 
                for mrow in multrow:
                    if mrow.checkMultirowInDF(row,columnind):
                        addmrow = mrow
            #append multicolrow
            if ((addmcol is not None) and (addmrow is not None)):
                for i in range(0,addmrow.get_length()):
                    for j in range(0,addmcol.get_length()):
                        data = data + str(dfcopy.iat[row+i,columnind+j])
                        #change hhline color and size
                        hhline[columnind+j].set_size(addmrow.get_length()-1-j)
                        hhline[columnind+j].set_color(getColorfromCommandString(cellcolor))
                        hhline[columnind+j].set_bordersamecolor(True)
                        if (i > 0)  and (i != addmrow.get_endrow()) :
                            dfcopy.iat[row+i,columnind+j] = "" 
                        else:
                            dfcopy.iat[row+i,columnind+j] = None 
                    #append more multicoloumns
                    if (i > 0) and (i < (addmrow.get_length()-1)):
                        multcolumn.append(Multicol((addmcol.get_row()+i),addmcol.get_startcolumn(), addmcol.get_endcolumn()))
                    #values.append(NoEscape(cellcolor))
                ####    
                addmcol.set_data(cellcolor)
                values.append(addmcol.get_data())
                ####
                addmcol.set_noescapedatafalse()
                addmrow.set_data(cellcolor+data)               
                addmcol.set_data(addmrow.get_data())       
                #values.append(NoEscape(cellcolor + addmcol.get_data()))
                dfcopy.iat[row+addmrow.get_length()-1,columnind] = addmcol.get_data()           
            #append multicol
            elif (addmcol is not None):
                for i in range(0,addmcol.get_length()):
                    data = data + str(dfcopy.iat[row,columnind+i])
                    dfcopy.iat[row,(columnind+i)] = None                 
                addmcol.set_data(cellcolor + data)
                values.append(addmcol.get_data())
            #append multirow
            elif (addmrow is not None):
                for i in range(0,addmrow.get_length()):
                    data = data + str(dfcopy.iat[row+i,columnind])
                    dfcopy.iat[row+i,columnind] = "" 
                addmrow.set_data(cellcolor + data) 
                values.append(NoEscape(cellcolor))
                dfcopy.iat[row+addmrow.get_length()-1,columnind] = addmrow.get_data()
                #change hhline color and size
                hhline[columnind].set_size(addmrow.get_length()-1)
                hhline[columnind].set_color(getColorfromCommandString(cellcolor))
                
                
            #append normal value
            elif ((addmcol is None) and (addmrow is None)):
                #Dont append None but append ''
                if ((dfcopy.iat[row,columnind] is not None) and ( "MultiRow" in str(dfcopy.iat[row,columnind]))):               
                     values.append(dfcopy.iat[row,columnind])
                elif ((dfcopy.iat[row,columnind] is not None) and (not("None" in dfcopy.iat[row,columnind]))) :
                    values.append(NoEscape(cellcolor + dfcopy.iat[row,columnind]))
        #print('Data:' + str(values))
        
        
        table3.add_row(values) 
        h=0
        for hhlinevalue in hhline:
            h=h+1
            if ((hhlinevalue.get_size()) > 0) and (h != len(hhline)):
                if (hhlinevalue.get_bordersamecolor()):
                    hhlinelatexcommand = hhlinelatexcommand + r">{\arrayrulecolor{" + hhlinevalue.get_color() + r"}}->{\arrayrulecolor{"+ hhlinevalue.get_color() +"}}|"
                else:
                    hhlinelatexcommand = hhlinelatexcommand + r">{\arrayrulecolor{" + hhlinevalue.get_color() + r"}}->{\arrayrulecolor{white}}|"
                hhlinevalue.set_size_minusone()
                if ((hhlinevalue.get_size()) == 0):
                    hhlinevalue.set_bordersamecolor(False)
                    hhlinevalue.set_color("white")
            else:
                hhlinelatexcommand = hhlinelatexcommand + r">{\arrayrulecolor{white}}-|"
        hhlinelatexcommand = hhlinelatexcommand + r"}"
        table3.append(NoEscape(hhlinelatexcommand))
        hhlinelatexcommand = r"\hhline{|"

        
    if not (label is None):
        table3.append(NoEscape("\label{tab:"+label+"}"))
    doc.append(table3)
    if fontsize is not None:
        doc.append(NoEscape(r'\end{'+fontsize+'}'))
      


def latexmainstart(doc,year):
    
    doc.preamble.append(Package("inputenc", "utf8"))
    doc.preamble.append(Package("babel", "ngerman"))
    doc.preamble.append(Package("fontenc", "T1"))
    doc.preamble.append(NoEscape(r"\renewcommand{\familydefault}{\sfdefault}"))
    doc.preamble.append(Package("helvet"))
    doc.preamble.append(Package("natbib"))
    doc.preamble.append(Package("textpos"))
    doc.preamble.append(Package("graphicx"))
    doc.preamble.append(Package("amsmath"))
    doc.preamble.append(Package("float"))
    doc.preamble.append(Package("etoolbox"))
    
    doc.preamble.append(Package("caption", "bf,justification=justified,singlelinecheck=false"))
    #doc.preamble.append(NoEscape(r"\counterwithout{figure}{chapter}"))
    #doc.preamble.append(NoEscape(r"\counterwithout{table}{chapter}"))
    #doc.preamble.append(Package("caption", "justification=justified,singlelinecheck=false"))
    #doc.preamble.append(NoEscape(r"\renewcommand*{\thefigure}{\textbf{\arabic{figure}}}"))
    #doc.preamble.append(NoEscape(r"\renewcommand*{\thetable}{\textbf{\arabic{table}}}"))
    
    doc.preamble.append(Package("adjustbox", "export"))
    doc.preamble.append(Package("xcolor"))
    doc.preamble.append(Package("fancyhdr"))
    doc.preamble.append(Package("marginnote"))
    
    doc.preamble.append(Package("geometry", "a4paper,outer=5cm,inner=1.5cm,top=2cm, bottom=3cm"))
    doc.preamble.append(Package("footnote"))
    
    doc.preamble.append(Package("multirow"))
    
    doc.preamble.append(NoEscape(r"\let\oldfbox\fbox"))
    doc.preamble.append(NoEscape(r"\renewcommand\fbox[1]{\savenotes\oldfbox{#1}\spewnotes}"))
    doc.preamble.append(Package("pdfpages"))
    
    doc.preamble.append(Package("parskip"))
    doc.preamble.append(Package("expl3"))
    doc.preamble.append(Package("afterpage"))
    doc.preamble.append(Package("wrapfig"))
    
    #Caption Paket
    #Customizing der Bild und Tabellencaptions (Einige Zusatzbefehle befinden sich als Parameter im usepackage Caption Bereich
    doc.preamble.append(NoEscape(r"\DeclareCaptionLabelFormat{meinLabel}{#1#2\hspace{1.5ex}}"))
    doc.preamble.append(NoEscape(r"\captionsetup[table]{labelformat=meinLabel,labelsep=space}"))
    doc.preamble.append(NoEscape(r"\captionsetup[figure]{labelformat=meinLabel,labelsep=space}"))
    
    doc.preamble.append(NoEscape(r"\renewcommand{\figurename}{\textbf{Abb.}}"))
    doc.preamble.append(NoEscape(r"\renewcommand{\tablename}{\textbf{Tab.}}"))
    
    #ListofFigures & ListofTables Zahleneintr??ge werden fettgedruckt

    doc.preamble.append(NoEscape(r"\ExplSyntaxOn"))
    doc.preamble.append(NoEscape(r"   \clist_map_inline:nn"))
    doc.preamble.append(NoEscape(r"     {figure,table}"))
    doc.preamble.append(NoEscape(r"     {\DeclareTOCStyleEntry["))
    doc.preamble.append(NoEscape(r"       pagenumberformat=\bfseries"))
    doc.preamble.append(NoEscape(r"       entryformat=\bfseries"))
    doc.preamble.append(NoEscape(r"       ]{section}{#1}"))
    doc.preamble.append(NoEscape(r"     }"))
    doc.preamble.append(NoEscape(r" \ExplSyntaxOff"))


 
 #Spacing vor neuen Chaptern entfernen
#\renewcommand*{\chapterheadstartvskip}{\vspace*{0cm}}

    doc.preamble.append(NoEscape(r"\RedeclareSectionCommand["))
    doc.preamble.append(NoEscape(r"  beforeskip=0pt,"))
    doc.preamble.append(NoEscape(r"  afterskip=\baselineskip,"))
    doc.preamble.append(NoEscape(r"  afterindent=false"))
    doc.preamble.append(NoEscape(r"  ]{chapter}"))
    doc.preamble.append(NoEscape(r"\counterwithout{figure}{chapter}"))
    doc.preamble.append(NoEscape(r"\counterwithout{table}{chapter}"))
    
    #Farben Colorbox
    doc.preamble.append(NoEscape(r"\definecolor{PaleAqua}{rgb}{0.74, 0.83, 0.9}")) 
    doc.preamble.append(NoEscape(r"\definecolor{LightBlue}{RGB}{217,226,238}")) 
    doc.preamble.append(NoEscape(r"\definecolor{Solitude}{HTML}{DCE6F1}")) 
    doc.preamble.append(NoEscape(r"\definecolor{TropicalBlue}{HTML}{B8CCE4}"))
    #Fancyhdr-Package genutzt zur Nutzung muss Pagestyle auf "Fancy" ge??ndert werden die vorangestellten Buchstaben gelten der Platzierung, f??r eine individuelle Kopf- und Fu??zeile im Dokument 
    doc.preamble.append(NoEscape(r"\pagestyle{fancy}")) 
    doc.preamble.append(NoEscape(r"\addtolength{\headwidth}{\marginparsep}")) 
    doc.preamble.append(NoEscape(r"\addtolength{\headwidth}{\marginparwidth}")) 
    doc.preamble.append(NoEscape(r"\fancyhead[RE,LO]{\textbf{Sozialatlas "+str(year)+"}}"))
    doc.preamble.append(NoEscape(r"\fancyhead[LE,RO]{\includegraphics[scale=0.5]{ABBILDUNGEN/Logo_Stadt_Flensburg.png}}"))
    doc.preamble.append(NoEscape(r"\fancyfoot[CE,CO]{}"))
    doc.preamble.append(NoEscape(r"\fancyfoot[LE,RO]{\thepage}"))
    
    doc.preamble.append(NoEscape(r"\renewcommand{\footrulewidth}{1pt}"))
    # Header u. Footer auch auf Chapter Pages anzeigen
    #??ndert das Pagestyle Argument auf fancyhdr in KOMA Klasse sccreprt (Etoolbox Package ben??tigt)
    doc.preamble.append(NoEscape(r"\renewcommand{\chapterpagestyle}{fancy}")) 
    #F??gt eine Leerzeile zwischen ??berschriften/Abs??tzen dazu
    doc.preamble.append(NoEscape(r"\setlength{\parindent}{0pt}")) 
    
    doc.preamble.append(NoEscape(r"\raggedbottom"))
    
    doc.preamble.append(Package(r"hhline"))
    doc.preamble.append(Package(r"colortbl"))
    
    
    #doc.append(NoEscape(r"\includepdf[pages={1}]{ABBILDUNGEN/Deckblatt Sozialatlas.pdf}"))

 

def latexmainend(doc):  
    doc.append(NoEscape(r"\listoffigures")) 
    doc.append(NoEscape(r"\addcontentsline{toc}{chapter}{\listfigurename}")) 
    doc.append(NoEscape(r"\thispagestyle{empty}")) 
    doc.append(NoEscape(r"\newpage")) 
    doc.append(NoEscape(r"\listoftables")) 
    doc.append(NoEscape(r"\addcontentsline{toc}{chapter}{\listtablename}")) 
    doc.append(NoEscape(r"\thispagestyle{empty}"))         

def ATitle(doc):
    doc.append(NoEscape(r"\begin{figure}")) 
    doc.append(NoEscape(r"\includegraphics[outer]{ABBILDUNGEN/Logo_Stadt_Flensburg.png}")) 
    doc.append(NoEscape(r"\end{figure}")) 
    doc.append(NoEscape(r"\begin{flushleft}")) 
    doc.append(NoEscape(r"\textbf{\LARGE{Sozialatlas 2020}}")) 
    doc.append(NoEscape(r"\end{flushleft}")) 
    doc.append(NoEscape(r"\begin{flushleft}")) 
    doc.append(NoEscape(r"\textbf{\Large{Datenbasis bis 31.12.2019}")) 
    doc.append(NoEscape(r"\end{flushleft}")) 
    doc.append(NoEscape(r"\begin{flushleft}")) 
    doc.append(NoEscape(r"\textbf{\Large{Stadt Flensburg \\")) 
    doc.append(NoEscape(r"Fachbereich Soziales und Gesundheit}}")) 
    doc.append(NoEscape(r"\end{flushleft}"))    

def einleitung(doc, year):
     emptyrow = '''
'''
     doc.append(NoEscape(r'\newpage'))
     doc.append(NoEscape(r"\textbf{\Large{Einleitung}}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\marginnote{\emph{Ziel: kontinuierliche Beobachtung der sozialen Lage}}[0.25cm]"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\emph{Mit dem Sozialatlas "+str(year)+" liegt die neunzehnte kleinr??umige Fortschreibung von Sozialstrukturdaten f??r die Stadt Flensburg und ihre 13 Stadtteile vor. Das Ziel ist eine differenzierte Beobachtung von relevanten Indikatoren, die Aufschluss ??ber die soziale Lage in Flensburg geben. Der Sozialatlas liefert damit wichtige Grundinformationen f??r Planungen, wie z.B. in der Jugendhilfe, im Bereich ??lterer Menschen oder der Stadtplanung.}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\marginnote{\emph{Orientierung nach Stadtteilen}}[0.25cm]"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\emph{Der Sozialatlas ist eine dauerhafte Aufgabe. Die Orientierung nach Sozialr??umen ??? in diesem Fall nach Stadtteilen ??? bedeutet, dass kleinr??umige, sozio??konomische und demografische Daten im Hinblick auf die soziale Lebensrealit??t der Bewohner*innen untersucht und analysiert werden. Durch die Fortschreibung k??nnen langfristig kleinr??umige Ver??nderungen nachgezeichnet werden, z.B. in der Altersstruktur der Bev??lkerung, der Erwerbst??tigkeit oder im Bezug von Sozialleistungen. Dabei erfolgt die Darstellung der Entwicklung der Bev??lkerungsdaten in einem 10-Jahresvergleich. Die themenspezifischen Informationen werden in der Regel in einem 5-Jahresr??ckblick betrachtet.}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\marginnote{\emph{Stichtag 31.12."+str(year-1)+"}}[0.25cm]"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\emph{Der Sozialatlas zum Stichtag 31.12."+str(year-1)+" ist in f??nf Kapitel unterteilt: 1. Bev??lkerung, 2. Arbeitsmarkt und Besch??ftigung, 3. Wohnen, 4. Soziale Sicherung und 5. Hilfen zur Erziehung. Den Kapiteln ist eine Zusammenfassung der wichtigsten Ergebnisse vorangestellt. Die umrandeten Textbl??cke weisen auf allgemeing??ltige Informationen hin. Am Ende befinden sich kurze Steckbriefe f??r die Stadt Flensburg und die 13 Stadtteile.} "))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\emph{Sofern nicht anders angegeben, handelt es sich bei allen Abbildungen und Tabellen um Darstellungen des Fachbereichs Soziales und Gesundheit auf Basis von Daten der Statistikstelle der Stadt Flensburg. Daten f??r die Jahre vor 2015 wurden in der Regel aus vorhergehenden Sozialatlanten ??bernommen.}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\textbf{\emph{Bev??lkerung}}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\marginnote{\emph{Zensus nicht ber??cksichtigt}}[0.25cm]"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\emph{Die Klage der Stadt Flensburg (aus 2015) gegen das Ergebnis des Zensus von 2011 befindet sich weiterhin in der rechtlichen Kl??rung. Daher werden auch weiterhin Daten auf Basis des st??dtischen Melderegisters verwendet (bis auf externe Quellen und Verweise). Im Gegensatz zu den Zahlen des Statistischen Amts f??r Hamburg und Schleswig-Holstein k??nnen die Daten des Melderegisters zudem kleinr??umig ausgewertet werden. Des Weiteren wird die Vergleichbarkeit zu den Daten der vorherigen Sozialatlanten gewahrt.}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\emph{Im Vordergrund steht die Entwicklung und strukturelle Zusammensetzung der Bev??lkerung nach Alter, Geschlecht und Herkunft. Dar??ber hinaus dargestellt sind die Geburtenentwicklung sowie wichtige Kennzahlen zur demografischen Entwicklung. Des Weiteren enth??lt der Sozialatlas Angaben zum Aufenthaltsstatus der in Flensburg lebenden ausl??ndischen Einwohner*innen und zu den Einb??rgerungen.}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\textbf{\emph{ und Besch??ftigung}}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\marginnote{\emph{sozialversicherungspflichtige Besch??ftigung und Arbeitslosigkeit}}[0.25cm]"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\emph{Kleinr??umige Daten sind f??r die Themen Arbeitslosigkeit und sozialversicherungspflichtig Besch??ftigte verf??gbar. Datengrundlage ist die Statistik der Bundesagentur f??r Arbeit.}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\newpage"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\textbf{\emph{Wohnen}}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\marginnote{\emph{Wohngeld und Wohnungs-hilfef??lle}}[0.25cm]"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\emph{Im Sozialatlas werden Daten zum Bezug von Wohngeld und ??ber Wohnungshilfef??lle ausgewertet. Sie werden vom B??rgerb??ro bzw. der Fachstelle f??r Wohnhilfen und Schuldnerberatung zur Verf??gung gestellt.}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\textbf{\emph{Soziale Sicherung}}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\marginnote{\emph{Leistungen nach SGB II, III und XII}}[0.25cm]"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r'\emph{Der Abschnitt enth??lt detaillierte Angaben zu den Bezieher*innen von Leistungen nach den Sozialgesetzb??chern (SGB) II, III und XII. Dargestellt werden im Wesentlichen die drei Altersgruppen ???unter 15 Jahren???, ???15 bis unter 65 Jahren??? und der Personen im Alter von ???65 Jahren und ??lter". Vor dem Hintergrund der Armutsdiskussion ist der Blick insbesondere darauf gerichtet, wie viele Personen im Bezug staatlicher Leistungen leben und damit ??berwiegend deutlich weniger Einkommen zur Verf??gung haben als der Durchschnitt.}'))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\textbf{\emph{Hilfen zur Erziehung}}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\marginnote{\emph{HzE-Daten f??r die Gesamtstadt}}[0.25cm]"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\emph{Die Darstellung der Hilfen zur Erziehung (HzE) beschr??nkt sich auf die Entwicklung der Fallzahlen f??r die Gesamtstadt. Als Datengrundlage sind ausschlie??lich zahlbare Leistungsf??lle verf??gbar.}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\textbf{\emph{Stadtteil-Steckbriefe}}"))
     doc.append(NoEscape(emptyrow))
     doc.append(NoEscape(r"\emph{Die Steckbriefe fassen die wesentlichen Ergebnisse f??r die einzelnen Stadtteile ??bersichtlich zusammen. Neben dem aktuellen Trend ist eine Vergleichszahl des aktuellen Jahres f??r die Stadt Flensburg eingepflegt. Die Trendpfeile stellen einen Vergleich des aktuellen Wertes mit dem Durchschnitt der letzten drei Jahre dar, eine Ver??nderung um mehr als 10\% des Durchschnittswertes wird dabei als relevant erachtet.}"))

def impressum(doc):
    from datetime import date
    emptyrow = '''
'''
    today = date.today()

    # dd.mm.YY
    d1 = today.strftime("%d.%m.%Y")

    doc.append(NoEscape(r'\newgeometry{asymmetric,outer=2cm,inner=2cm,top=2cm,bottom=3cm}'))
    doc.append(NoEscape(r'\vspace*{14cm}'))
    doc.append(NoEscape(r'\textbf{Herausgebend:} \newline'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'Stadt Flensburg \newline'))
    doc.append(NoEscape(r'- Die Oberb??rgermeisterin - \newline'))
    doc.append(NoEscape(r'Fachbereich Soziales und Gesundheit \newline'))
    doc.append(NoEscape(r'Rathausplatz 1 \par'))
    doc.append(NoEscape(r'24937 Flensburg \par'))
    doc.append(NoEscape(r'Telefon: 0461 85-1241'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'Stand: ' + str(d1) ))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'\restoregeometry'))

def vorwort(doc, year):
    emptyrow = '''
'''
    doc.append(NoEscape(r'\newgeometry{asymmetric,outer=1.5cm,inner=1.5cm,top=1.5cm,bottom=1.5cm}'))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(r'\textbf{\large{Vorwort}}'))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(r'\begin{wrapfigure}{r}{5.5cm}'))
    doc.append(NoEscape(r'\centering'))
    doc.append(NoEscape(r'\includegraphics[scale=0.3]{ABBILDUNGEN/2020_Dezernentin_KWN.png}'))
    doc.append(NoEscape(r'\end{wrapfigure}'))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(r'Liebe Leser*innen,'))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(r'mit dem Sozialatlas '+str(year)+' legt der Fachbereich Soziales und Gesundheit der Stadtverwaltung Flensburg die 19. kleinr??umige Fortschreibung von bev??lkerungs- und sozialstrukturbezogenen Daten f??r die Stadt Flensburg vor. '))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(r'Ziel der Flensburger Sozialberichterstattung ist es, die soziale Lebenswirklichkeit der Flensburger*innen mittels aussagekr??ftiger Kennzahlen abzubilden, um gem???? dem Leitspruch \glqq Daten f??r Taten\grqq{} empirisch fundierte Planungs- und Steuerungsprozesse zu initiieren und auf diese Weise zielgerichtetetes Planen und Handeln, auf Politik-, Verwaltungs- sowie auch Tr??gerebene, zu erm??glichen. '))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(r'Die j??hrliche Fortschreibung sowie die kleinr??umige Betrachtung der Daten erlauben das Sichtbarmachen demografischer und sozialstruktureller Entwicklungen. Indem bspw. Ver??nderungen in der Altersstruktur der Bev??lkerung, der Erwerbst??tigkeit oder bei dem Bezug von Sozialleistungen aufgezeigt werden, k??nnen soziale ??nderungsprozesse registriert und ggf. gegensteuernde Ma??nahmen eingeleitet werden.'))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(r'Vor dem Hintergrund der Armutsdiskussion wird ein besonderes Augenmerk auf die Bezieher*innen von Leistungen nach den Sozialgesetzb??chern (SGB) II, III und XII gerichtet (s. Kap. 4 Soziale Sicherung). Es gibt mittlerweile zahlreiche empirische Befunde, die auf den Zusammenhang zwischen Einkommen, Bildung und Gesundheit hinweisen, die dabei in komplexen Wechselwirkungen stehen. Dabei zeigt sich immer wieder, dass Ma??nahmen ihre Wirksamkeit nicht isoliert entfalten, sondern oft auch Auswirkungen in benachbarten Bereichen mit sich bringen. Angesichts verschiedener sozialer Problemlagen stellt dies eine gro??e Herausforderung dar, die aber auch als Chance zu verstehen und wahrzunehmen ist. Denn wenn durch datenbasiertes und zielgerichtetes Handeln positive Entwicklungen angesto??en werden, kann davon ausgegangen werden, dass diese - wie eine positive Kettenreaktion - zu einem erweiterten Wirkungskreis einzelner Ma??nahmen f??hren.'))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(r'Soziale Wirklichkeit wird tagt??glich vor Ort, in den einzelnen Lebenswelten, ob in der Schule, bei der Arbeit oder im Verein, in den Sozialr??umen dieser Stadt, in jedem der 13 Stadtteile geschaffen. Damit liegt ein bedeutender Anteil der Gestaltungshoheit sozialer Prozesse auch bei Ihnen, den Einwohner*innen dieser wunderbaren Stadt zwischen Himmel und F??rde. Daher m??chte ich Sie dazu einladen, die Lekt??re dieses Berichtes zum Anlass zu nehmen, sich Ihrer Bedeutung und Ihres Wirkpotenzials als Mitglied einer gr????eren Verantwortungsgemeinschaft bewusst zu werden. Sie nehmen es bereits tagt??glich wahr, sei es in Form von Nachbarschaftshilfe, Vereinsarbeit oder in jeglicher anderen Form sozialen Engagements, und wir danken Ihnen an dieser Stelle ausdr??cklich f??r Ihren t??glichen Beitrag!'))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(r'Flensburg, November '+ str(year)))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(r'\begin{figure}[H]'))
    doc.append(NoEscape(r' \includegraphics[scale=0.25]{ABBILDUNGEN/Unterschrift_Dezernentin_KWN.png}'))
    doc.append(NoEscape(r'\end{figure}'))
    doc.append(NoEscape(r'Karen Welz-Nettlau,'))
    doc.append(NoEscape(r'Dezernentin f??r Jugend, Soziales, Gesundheit und Zentrale Dienste'))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(r'\underline{\textcolor{red}{Hinweis zur Corona-Pandemie:}}Da diese Ausgabe sich auf die Datenlage zum Stichtag 31.12.2019 bezieht, k??nnen anhand der Datenauswertungen noch keine Hinweise auf die sozialen Auswirkungen der Covid-19-Pandemie abgeleitet werden. So werden fr??hestens mit dem Sozialatlas 2021 (auf Datenbasis zum 31.12.2020) erste Hinweise auf pandemiebedingte Auswirkungen sichtbar werden. Nichtsdestotrotz steht bereits zum heutigen Zeitpunkt fest, dass infolge der Pandemielage mit nachhaltigen, sozialen Auswirkungen zu rechnen ist. Es ist davon auszugehen, dass diese sich besonders in den Bereichen Besch??ftigung und soziale Sicherung, aber bspw. auch bei den Hilfen zur Erziehung niederschlagen werden.'))
    doc.append(NoEscape(emptyrow))
    doc.append(NoEscape(r'\afterpage{\aftergroup\restoregeometry}'))
    doc.append(NoEscape(r'\cleardoubleoddpage'))
    
def zusammenfassung(doc):
    doc.append(NoEscape(r'\textbf{\large{Zusammenfassung}}'))
    


def bevoelkerung(doc, year,directoryabbbevoelkerung,directoryabbbevoelkerungserver,datadirectory ):
    """Add a section, a subsection and some text to the document.

    :param doc: the document
    :type doc: :class:`pylatex.document.Document` instance
    """
    emptyrow = '''
'''
    with doc.create(Chapter('Bev??lkerung', label="1Bevoelkerung")):
        with doc.create(Section('Bev??lkerungsentwicklung')):
            doc.append(NoEscape(emptyrow))
            doc.append(NoEscape(r'\marginnote{\emph{Einfl??sse auf die ' +
                                'Bev??lkerungsentwicklung}}[0.2cm]'))
            doc.append(NoEscape(emptyrow))
            doc.append(NoEscape(r'\fbox{\parbox{\columnwidth}{\emph{Die Bev??lkerungsentwicklung ergibt sich aus der Differenz zwischen Geburten- und Sterberate in Verbindung mit dem Wanderungssaldo. Dieser wird von verschiedenen Faktoren beeinflusst: von globalen politischen Entwicklungen, Tendenzen auf dem Arbeitsmarkt (z.B. Anzahl der offenen und vermittelbaren Stellen), dem Wohnraumangebot (z.B. Mietpreise, freie Wohnkapazit??ten, Wohnraumqualit??t), durch die Bildungsinfrastruktur (z.B. Angebot an Kindertagesst??tten und Schulen bzw. Hochschulen), das Angebot an beruflichen Ausbildungen sowie durch pers??nliche oder famili??re Entscheidungen ??ber den Hauptwohnsitz.}}}'))
            doc.append(NoEscape(emptyrow))
            doc.append(NoEscape(r' \marginnote{\scriptsize{*Hinweis: Um die Ver??nderungen besser sichtbar zu machen, beginnt die y-Achse bei 84.000 statt bei 0.}} [0.25cm]'))
            
            #######################################################Get Data Population ##########################################################
            population = pd.read_csv(datadirectory+"//Population//Population.csv", sep = ";")

            filenameabb1 = 'Bev??lkerungsentwicklung '+ str(year-11) +' bis '+ str(year-1) +'.png'
            pathabb1 = directoryabbbevoelkerung + '//' + filenameabb1
            directoryabbbevoelkerungserver = directoryabbbevoelkerung + "//" + filenameabb1            
        
            filteryears1 = (population['year'] >= year-11) 
            
            populationofelevenyears = population[filteryears1]
            
            
            filteryears2 = (populationofelevenyears['year'] < year)
            populationofelevenyears = populationofelevenyears[filteryears2]          
            minpopulation = populationofelevenyears["population"].min()  
            maxpopulation = populationofelevenyears["population"].max()
            currentpopulation = populationofelevenyears['population'][(populationofelevenyears['year'] == year-1)].values[0]
            prevpopulation = populationofelevenyears['population'][(populationofelevenyears['year'] == year-11)].values[0]
            prevyear = populationofelevenyears['population'][(populationofelevenyears['year'] == year-2)].values[0]

            ################################################Population Chart####################################
            bars = alt.Chart(populationofelevenyears).mark_bar(size = 20, color='#4F81BD').encode(
            alt.X('year:O',
                  axis=alt.Axis(labels=True, title='Jahr'),
                  #scale=alt.Scale(domain=(year-11, year-1))
                  ),
            alt.Y('population:Q',
                axis=alt.Axis(labels=True, title='Bev??lkerung'),
                scale=alt.Scale(domain=(round(minpopulation*0.97), round(maxpopulation*1.03))),
                ),
            )
            
            text = bars.mark_text(
                align='center',
                baseline='bottom',
                #dx=3  # Nudges text to right so it doesn't appear on top of the bar
                dy = -10
            ).encode(
                text='population'
            )
            
            (bars + text).properties(width=500).save(directoryabbbevoelkerungserver)
        
            ################################################Population Chart####################################
            
                   
            
            doc.append(NoEscape(r'\begin{figure}[H]'))
            doc.append(NoEscape(r' \includegraphics[width=\textwidth]{'+pathabb1+'}'))
            doc.append(NoEscape(r' \caption{\textbf{Bev??lkerungsentwicklung '+ str(year-11) +' bis '+ str(year-1) +' (ohne Ber??cksichtigung Zensus 2011).}}'))
            doc.append(NoEscape(r' \label{fig:Abbildung_1}'))
            doc.append(NoEscape(r'\end{figure}'))
            doc.append(NoEscape(emptyrow))
            doc.append(NoEscape(r'\marginnote{\emph{Bev??lkerungszunahme seit dem Jahr '+ str(year-11) +', aktueller Stand: '+ str(currentpopulation) +'}}[0.2cm]'))
            doc.append(NoEscape(emptyrow))   
            currentpopulationstring = r'Zum 31.12.'+ str(year-1) + ' waren '+ str(currentpopulation) +' Personen mit ihrem Hauptwohnsitz in Flensburg gemeldet. ' 
            highest_population = r'Das ist der h??chste Bev??lkerungsstand seit fast 60 Jahren. '     
            if (prevpopulation > currentpopulation):
                 compare_curr_prev_population = r"Im Vergleich zu "+ str(year-11) +" ist die Einwohnerzahl Flensburgs um "+ str(prevpopulation-currentpopulation) +" Personen gesunken (vgl. Abb. \\ref{fig:Abbildung_1}).";
            elif(prevpopulation < currentpopulation):
                compare_curr_prev_population = r"Im Vergleich zu "+ str(year-11) +" ist die Einwohnerzahl Flensburgs um "+ str(currentpopulation-prevpopulation) +" Personen angewachsen (vgl. Abb. \\ref{fig:Abbildung_1}).";
            elif (prevpopulation == currentpopulation) :
                compare_curr_prev_population = r"Im Vergleich zu "+ str(year-11) +" ist die Einwohnerzahl Flensburgs konstant geblieben (vgl. Abb. \\ref{fig:Abbildung_1}).";
            
            average_population = r'Zwischen 2010 und 2014 ist die Stadt durchschnittlich um ca. 500 Personen j??hrlich gewachsen. In den Jahren 2015 bis 2017 sind mehr als 1.000 Einwohner*innen pro Jahr hinzugekommen, in 2015 waren es sogar 1.796.'
                                
            prevyear = r'Im Vergleich zum Vorjahr \marginnote{\emph{h??chste Einwohnerzahl seit fast 60 Jahren}}[0cm] steigt die Bev??lkerung um 0,74\% (+716 Personen). Damit bewegt sich der Zuwachs auf dem Vorjahresniveau.'
            
            population_development = currentpopulationstring + highest_population + compare_curr_prev_population + average_population + prevyear

            doc.append(NoEscape(population_development))
            doc.append(NoEscape(emptyrow))
            
            doc.append(NoEscape(r'\fcolorbox{black}{PaleAqua}{\parbox{\columnwidth}{\footnotesize{\underline{Hinweis}: F??r die Jahre ab 2011 hat das Statistische Amt f??r Hamburg und Schleswig-Holstein (Statistikamt Nord) auf Grundlage der Ergebnisse des Zensus 2011 eine deutlich unter den bisherigen Ergebnissen liegende Bev??lkerungszahl (82.258 zum Stichtag 31.12.2011) f??rmlich festgesetzt. Das Flensburger Einwohnermelderegister wies im Vergleich eine Einwohner*innenzahl von 89.532 Personen aus. Die Stadt Flensburg hat 2015 gegen die Ergebnisse des Zensus geklagt. Die rechtliche Kl??rung des Sachverhaltes ist derzeit noch nicht abgeschlossen. Daher beziehen sich alle nachfolgenden Angaben zu den Bev??lkerungszahlen weiterhin auf Datenbest??nde des st??dtischen Einwohnermelderegisters. Im Gegensatz zu den Zahlen des Statistikamts Nord k??nnen die Daten des Melderegisters zudem kleinr??umig ausgewertet werden. Des Weiteren wird die Vergleichbarkeit mit den Daten der Sozialatlanten der Vorjahre gew??hrleistet.}}}'))
            doc.append(NoEscape(emptyrow))
            
            
            
            doc.append(NoEscape(r'\subsubsection{a) kleinr??umige Entwicklung}'))
            doc.append(NoEscape(emptyrow))
            doc.append(NoEscape(r'\marginnote{\emph{Zunahme der Bev??lkerung in fast allen Stadtteilen}}[0.25cm]')) 
            doc.append(NoEscape(emptyrow))
            doc.append(NoEscape(r'Im Vergleich zu 2009 weisen fast alle Stadtteile ein kontinuierliches Wachstum auf (vgl. Tab. \ref{tab:Tabelle_1} und Abb. \ref{fig:Abbildung_2}). Es zeigen sich jedoch gro??e Unterschiede hinsichtlich der Wachstumsraten in den einzelnen Stadtteilen. So ist f??r die Nordstadt, M??rwik und Tarup ein Zuwachs um mehr als 1.000 Einwohner*innen seit 2009 zu verzeichnen. Auch in den anderen Stadtteilen ist eine allgemeine Zunahme der Bev??lkerungszahlen zu konstatieren. Lediglich in Engelsby ist die Einwohner*innenzahl zur??ckgegangen (-377 gg??. 2009). Im Vergleich zum Vorjahr sind insbesondere M??rwik (+286) und Tarup (+188) gewachsen. Am st??rksten geht die Einwohner*innenzahl im Stadtteil Engelsby zur??ck (-94 gg??. 2018).'))            
            doc.append(NoEscape(emptyrow))
            doc.append(NoEscape(emptyrow))
            doc.append(NoEscape(r'\marginnote{\scriptsize{* Einwohner*innen mit Haupt- oder alleiniger Wohnung. Geringf??gige Abweichungen ergeben sich durch nicht zuordenbare Personen.}}[1.25cm]'))
            
            #Districts
            districts = pd.read_csv(datadirectory+"//Population//District.csv", sep = ";",encoding = "ISO-8859-1")
            districts.loc[13] = districts.sum()
            districts.loc[13," "] = "Flensburg"
            difference_coloumn = districts.iloc[:,6] -districts.iloc[:,1]
            procentual_coloumn = round(100 / districts.iloc[:,1] * difference_coloumn,1)
            districts.insert(7, 'Ver??nderung '+ str(year-11) + ' - ' + str(year-1), difference_coloumn)
            districts.insert(8, "", procentual_coloumn)
            districtscopy = districts.copy(deep=True)
            
            districts.loc[-2] = districts.columns
            districts.loc[-1] = ['', '', '','', '', '','', 'absolut', 'prozentual']  # adding a row
            districts.index = districts.index + 2 
            districts.sort_index(inplace=True)
            

            #print(str(districts))
            #multirowlist = [Multirow(1,0,1),Multirow(2,0,1),Multirow(3,0,1),Multirow(4,0,1),Multirow(5,0,1),Multirow(6,0,1),Multirow(7,0,1),Multirow(8,0,1)]
            multirowlist = [Multirow(0,0,1),Multirow(1,0,1),Multirow(2,0,1),Multirow(3,0,1),Multirow(4,0,1),Multirow(5,0,1),Multirow(6,0,1)]         
            multicolumlist = [Multicol(0,7,8)]
            createtable(doc,districts,fontsize="scriptsize", columncolor = \
                        [(0,"Solitude"),(1,"Solitude"),(2,"Solitude"),(3,"Solitude") \
                         ,(4,"Solitude"),(5,"Solitude"),(6,"Solitude"), \
                         (7,"TropicalBlue"),(8,"TropicalBlue"),], framecolor = \
                            "white", align = [(0,"l")], rowstretch=1.5, fontweightrowlist = [0,15], addcoloumnnames= False, multrow = multirowlist, multcolumn = multicolumlist, label = "Tabelle_1")
            


            districtscopy = districtscopy.rename(columns = {'': 'bve'}, inplace = False)
            districtscopy.insert(9, "bve2", districtscopy.iloc[:,8].divide(100))
            print(districtscopy.iloc[:,8].divide(100))
            filenameabb2 = 'Bev??lkerungsentwicklung in den Stadtteilen '+ str(year-11) +' bis '+ str(year-1) +'.png'
            pathabb2 = directoryabbbevoelkerung + '//' + filenameabb2
            directoryabbbevoelkerungserver = directoryabbbevoelkerung + "//" + filenameabb2 

            print(districtscopy.head())
            
            base = alt.Chart(
                districtscopy, 
            ).properties(width=700)
            
            bars = base.mark_bar().encode(
                x=alt.X(
                    " ",
                    axis=alt.Axis(labels=True, title='Stadtteile',labelAngle=-45),
                ),
                y=alt.Y('bve2',
                axis=alt.Axis(labels=True, title='Bev??lkerungsentwicklung', format='%'),
                scale=alt.Scale(domain=[
                    districtscopy['bve2'].min()-0.1,
                    districtscopy['bve2'].max()+0.1
                    ])
                
                ),
            )
            text = base.encode(
                x=alt.X(" "),
                y="bve2",
                text=alt.Text('bve2', format='.1%'),
            )
            
            textAbove = text.transform_filter(alt.datum.bve2 > 0).mark_text(
                align='center',
                baseline='middle',
                dy=-10
                
            )
            
            textBelow = text.transform_filter(alt.datum.bve2 < 0).mark_text(
                align='center',
                baseline='middle',
                dy=12
            )
            
            (bars + textAbove + textBelow).properties(width=500).save(directoryabbbevoelkerungserver)
           
            
            doc.append(NoEscape(r'\begin{figure}[H]'))
            doc.append(NoEscape(r'\includegraphics[width=\textwidth]{'+pathabb2+'}'))
            doc.append(NoEscape('\\caption{\\textbf{Bev??lkerungsentwicklung in den Stadtteilen '+ str(year-11) +' bis '+ str(year-1)+'}}'))
            doc.append(NoEscape(r'\label{fig:Abbildung_2}'))
            doc.append(NoEscape(r'\end{figure}'))
            doc.append(NoEscape(emptyrow))
            doc.append(NoEscape(r'Die st??rksten, prozentualen Zuw??chse seit '+ str(year-11) + r' weisen die Stadtteile Tarup, Neustadt und Weiche  auf (s. Abb. \ref{fig:Abbildung_2}). Allein in Engelsby ist ein R??ckgang der Einwohner*innenzahl zu verzeichnen.'))
            doc.append(NoEscape(emptyrow))
            doc.append(NoEscape(r'\subsubsection{b) Geburtenentwicklung}'))
            doc.append(NoEscape(emptyrow))
            doc.append(NoEscape(r'\marginnote{\emph{Anstieg der Geburtenzahl in 2019}}[0.25cm]'))
            doc.append(NoEscape(r'In 2019 lag die Geburtenzahl bei 1.007 Geburten (s. Abb. \ref{fig:Abbildung_3}). Das sind 96 Geburten mehr als im Vorjahr (+10,5\% gg??. 2018). Im Vergleich zu 2009 ist die Zahl der Geburten um 23,0\% gestiegen (+188 Geburten). Zwischen 2009 und 2018 waren es durchschnittlich 843 Geburten pro Jahr. Die Geburtenzahlen der letzten 4 Jahren liegen deutlich dar??ber.'))
            doc.append(NoEscape(r'\marginnote{\scriptsize{*Hinweis: Um die Ver??nderungen besser sichtbar zu machen, beginnt die y-Achse bei 500 statt bei 0.}} [0.25cm]'))
            
             #######################################################Get Data Birthdayrate data ##########################################################
            birthrate = pd.read_csv(datadirectory+"//Population//Birth-Rate.csv", sep = ";")

            filenameabb1 = 'Geburtentwicklung '+ str(year-11) +' bis '+ str(year-1) +'.png'
            pathabb1 = directoryabbbevoelkerung + '//' + filenameabb1
            directoryabbbevoelkerungserver = directoryabbbevoelkerung + "//" + filenameabb1            
        
            filteryears1 = (birthrate['year'] >= year-11) 
            
            birthrateofelevenyears = birthrate[filteryears1]

            ################################################Birth rate Chart####################################
            bars = alt.Chart(birthrateofelevenyears).mark_bar(size = 20, color='#4F81BD').encode(
            alt.X('year:O',
                  axis=alt.Axis(labels=True, title='Jahr'),
                  #scale=alt.Scale(domain=(year-11, year-1))
                  ),
            alt.Y('birthrate:Q',
                axis=alt.Axis(labels=True, title='Anzahl der Geburten'),
                                    scale=alt.Scale(domain=[
                                    birthrateofelevenyears['birthrate'].min()*0.7,
                                    birthrateofelevenyears['birthrate'].max()*1.1
                                    ])
                ),
            )
            
            text = bars.mark_text(
                align='center',
                baseline='bottom',
                #dx=3  # Nudges text to right so it doesn't appear on top of the bar
                dy = -10
            ).encode(
                text='birthrate'
            )
            
            (bars + text).properties(width=500).save(directoryabbbevoelkerungserver)
        
            ################################################BirthRate Chart####################################
            doc.append(NoEscape(r'\begin{figure}[H]'))
            doc.append(NoEscape(r' \includegraphics[width=\textwidth]{'+pathabb1+'}'))
            doc.append(NoEscape(r' \caption{\textbf{Geburtenentwicklung '+ str(year-11) +' bis '+ str(year-1) +'.}}'))
            doc.append(NoEscape(r' \label{fig:Abbildung_3}'))
            doc.append(NoEscape(r'\end{figure}'))
            doc.append(NoEscape(emptyrow))
            doc.append(NoEscape('Die Geburtenquote stellt eine bedeutende demografische Kennziffer dar. Sie gibt die Anzahl der Geburten je 1.000 Frauen im Alter zwischen 15 und unter 45 Jahren an. '))
            doc.append(NoEscape('In 2019 liegt die Geburtenquote in Flensburg bei 51,4. Vergleichsweise hohe Geburtenquoten weisen M??rwik (64,5), Tarup (62,2) und die Nordstadt (60,8) auf. Die niedrigsten Geburtenquoten sind in den Stadtteilen Sandberg (32,7) und Altstadt (39,1) zu verzeichnen.'))
            doc.append(NoEscape(emptyrow))
            
            birthratedistricts = pd.read_csv(datadirectory+"//Population//Birth-RateDistricts.csv", sep = ";",encoding = "ISO-8859-1", header=None)
            #replace NAs
            birthratedistricts = birthratedistricts.fillna("")
            #Latex correct %
            birthratedistricts = birthratedistricts.replace(regex=r'%', value='\%')
            print(birthratedistricts)
            
            multirowlist = [Multirow(0,0,2),Multirow(1,0,1),Multirow(3,0,1),Multirow(5,0,1),Multirow(7,0,1),Multirow(9,0,1),Multirow(11,0,1)]         
            multicolumlist = [Multicol(0,1,2),Multicol(0,3,4),Multicol(0,5,6),Multicol(0,7,8),Multicol(0,9,10),Multicol(0,11,12),Multicol(0,13,14),Multicol(1,13,14)]
            createtable(doc,birthratedistricts,fontsize="tiny", columncolor = \
                        [(0,"Solitude"),(1,"Solitude"),(2,"Solitude"),(3,"Solitude") \
                         ,(4,"Solitude"),(5,"Solitude"),(6,"Solitude"), \
                         (7,"Solitude"),(8,"Solitude"),(9,"Solitude") \
                         ,(10,"Solitude"),(11,"Solitude"),(12,"Solitude"), \
                         (13,"TropicalBlue"),(14,"TropicalBlue"),], framecolor = \
                            "white", align = [(0,"l")], rowstretch=1.5, fontweightrowlist = [0,15], addcoloumnnames= False, multrow = multirowlist, multcolumn = multicolumlist, label = "Tabelle_1")
            
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))
            doc.append(NoEscape(''))


if __name__ == '__main__':
    
    t = [(3,4,5),(1,2,3)]

   
    #set chromedriver Path
    try:
        driver = webdriver.Chrome(r'C:\Program Files\Google\Chromedriver\chromedriver.exe')
        driver.close()
    except:
        driver = webdriver.Chrome()     
    
    # variable deklaration
    year = 2021; 
    
    directory = os.path.dirname(__file__) + '\\Sozialatlas_generated'
    directoryabb1 = directory + '\\ABBILDUNGEN'
    directoryabbbevoelkerung = directoryabb1 + '\\Abbildungen_1_Bev??lkerung'
    
    directoryabb1latex = 'ABBILDUNGEN'
    directoryabbbevoelkerunglatex = directoryabb1latex + '//Abbildungen_1_Bev??lkerung' 
    
    datadirectory = directory + '\\Data'
    
    #### Create Folder######
    try:
        os.stat(directory)
    except:
        os.mkdir(directory) 
    try:
        os.stat(directoryabb1)
    except:
        os.mkdir(directoryabb1) 
    try:
        os.stat(directoryabbbevoelkerung)
    except:
        os.mkdir(directoryabbbevoelkerung) 
        
    
    #output Path of the tex data
    os.chdir(path = directory )
    
    ######################################Create Diagrams

    
    
    
    
    # Main document
    ma = Document('main', documentclass="scrreprt",document_options="a4paper, 10.5pt, twoside, listof=entryprefix", lmodern=False, textcomp=False,page_numbers=False )
    latexmainstart(ma,year)
    
    
    
    ma.append(NoEscape(r"\thispagestyle{empty}")) 
#    doc.append(NoEscape(r"\include{0_Impressum}")) 
    impressum(ma)
    ma.append(NoEscape(r"\thispagestyle{empty}")) 
#    doc.append(NoEscape(r"\include{0_Vorwort}")) 
    vorwort(ma,year)

        
    # Fancyhdr zum Inhaltsverzeichnis,ListofTables und ListofFigures hinzuf??gen
    ma.append(NoEscape(r"\addtocontents{toc}{\protect\thispagestyle{fancy}}")) 
    ma.append(NoEscape(r"\addtocontents{lot}{\protect\thispagestyle{fancy}}")) 
    ma.append(NoEscape(r"\addtocontents{lof}{\protect\thispagestyle{fancy}}")) 
    ma.append(NoEscape(r"\tableofcontents")) 
    
    # Seitennummerierung f??r Inhaltsverzeichnis romanisch
    ma.append(NoEscape(r"\pagenumbering{arabic}")) 
    
#    doc.append(NoEscape(r"\include{0_Einleitung}"))
    einleitung(ma,year)
    
    ma.append(NoEscape(r"\addcontentsline{toc}{chapter}{Zusammenfassung}")) 
    
#    doc.append(NoEscape(r"\include{0_Zusammenfassung}"))
#    doc.append(NoEscape(r"\include{1Bevoelkerung}")) 
    bevoelkerung(ma,year,directoryabbbevoelkerunglatex, directoryabbbevoelkerung,datadirectory)

#    doc.append(NoEscape(r"\include{2_ArbeitsmarktundBesch??ftigung}")) 
#    doc.append(NoEscape(r"\include{3_Wohnen}")) 
#    doc.append(NoEscape(r"\include{4_SozialeSicherung}")) 
#    doc.append(NoEscape(r"\include{5_HilfenzurErziehung}")) 
#    doc.append(NoEscape(r"\input{6_??bersicht??berdieStadtteile}")) 
    ma.append(NoEscape(r"\newgeometry{left=2cm,right=2cm,top=2cm,bottom=3cm}")) 
#    doc.append(NoEscape(r"\include{6_??bersicht??berdieStadtteile}"))
    latexmainend(ma)
    ma.generate_tex()  
    
    
    
    
    # modularization with include didn't work
    # Bevoelkerung document
    #bev = Document('bev??lkerung', documentclass= False,fontenc=None,inputenc=None, font_size=None, lmodern=False, textcomp=False,page_numbers=False )
    #bevoelkerung(bev)
    #bev.generate_tex()    
    
    
    #ma.generate_pdf(clean_tex=False, compiler='pdflatex')
    

    # Document with `\maketitle` command activated


    #doc.preamble.append(Command('title', 'Sozialatlas' + str(jahr)))
    #doc.preamble.append(Command('author', 'Anonymous author'))
    #doc.preamble.append(Command('date', NoEscape(r'\today')))
    #doc.append(NoEscape(r'\maketitle'))
