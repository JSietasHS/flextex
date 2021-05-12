from pylatex import Document, Section, Subsection, Command, Package, \
    MultiRow, Tabular, MultiColumn
from pylatex.utils import italic, NoEscape
from pylatex.section import Chapter
import  altair  as  alt 
import numpy as np
import pandas as pd
import os

from altair_saver import save

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

    def __init__(self, row, startcolumn, endcolumn):
        self.row = row
        self.startcolumn = startcolumn
        self.endcolumn = endcolumn
        
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
    
    def set_data(self, data):
        self.data = MultiColumn(self.get_length(), align='|c|', data=data)
    
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
        self.data = MultiRow(self.get_length(), data=data)

    def get_length(self):
        return self.endrow +1 - self.startrow
 
    def checkMultirowInDF(self, row,column):
        return ((column == self.get_col()) and (row == self.get_startrow()))
    
def getColorstringFromList(colorlist, checkvalue):
    res = ""
    if colorlist is not None:
           for color in colorlist:
               if color[0] == checkvalue and color[1] is not None and color[2] is not None :
                   res = r"\cellcolor["+ str(color[1]) +"]{"+ str(color[2]) +"} "
    return res

#Create a Table out of a pandas Dataframe 
#input  document
#       dataframe: a pandas dataframe
#       coloumncolor: List 
#       coloumwithcolor list of integer #Paint the number of the coloumn in the 
#                               coloumncolor default value all coloumns
#       multirows # List of Triple [Rownumber,start coloumn, end coloumn]
#                   rownumber start at 0 the row of the table
#                   coloumnnumber start at 0 the start coloumn to connect coloumns
#       multicoloumns # # List of MultiCol elements to add multicoloumns        
def createtable(doc, df, multcolumn=None, multrow=None, columncolor = None, rowcolor=None):
    dfcopy = df.copy(deep=False) #copy of the dataframe
    #multcolumn = [Multicol(0,0,1)] #test
    #multrow = [Multirow(0,0,1)] #test
    #columncolor = [(1,"gray",0.8),(3,"gray",0.8)]
    #rowcolor = [(1,"gray",0.8),(3,"gray",0.8)]
    rowcellcolor = ""
    columncellcolor = ""
    addmcol = None # add multicoloumn
    addmrow = None # add multirow
    values = [] #add values to table
    i=0 #counter
    c='|' #coloumns
    
    all_columns = list(dfcopy) # Creates list of all column headers
    dfcopy[all_columns] = dfcopy[all_columns].astype(str) #Set all coloumnsto String
    
    #create table coloumns
    while i in range(0,len(df.columns)):
        c = c + 'c|'
        i = i+1
        
    table3 = Tabular(c)
    #iterate through rows
    for row in dfcopy.index:
        rowcellcolor = ""
        #print('row:' + str(row))
        values = []
        #iterate through coloumns
        rowcellcolor = getColorstringFromList(rowcolor,row)
        for columnind, column in enumerate(df.columns):
            addmcol = None
            addmrow = None
            cellcolor= ""
            columncellcolor = ""
            columncellcolor = getColorstringFromList(columncolor,columnind)
            if columncellcolor == "" :
                cellcolor = rowcellcolor
            else:
                cellcolor = columncellcolor
            
            if cellcolor != "":
                 doc.preamble.append(Package("colortbl"))
            
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
            #appen multiolrow
            if ((addmcol is not None) and (addmrow is not None)):
                for i in range(0,addmrow.get_length()):
                    for j in range(0,addmcol.get_length()):
                        data = data + str(dfcopy.iat[row+i,columnind+j])
                        if i > 0:
                            dfcopy.iat[row+i,columnind+j] = "" 
                        else:
                            dfcopy.iat[row+i,columnind+j] = None 
                addmrow.set_data(data)
                addmcol.set_data(addmrow.get_data())
                values.append(NoEscape(cellcolor + addmcol.get_data()))
            #append multicol
            elif (addmcol is not None):
                for i in range(0,addmcol.get_length()):
                    data = data + str(dfcopy.iat[row,columnind+i])
                    dfcopy.iat[row,(columnind+i)] = None                 
                addmcol.set_data(data) 
                values.append(NoEscape(cellcolor + addmcol.get_data()))
            #append miltirow
            elif (addmrow is not None):
                for i in range(0,addmrow.get_length()):
                    data = data + str(dfcopy.iat[row+i,columnind])
                    dfcopy.iat[row+i,columnind] = "" 
                addmrow.set_data(data) 
                values.append(NoEscape(cellcolor + addmrow.get_data()))
            #append normal value
            elif ((addmcol is None) and (addmrow is None)):
                #Dont append None but append ''
                if dfcopy.iat[row,columnind] is not None:
                    values.append(NoEscape(cellcolor + dfcopy.iat[row,columnind]))
        #print('Data:' + str(values))
        table3.add_row(values)
        table3.add_hline()
                
    #print(dfcopy)
    doc.append(table3)
                
    
      


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
    
    #ListofFigures & ListofTables Zahleneinträge werden fettgedruckt

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
    #Fancyhdr-Package genutzt zur Nutzung muss Pagestyle auf "Fancy" geändert werden die vorangestellten Buchstaben gelten der Platzierung, für eine individuelle Kopf- und Fußzeile im Dokument 
    doc.preamble.append(NoEscape(r"\pagestyle{fancy}")) 
    doc.preamble.append(NoEscape(r"\addtolength{\headwidth}{\marginparsep}")) 
    doc.preamble.append(NoEscape(r"\addtolength{\headwidth}{\marginparwidth}")) 
    doc.preamble.append(NoEscape(r"\fancyhead[RE,LO]{\textbf{Sozialatlas "+str(year)+"}}"))
    doc.preamble.append(NoEscape(r"\fancyhead[LE,RO]{\includegraphics[scale=0.5]{ABBILDUNGEN/Logo_Stadt_Flensburg.png}}"))
    doc.preamble.append(NoEscape(r"\fancyfoot[CE,CO]{}"))
    doc.preamble.append(NoEscape(r"\fancyfoot[LE,RO]{\thepage}"))
    
    doc.preamble.append(NoEscape(r"\renewcommand{\footrulewidth}{1pt}"))
    # Header u. Footer auch auf Chapter Pages anzeigen
    #Ändert das Pagestyle Argument auf fancyhdr in KOMA Klasse sccreprt (Etoolbox Package benötigt)
    doc.preamble.append(NoEscape(r"\renewcommand{\chapterpagestyle}{fancy}")) 
    #Fügt eine Leerzeile zwischen Überschriften/Absätzen dazu
    doc.preamble.append(NoEscape(r"\setlength{\parindent}{0pt}")) 
    
    doc.preamble.append(NoEscape(r"\raggedbottom"))
    
    
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
     doc.append(NoEscape(r"\textbf{\Large{Einleitung}}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\marginnote{\emph{Ziel: kontinuierliche Beobachtung der sozialen Lage}}[0.25cm]"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\emph{Mit dem Sozialatlas "+str(year)+" liegt die neunzehnte kleinräumige Fortschreibung von Sozialstrukturdaten für die Stadt Flensburg und ihre 13 Stadtteile vor. Das Ziel ist eine differenzierte Beobachtung von relevanten Indikatoren, die Aufschluss über die soziale Lage in Flensburg geben. Der Sozialatlas liefert damit wichtige Grundinformationen für Planungen, wie z.B. in der Jugendhilfe, im Bereich älterer Menschen oder der Stadtplanung.}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\marginnote{\emph{Orientierung nach Stadtteilen}}[0.25cm]"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\emph{Der Sozialatlas ist eine dauerhafte Aufgabe. Die Orientierung nach Sozialräumen – in diesem Fall nach Stadtteilen – bedeutet, dass kleinräumige, sozioökonomische und demografische Daten im Hinblick auf die soziale Lebensrealität der Bewohner*innen untersucht und analysiert werden. Durch die Fortschreibung können langfristig kleinräumige Veränderungen nachgezeichnet werden, z.B. in der Altersstruktur der Bevölkerung, der Erwerbstätigkeit oder im Bezug von Sozialleistungen. Dabei erfolgt die Darstellung der Entwicklung der Bevölkerungsdaten in einem 10-Jahresvergleich. Die themenspezifischen Informationen werden in der Regel in einem 5-Jahresrückblick betrachtet.}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\marginnote{\emph{Stichtag 31.12."+str(year-1)+"}}[0.25cm]"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\emph{Der Sozialatlas zum Stichtag 31.12."+str(year-1)+" ist in fünf Kapitel unterteilt: 1. Bevölkerung, 2. Arbeitsmarkt und Beschäftigung, 3. Wohnen, 4. Soziale Sicherung und 5. Hilfen zur Erziehung. Den Kapiteln ist eine Zusammenfassung der wichtigsten Ergebnisse vorangestellt. Die umrandeten Textblöcke weisen auf allgemeingültige Informationen hin. Am Ende befinden sich kurze Steckbriefe für die Stadt Flensburg und die 13 Stadtteile.} "))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\emph{Sofern nicht anders angegeben, handelt es sich bei allen Abbildungen und Tabellen um Darstellungen des Fachbereichs Soziales und Gesundheit auf Basis von Daten der Statistikstelle der Stadt Flensburg. Daten für die Jahre vor 2015 wurden in der Regel aus vorhergehenden Sozialatlanten übernommen.}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\textbf{\emph{Bevölkerung}}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\marginnote{\emph{Zensus nicht berücksichtigt}}[0.25cm]"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\emph{Die Klage der Stadt Flensburg (aus 2015) gegen das Ergebnis des Zensus von 2011 befindet sich weiterhin in der rechtlichen Klärung. Daher werden auch weiterhin Daten auf Basis des städtischen Melderegisters verwendet (bis auf externe Quellen und Verweise). Im Gegensatz zu den Zahlen des Statistischen Amts für Hamburg und Schleswig-Holstein können die Daten des Melderegisters zudem kleinräumig ausgewertet werden. Des Weiteren wird die Vergleichbarkeit zu den Daten der vorherigen Sozialatlanten gewahrt.}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\emph{Im Vordergrund steht die Entwicklung und strukturelle Zusammensetzung der Bevölkerung nach Alter, Geschlecht und Herkunft. Darüber hinaus dargestellt sind die Geburtenentwicklung sowie wichtige Kennzahlen zur demografischen Entwicklung. Des Weiteren enthält der Sozialatlas Angaben zum Aufenthaltsstatus der in Flensburg lebenden ausländischen Einwohner*innen und zu den Einbürgerungen.}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\textbf{\emph{Arbeitsmarkt und Beschäftigung}}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\marginnote{\emph{sozialversicherungspflichtige Beschäftigung und Arbeitslosigkeit}}[0.25cm]"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\emph{Kleinräumige Daten sind für die Themen Arbeitslosigkeit und sozialversicherungspflichtig Beschäftigte verfügbar. Datengrundlage ist die Statistik der Bundesagentur für Arbeit.}"))
     doc.append(NoEscape(r"\newpage"))
     doc.append(NoEscape(''))
     doc.append(NoEscape(r"\textbf{\emph{Wohnen}}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\marginnote{\emph{Wohngeld und Wohnungs-hilfefälle}}[0.25cm]"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\emph{Im Sozialatlas werden Daten zum Bezug von Wohngeld und über Wohnungshilfefälle ausgewertet. Sie werden vom Bürgerbüro bzw. der Fachstelle für Wohnhilfen und Schuldnerberatung zur Verfügung gestellt.}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\textbf{\emph{Soziale Sicherung}}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\marginnote{\emph{Leistungen nach SGB II, III und XII}}[0.25cm]"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r'\emph{Der Abschnitt enthält detaillierte Angaben zu den Bezieher*innen von Leistungen nach den Sozialgesetzbüchern (SGB) II, III und XII. Dargestellt werden im Wesentlichen die drei Altersgruppen „unter 15 Jahren“, „15 bis unter 65 Jahren“ und der Personen im Alter von „65 Jahren und älter". Vor dem Hintergrund der Armutsdiskussion ist der Blick insbesondere darauf gerichtet, wie viele Personen im Bezug staatlicher Leistungen leben und damit überwiegend deutlich weniger Einkommen zur Verfügung haben als der Durchschnitt.}'))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\textbf{\emph{Hilfen zur Erziehung}}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\marginnote{\emph{HzE-Daten für die Gesamtstadt}}[0.25cm]"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\emph{Die Darstellung der Hilfen zur Erziehung (HzE) beschränkt sich auf die Entwicklung der Fallzahlen für die Gesamtstadt. Als Datengrundlage sind ausschließlich zahlbare Leistungsfälle verfügbar.}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\textbf{\emph{Stadtteil-Steckbriefe}}"))
     doc.append(NoEscape(r'\\'))
     doc.append(NoEscape(r"\emph{Die Steckbriefe fassen die wesentlichen Ergebnisse für die einzelnen Stadtteile übersichtlich zusammen. Neben dem aktuellen Trend ist eine Vergleichszahl des aktuellen Jahres für die Stadt Flensburg eingepflegt. Die Trendpfeile stellen einen Vergleich des aktuellen Wertes mit dem Durchschnitt der letzten drei Jahre dar, eine Veränderung um mehr als 10\% des Durchschnittswertes wird dabei als relevant erachtet.}"))

def impressum(doc):
    from datetime import date

    today = date.today()

    # dd.mm.YY
    d1 = today.strftime("%d.%m.%Y")

    doc.append(NoEscape(r'\newgeometry{asymmetric,outer=2cm,inner=2cm,top=2cm,bottom=3cm}'))
    doc.append(NoEscape(r'\vspace*{14cm}'))
    doc.append(NoEscape(r'\textbf{Herausgebend:} \newline'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'Stadt Flensburg \newline'))
    doc.append(NoEscape(r'- Die Oberbürgermeisterin - \newline'))
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
    doc.append(NoEscape(r'\newgeometry{asymmetric,outer=1.5cm,inner=1.5cm,top=1.5cm,bottom=1.5cm}'))
    doc.append(NoEscape(r'\textbf{\large{Vorwort}}'))
    doc.append(NoEscape(r'\begin{wrapfigure}{r}{5.5cm}'))
    doc.append(NoEscape(r'\centering'))
    doc.append(NoEscape(r'\includegraphics[scale=0.3]{ABBILDUNGEN/2020_Dezernentin_KWN.png}'))
    doc.append(NoEscape(r'\end{wrapfigure}'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'Liebe Leser*innen,'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'mit dem Sozialatlas '+str(year)+' legt der Fachbereich Soziales und Gesundheit der Stadtverwaltung Flensburg die 19. kleinräumige Fortschreibung von bevölkerungs- und sozialstrukturbezogenen Daten für die Stadt Flensburg vor. '))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'Ziel der Flensburger Sozialberichterstattung ist es, die soziale Lebenswirklichkeit der Flensburger*innen mittels aussagekräftiger Kennzahlen abzubilden, um gemäß dem Leitspruch \glqq Daten für Taten\grqq{} empirisch fundierte Planungs- und Steuerungsprozesse zu initiieren und auf diese Weise zielgerichtetetes Planen und Handeln, auf Politik-, Verwaltungs- sowie auch Trägerebene, zu ermöglichen. '))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'Die jährliche Fortschreibung sowie die kleinräumige Betrachtung der Daten erlauben das Sichtbarmachen demografischer und sozialstruktureller Entwicklungen. Indem bspw. Veränderungen in der Altersstruktur der Bevölkerung, der Erwerbstätigkeit oder bei dem Bezug von Sozialleistungen aufgezeigt werden, können soziale Änderungsprozesse registriert und ggf. gegensteuernde Maßnahmen eingeleitet werden.'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'Vor dem Hintergrund der Armutsdiskussion wird ein besonderes Augenmerk auf die Bezieher*innen von Leistungen nach den Sozialgesetzbüchern (SGB) II, III und XII gerichtet (s. Kap. 4 Soziale Sicherung). Es gibt mittlerweile zahlreiche empirische Befunde, die auf den Zusammenhang zwischen Einkommen, Bildung und Gesundheit hinweisen, die dabei in komplexen Wechselwirkungen stehen. Dabei zeigt sich immer wieder, dass Maßnahmen ihre Wirksamkeit nicht isoliert entfalten, sondern oft auch Auswirkungen in benachbarten Bereichen mit sich bringen. Angesichts verschiedener sozialer Problemlagen stellt dies eine große Herausforderung dar, die aber auch als Chance zu verstehen und wahrzunehmen ist. Denn wenn durch datenbasiertes und zielgerichtetes Handeln positive Entwicklungen angestoßen werden, kann davon ausgegangen werden, dass diese - wie eine positive Kettenreaktion - zu einem erweiterten Wirkungskreis einzelner Maßnahmen führen.'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'Soziale Wirklichkeit wird tagtäglich vor Ort, in den einzelnen Lebenswelten, ob in der Schule, bei der Arbeit oder im Verein, in den Sozialräumen dieser Stadt, in jedem der 13 Stadtteile geschaffen. Damit liegt ein bedeutender Anteil der Gestaltungshoheit sozialer Prozesse auch bei Ihnen, den Einwohner*innen dieser wunderbaren Stadt zwischen Himmel und Förde. Daher möchte ich Sie dazu einladen, die Lektüre dieses Berichtes zum Anlass zu nehmen, sich Ihrer Bedeutung und Ihres Wirkpotenzials als Mitglied einer größeren Verantwortungsgemeinschaft bewusst zu werden. Sie nehmen es bereits tagtäglich wahr, sei es in Form von Nachbarschaftshilfe, Vereinsarbeit oder in jeglicher anderen Form sozialen Engagements, und wir danken Ihnen an dieser Stelle ausdrücklich für Ihren täglichen Beitrag!'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'Flensburg, November '+ str(year)))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'\begin{figure}[H]'))
    doc.append(NoEscape(r' \includegraphics[scale=0.25]{ABBILDUNGEN/Unterschrift_Dezernentin_KWN.png}'))
    doc.append(NoEscape(r'\end{figure}'))
    doc.append(NoEscape(r'Karen Welz-Nettlau,'))
    doc.append(NoEscape(r'Dezernentin für Jugend, Soziales, Gesundheit und Zentrale Dienste'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'\underline{\textcolor{red}{Hinweis zur Corona-Pandemie:}}Da diese Ausgabe sich auf die Datenlage zum Stichtag 31.12.2019 bezieht, können anhand der Datenauswertungen noch keine Hinweise auf die sozialen Auswirkungen der Covid-19-Pandemie abgeleitet werden. So werden frühestens mit dem Sozialatlas 2021 (auf Datenbasis zum 31.12.2020) erste Hinweise auf pandemiebedingte Auswirkungen sichtbar werden. Nichtsdestotrotz steht bereits zum heutigen Zeitpunkt fest, dass infolge der Pandemielage mit nachhaltigen, sozialen Auswirkungen zu rechnen ist. Es ist davon auszugehen, dass diese sich besonders in den Bereichen Beschäftigung und soziale Sicherung, aber bspw. auch bei den Hilfen zur Erziehung niederschlagen werden.'))
    doc.append(NoEscape(r'\\'))
    doc.append(NoEscape(r'\afterpage{\aftergroup\restoregeometry}'))
    doc.append(NoEscape(r'\cleardoubleoddpage'))
    
def zusammenfassung(doc):
    doc.append(NoEscape(r'\textbf{\large{Zusammenfassung}}'))
    


def bevoelkerung(doc, year,directoryabbbevoelkerung,directoryabbbevoelkerungserver,datadirectory ):
    """Add a section, a subsection and some text to the document.

    :param doc: the document
    :type doc: :class:`pylatex.document.Document` instance
    """
    with doc.create(Chapter('Bevölkerung', label="1Bevoelkerung")):
        with doc.create(Section('Bevölkerungsentwicklung')):
            doc.append(NoEscape(r'\textbf{\large{Beispiel}}'))
            doc.append(NoEscape(r'\marginnote{\emph{Einflüsse auf die ' +
                                'Bevölkerungsentwicklung}}[0.2cm]'))
            doc.append(NoEscape(r'\fbox{\parbox{\columnwidth}{\emph{Die Bevölkerungsentwicklung ergibt sich aus der Differenz zwischen Geburten- und Sterberate in Verbindung mit dem Wanderungssaldo. Dieser wird von verschiedenen Faktoren beeinflusst: von globalen politischen Entwicklungen, Tendenzen auf dem Arbeitsmarkt (z.B. Anzahl der offenen und vermittelbaren Stellen), dem Wohnraumangebot (z.B. Mietpreise, freie Wohnkapazitäten, Wohnraumqualität), durch die Bildungsinfrastruktur (z.B. Angebot an Kindertagesstätten und Schulen bzw. Hochschulen), das Angebot an beruflichen Ausbildungen sowie durch persönliche oder familiäre Entscheidungen über den Hauptwohnsitz.}}}'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r' \marginnote{\scriptsize{*Hinweis: Um die Veränderungen besser sichtbar zu machen, beginnt die y-Achse bei 84.000 statt bei 0.}} [0.25cm]'))
            
            #######################################################Get Data Population ##########################################################
            population = pd.read_csv(datadirectory+"//Population//Population.csv", sep = ";")

            filenameabb1 = 'Bevölkerungsentwicklung '+ str(year-11) +' bis '+ str(year-1) +'.png'
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
            #########################################################################################################
            populationchart = alt.Chart(populationofelevenyears).mark_bar(size = 20, color='#4F81BD').encode(
            alt.X('year',
                  axis=alt.Axis(labels=True, title='Jahr'),
                  scale=alt.Scale(domain=(year-10.5, year+0.5))
                  ),
            alt.Y('population',
                axis=alt.Axis(labels=True, title='Bevölkerung'),
                scale=alt.Scale(domain=(round(minpopulation*0.97), round(maxpopulation*1.03)))
                ),
            ).configure_axisX(
                labelAngle=90,
            )
            ################################################################################################################
            
            ################################################Population Chart####################################
            bars = alt.Chart(populationofelevenyears).mark_bar(size = 20, color='#4F81BD').encode(
            alt.X('year',
                  axis=alt.Axis(labels=True, title='Jahr', format = (" f")),
                  scale=alt.Scale(domain=(year-11, year-1))
                  ),
            alt.Y('population',
                axis=alt.Axis(labels=True, title='Bevölkerung'),
                scale=alt.Scale(domain=(round(minpopulation*0.97), round(maxpopulation*1.03)))
                ),
            )
            
            text = bars.mark_text(
                align='center',
                baseline='middle',
                #dx=3  # Nudges text to right so it doesn't appear on top of the bar
                dy = -10
            ).encode(
                text='population'
            )
            
            (bars + text).properties(width=500).save(directoryabbbevoelkerungserver)
        
            ################################################Population Chart####################################
            
                   
            
            doc.append(NoEscape(r'\begin{figure}[H]'))
            doc.append(NoEscape(r' \includegraphics[width=\textwidth]{'+pathabb1+'}'))
            doc.append(NoEscape(r' \caption{\textbf{Bevölkerungsentwicklung '+ str(year-11) +' bis '+ str(year-1) +' (ohne Berücksichtigung Zensus 2011).}}'))
            doc.append(NoEscape(r' \label{fig:Abbildung_1}'))
            doc.append(NoEscape(r'\end{figure}'))

            doc.append(NoEscape(r'\marginnote{\emph{Bevölkerungszunahme seit dem Jahr '+ str(year-11) +', aktueller Stand: '+ str(currentpopulation) +'}}[0.2cm]'))

            
            currentpopulationstring = r'Zum 31.12.'+ str(year-1) + ' waren '+ str(currentpopulation) +' Personen mit ihrem Hauptwohnsitz in Flensburg gemeldet. ' 
            highest_population = r'Das ist der höchste Bevölkerungsstand seit fast 60 Jahren. '
            
            if (prevpopulation > currentpopulation):
                 compare_curr_prev_population = r"Im Vergleich zu "+ str(year-11) +" ist die Einwohnerzahl Flensburgs um "+ str(prevpopulation-currentpopulation) +" Personen gesunken (vgl. Abb. \ref{fig:Abbildung_1}).";
            elif(prevpopulation < currentpopulation):
                compare_curr_prev_population = r"Im Vergleich zu "+ str(year-11) +" ist die Einwohnerzahl Flensburgs um "+ str(currentpopulation-prevpopulation) +" Personen angewachsen (vgl. Abb. \ref{fig:Abbildung_1}).";
            elif (prevpopulation == currentpopulation) :
                compare_curr_prev_population = r"Im Vergleich zu "+ str(year-11) +" ist die Einwohnerzahl Flensburgs konstant geblieben (vgl. Abb. \\ref{fig:Abbildung_1}).";
            
            average_population = r'Zwischen 2010 und 2014 ist die Stadt durchschnittlich um ca. 500 Personen jährlich gewachsen. In den Jahren 2015 bis 2017 sind mehr als 1.000 Einwohner*innen pro Jahr hinzugekommen, in 2015 waren es sogar 1.796.'
                                
            prevyear = r'Im Vergleich zum Vorjahr \marginnote{\emph{höchste Einwohnerzahl seit fast 60 Jahren}}[0cm] steigt die Bevölkerung um 0,74\% (+716 Personen). Damit bewegt sich der Zuwachs auf dem Vorjahresniveau.'
            
            population_development = currentpopulationstring + highest_population + compare_curr_prev_population + average_population + prevyear

            doc.append(NoEscape(population_development))
            doc.append(NoEscape(r'\\'))
            
            doc.append(NoEscape(r'\fcolorbox{black}{PaleAqua}{\parbox{\columnwidth}{\footnotesize{\underline{Hinweis}: Für die Jahre ab 2011 hat das Statistische Amt für Hamburg und Schleswig-Holstein (Statistikamt Nord) auf Grundlage der Ergebnisse des Zensus 2011 eine deutlich unter den bisherigen Ergebnissen liegende Bevölkerungszahl (82.258 zum Stichtag 31.12.2011) förmlich festgesetzt. Das Flensburger Einwohnermelderegister wies im Vergleich eine Einwohner*innenzahl von 89.532 Personen aus. Die Stadt Flensburg hat 2015 gegen die Ergebnisse des Zensus geklagt. Die rechtliche Klärung des Sachverhaltes ist derzeit noch nicht abgeschlossen. Daher beziehen sich alle nachfolgenden Angaben zu den Bevölkerungszahlen weiterhin auf Datenbestände des städtischen Einwohnermelderegisters. Im Gegensatz zu den Zahlen des Statistikamts Nord können die Daten des Melderegisters zudem kleinräumig ausgewertet werden. Des Weiteren wird die Vergleichbarkeit mit den Daten der Sozialatlanten der Vorjahre gewährleistet.}}}'))
            doc.append(NoEscape(r'\\'))
            
            
            
            doc.append(NoEscape(r'\subsubsection{a) kleinräumige Entwicklung}'))
            doc.append(NoEscape(r'\\'))           
            doc.append(NoEscape(r'\marginnote{\emph{Zunahme der Bevölkerung in fast allen Stadtteilen}}[0.25cm]')) 
            doc.append(NoEscape(r'\\'))  
            doc.append(NoEscape(r'Im Vergleich zu 2009 weisen fast alle Stadtteile ein kontinuierliches Wachstum auf (vgl. Tab. \ref{tab:Tabelle_1} und Abb. \ref{fig:Abbildung_2}). Es zeigen sich jedoch große Unterschiede hinsichtlich der Wachstumsraten in den einzelnen Stadtteilen. So ist für die Nordstadt, Mürwik und Tarup ein Zuwachs um mehr als 1.000 Einwohner*innen seit 2009 zu verzeichnen. Auch in den anderen Stadtteilen ist eine allgemeine Zunahme der Bevölkerungszahlen zu konstatieren. Lediglich in Engelsby ist die Einwohner*innenzahl zurückgegangen (-377 ggü. 2009). Im Vergleich zum Vorjahr sind insbesondere Mürwik (+286) und Tarup (+188) gewachsen. Am stärksten geht die Einwohner*innenzahl im Stadtteil Engelsby zurück (-94 ggü. 2018).'))            
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\marginnote{\scriptsize{* Einwohner*innen mit Haupt- oder alleiniger Wohnung. Geringfügige Abweichungen ergeben sich durch nicht zuordenbare Personen.}}[1.25cm]'))
            
            #Districts
            districts = pd.read_csv(datadirectory+"//Population//District.csv", sep = ";",encoding = "ISO-8859-1")
            districts.loc[13] = districts.sum()
            districts.loc[13," "] = "Flensburg"
            #districts.append(df.sum().rename('Flensburg'))
            print(districts)
            createtable(doc,districts)
            
            
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))
            doc.append(NoEscape(r'\\'))

                #example dataframe
            d = {'col1': [1, 2], 'col2': [3, 4], 'col3': [3, 4]}
            df = pd.DataFrame(data=d)
            df
            createtable(doc,df)
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
    year = 2020; 
    
    directory = os.path.dirname(__file__) + '\\Sozialatlas_generated'
    directoryabb1 = directory + '\\ABBILDUNGEN'
    directoryabbbevoelkerung = directoryabb1 + '\\Abbildungen_1_Bevölkerung'
    
    directoryabb1latex = 'ABBILDUNGEN'
    directoryabbbevoelkerunglatex = directoryabb1latex + '//Abbildungen_1_Bevölkerung' 
    
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

        
    # Fancyhdr zum Inhaltsverzeichnis,ListofTables und ListofFigures hinzufügen
    ma.append(NoEscape(r"\addtocontents{toc}{\protect\thispagestyle{fancy}}")) 
    ma.append(NoEscape(r"\addtocontents{lot}{\protect\thispagestyle{fancy}}")) 
    ma.append(NoEscape(r"\addtocontents{lof}{\protect\thispagestyle{fancy}}")) 
    ma.append(NoEscape(r"\tableofcontents")) 
    
    # Seitennummerierung für Inhaltsverzeichnis romanisch
    ma.append(NoEscape(r"\pagenumbering{arabic}")) 
    
#    doc.append(NoEscape(r"\include{0_Einleitung}"))
    einleitung(ma,year)
    
    ma.append(NoEscape(r"\addcontentsline{toc}{chapter}{Zusammenfassung}")) 
    
#    doc.append(NoEscape(r"\include{0_Zusammenfassung}"))
#    doc.append(NoEscape(r"\include{1Bevoelkerung}")) 
    bevoelkerung(ma,year,directoryabbbevoelkerunglatex, directoryabbbevoelkerung,datadirectory)

#    doc.append(NoEscape(r"\include{2_ArbeitsmarktundBeschäftigung}")) 
#    doc.append(NoEscape(r"\include{3_Wohnen}")) 
#    doc.append(NoEscape(r"\include{4_SozialeSicherung}")) 
#    doc.append(NoEscape(r"\include{5_HilfenzurErziehung}")) 
#    doc.append(NoEscape(r"\input{6_ÜbersichtüberdieStadtteile}")) 
    ma.append(NoEscape(r"\newgeometry{left=2cm,right=2cm,top=2cm,bottom=3cm}")) 
#    doc.append(NoEscape(r"\include{6_ÜbersichtüberdieStadtteile}"))
    latexmainend(ma)
    ma.generate_tex()  
    
    # modularization with include didn't work
    # Bevoelkerung document
    #bev = Document('bevölkerung', documentclass= False,fontenc=None,inputenc=None, font_size=None, lmodern=False, textcomp=False,page_numbers=False )
    #bevoelkerung(bev)
    #bev.generate_tex()    
    
    
    #ma.generate_pdf(clean_tex=False, compiler='pdflatex')
    

    # Document with `\maketitle` command activated


    #doc.preamble.append(Command('title', 'Sozialatlas' + str(jahr)))
    #doc.preamble.append(Command('author', 'Anonymous author'))
    #doc.preamble.append(Command('date', NoEscape(r'\today')))
    #doc.append(NoEscape(r'\maketitle'))
