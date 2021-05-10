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
        print('Data:' + str(values))
        table3.add_row(values)
        table3.add_hline()
                
    print(dfcopy)
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

#    doc.preamble.append(NoEscape(r"\ExplSyntaxOn"))
#    doc.preamble.append(NoEscape(r"   \clist_map_inline:nn"))
#    doc.preamble.append(NoEscape(r"     {figure,table}"))
#    doc.preamble.append(NoEscape(r"     {\DeclareTOCStyleEntry["))
#    doc.preamble.append(NoEscape(r"       pagenumberformat=\bfseries"))
#    doc.preamble.append(NoEscape(r"       entryformat=\bfseries"))
#    doc.preamble.append(NoEscape(r"       ]{section}{#1}"))
#    doc.preamble.append(NoEscape(r"     }"))
#    doc.preamble.append(NoEscape(r" \ExplSyntaxOff"))


 
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
    doc.preamble.append(NoEscape(r"\fancyhead[LE,RO]{\includegraphics[scale=0.5]{ABBILDUNGEN/Abbildungen/Logo Stadt Flensburg.png}}"))
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

    doc.append(NoEscape(r"\thispagestyle{empty}")) 
#    doc.append(NoEscape(r"\include{0_Impressum}")) 
    doc.append(NoEscape(r"\thispagestyle{empty}")) 
#    doc.append(NoEscape(r"\include{0_Einleitung}")) 

#%%%% Fancyhdr zum Inhaltsverzeichnis,ListofTables und ListofFigures hinzufügen
    doc.append(NoEscape(r"\addtocontents{toc}{\protect\thispagestyle{fancy}}")) 
    doc.append(NoEscape(r"\addtocontents{lot}{\protect\thispagestyle{fancy}}")) 
    doc.append(NoEscape(r"\addtocontents{lof}{\protect\thispagestyle{fancy}}")) 
    doc.append(NoEscape(r"\tableofcontents")) 
    
    doc.append(NoEscape(r"\tableofcontents")) 
#%%%%% Seitennummerierung für Inhaltsverzeichnis romanisch
    doc.append(NoEscape(r"\pagenumbering{arabic}")) 
    
#    doc.append(NoEscape(r"\include{0_Einleitung}"))
    
    doc.append(NoEscape(r"\addcontentsline{toc}{chapter}{Zusammenfassung}")) 
    
#    doc.append(NoEscape(r"\include{0_Zusammenfassung}"))
#    doc.append(NoEscape(r"\include{1Bevoelkerung}")) 

#    doc.append(NoEscape(r"\include{2_ArbeitsmarktundBeschäftigung}")) 
#    doc.append(NoEscape(r"\include{3_Wohnen}")) 
#    doc.append(NoEscape(r"\include{4_SozialeSicherung}")) 
#    doc.append(NoEscape(r"\include{5_HilfenzurErziehung}")) 
#    doc.append(NoEscape(r"\input{6_ÜbersichtüberdieStadtteile}")) 
    doc.append(NoEscape(r"\newgeometry{left=2cm,right=2cm,top=2cm,bottom=3cm}")) 
#    doc.append(NoEscape(r"\include{6_ÜbersichtüberdieStadtteile}")) 
    doc.append(NoEscape(r"\listoffigures")) 
    doc.append(NoEscape(r"\addcontentsline{toc}{chapter}{\listfigurename}")) 
    doc.append(NoEscape(r"\thispagestyle{empty}")) 
    doc.append(NoEscape(r"\newpage")) 
    doc.append(NoEscape(r"\listoftables")) 
    doc.append(NoEscape(r"\addcontentsline{toc}{chapter}{\listtablename}")) 
    doc.append(NoEscape(r"\thispagestyle{empty}"))         










def bevoelkerung(doc, year, populationofelevenyears):
    """Add a section, a subsection and some text to the document.

    :param doc: the document
    :type doc: :class:`pylatex.document.Document` instance
    """
    with doc.create(Chapter('Bevölkerung', label="1Bevoelkerung")):
        with doc.create(Section('Bevölkerungsentwicklung')):
            doc.append(NoEscape(r'\marginnote{\emph{Einflüsse auf die ' +
                                'Bevölkerungsentwicklung}}[0.2cm]'))
            doc.append(NoEscape(r'\fbox{\parbox{\columnwidth}{\emph{Die ' +
                                'Bevölkerungsentwicklung ergibt sich aus der ' +
                                'Differenz zwischen Geburtenrate und Sterberate '+ 
                                'in Verbindung mit dem Wanderungssaldo. Dieser ' +
                                'wiederum wird von verschiedenen Faktoren ' +
                                'beeinflusst: von globalen politischen ' +
                                'Entwicklungen, Tendenzen auf dem Arbeitsmarkt ' + 
                                '(z.B. Anzahl der offenen und vermittelbaren ' +
                                ' Stellen), dem Wohnraumangebot (z.B. Mietpreise, ' + 
                                'freie Wohnkapazitäten, Wohnraumqualität), durch ' +
                                'die Bildungsinfrastruktur (z.B. Angebot an ' +
                                'Kindertagesstätten und Schulen bzw. Hochschulen), ' + 
                                'das Angebot an beruflichen Ausbildungen sowie ' +
                                'durch persönliche oder familiäre Entscheidungen ' + 
                                'über den Hauptwohnsitz.}}} '))
            doc.append(NoEscape(r'\begin{figure}[htbp]'))
            s = r'\includegraphics[width=\textwidth]{ABBILDUNGEN/Abbildungen/Bevölkerungsentwicklung '+ str(year-10) +' bis '+ str(year) +'.png}'
            doc.append(NoEscape(s))
            doc.append(NoEscape(r'\caption{\textbf{Bevölkerungsentwicklung '+ str(year-10) +' bis '+ str(year) +' (ohne Berücksichtigung Zensus 2011).}}'))
            doc.append(NoEscape(r' \label{fig:Abbildung_1}'))
            doc.append(NoEscape(r'\end{figure}'))
            #Was war die Einwohnerzahl vor 40 Jahren?
            doc.append(NoEscape(r'\marginnote{\emph{Bevölkerungszunahme seit dem Jahr 2010, aktueller Stand: 96.204}}[0cm]'))
            doc.append(NoEscape(r'Im Rückblick zeigt sich, dass die Einwohnerzahl im Vergleich zum Jahr 2008 um 7.071 Personen angestiegen ist (vgl. Abb. 1). Von 2010 bis 2014 wuchs die Stadt jährlich um ca. 500 Personen. Zwischen 2015 und 2017 sind mehr als 1.000 EinwohnerInnen je Jahr hinzugekommen, in 2015 sogar 1.796. Im Jahr 2018 stieg \marginnote{\emph{höchste Einwohnerzahl seit über 40 Jahren}}[0cm] die Bevölkerungszahl um 735 Personen, also nicht mehr ganz so stark wie zuvor. Mit einer Einwohnerzahl von 96.204 verfügt Flensburg über den höchsten Bevölkerungsstand seit über 40 Jahren.'))
            doc.append(NoEscape(r'\fcolorbox{black}{cyan}{\parbox{\columnwidth}{\footnotesize{\underline{Hinweis}: Für die Jahre ab 2011 hat das Statistikamt Nord auf Grundlage der Ergebnisse des Zensus 2011 eine deutlich unter den bisherigen Ergebnissen liegende Bevölkerungszahl (82.258 zum Stichtag 31.12.2011) förmlich festgesetzt. Das Flensburger Einwohnermeldere-gister wies im Vergleich eine Einwohnerzahl von 89.532 Personen aus. Alle nachfolgenden Angaben zu den Bevölkerungszahlen beziehen sich aber weiterhin auf Datenbestände des städtischen Einwohnermelderegisters.}}}'))
            doc.append(NoEscape(r'\textbf{a) kleinräumige Entwicklung}'))
            doc.append(NoEscape(r'\marginnote{\emph{Zunahme der Bevölkerung in fast allen Stadtteilen}}[0cm]'))
            doc.append(NoEscape(r'Die Bevölkerungsentwicklung verläuft im Zehnjahresvergleich in fast allen Stadtteilen positiv (vgl. Tab 1 und Abb. 2), wenn sich auch große Unterschiede hinsichtlich der Intensität des Wachstums zeigen. In der Nordstadt, Weiche und Tarup sind deutlich mehr Personen mit Hauptwohnsitz gemeldet als vor zehn Jahren (Zuwachs um jeweils mehr als 1.000 Einwohner). Mit Ausnahme von Engelsby hat die Bevölkerung auch in allen anderen Stadtteilen zugenommen. Im Vergleich zum Vorjahr hat der Friesische Berg 160 EinwohnerInnen verloren, in Engelsby gab es hingegen einen Zuwachs um 31 Personen.'))
            
            doc.append(NoEscape(r'\begin{table}[htbp]'))
            doc.append(NoEscape(r'    \caption{\textbf{EinwohnerInnen in den Stadtteilen 2008 bis 2018*.}}'))
            doc.append(NoEscape(r'    \includegraphics[width=\textwidth]{TABELLEN/Tabellen/EinwohnerInnen in den Stadtteilen 2008 bis 2018.png}'))
            doc.append(NoEscape(r'    \label{tab:Tabelle_1}'))
            doc.append(NoEscape(r'\end{table}'))
            
                #example dataframe
            d = {'col1': [1, 2], 'col2': [3, 4], 'col3': [3, 4]}
            df = pd.DataFrame(data=d)
            df
            createtable(doc,df)
            print(df)
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
    year = 2018; 
    populationofelevenyears =  [89133, 88812, 88959, 89532, 90179, 90641, 91316, 93112, 94227, 95469, 96204]
    minpoulation = min(populationofelevenyears)
    maxpoulation = max(populationofelevenyears)
    
    directory = os.path.dirname(__file__) + '\\Sozialatlas_generated'
    directoryabb1 = directory + '\\ABBILDUNGEN'
    directoryabb2 = directoryabb1 + '\\Abbildungen'
    
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
        os.stat(directoryabb2)
    except:
        os.mkdir(directoryabb2) 
        
    
    #output Path of the tex data
    os.chdir(path = directory )
    
    ######################################Create Diagrams
    stringpath = directoryabb2 + '\\Bevölkerungsentwicklung '+ str(year-10) +' bis '+ str(year) +'.png'
    print(stringpath)
    #Create Population Chart of the last 11 years
    x=list(range(year-10, year+1, 1))
    population  =  pd . DataFrame ({ 
    'years' :  x, 
    'population' :  populationofelevenyears,
    })
    populationchart = alt.Chart(population).mark_bar(size = 20, color='blue').encode(
    alt.X('years',
          axis=alt.Axis(labels=False),
          ),
    alt.Y('population',
        axis=alt.Axis(labels=False),
        scale=alt.Scale(domain=(round(minpoulation*0.97), round(maxpoulation*1.03)))
        ),
    )
    populationchart.save(directoryabb2 + '\\Bevölkerungsentwicklung '+ str(year-10) +' bis '+ str(year) +'.png')
    
    
    
    
    # Main document
    ma = Document('main', documentclass="scrreprt",document_options="a4paper, 10.5pt, twoside, listof=entryprefix", lmodern=False, textcomp=False,page_numbers=False )
    latexmainstart(ma,year)
    bevoelkerung(ma,year,populationofelevenyears)
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
