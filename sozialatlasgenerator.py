from pylatex import Document, Section, Subsection, Command, Package, \
    MultiRow, Tabular, MultiColumn
from pylatex.utils import italic, NoEscape
from pylatex.section import Chapter
import  altair  as  alt 
import numpy as np
import pandas as pd
import os
from vega_datasets import data

from altair_saver import save

from selenium import webdriver


#row - The row in which the multi col start 
#start coloumn - the start coloumn to connect the row
#end coloumn - the end coloumn to connect the row
class Multicol():
    """A class that saves Values to define a Multicoloumn."""
    
    col = 0
    startcoloumn = 0
    endcoloumn = 0

    def __init__(self, col, startcoloumn, endcoloumn):
        self.col = col
        self.startcoloumn = startcoloumn
        self.endcoloumn = endcoloumn
        
    def get_col(self):
        return self.col
    
    def get_startcoloumn(self):
        return self.startcoloumn
    
    def get_endcoloumn(self):
        return self.endcoloumn

#col - The coloumn in which the multi row start 
#startrow - the start coloumn to connect the coloumns
#endrow - the end coloumn to connect the coloumns
class Multirow():
    """A class that saves Values to define a Multicoloumn."""
    
    row = 0
    startrow = 0
    endrow = 0

    def __init__(self, row, startrow, endcrow):
        self.row = row
        self.startrow = startrow
        self.endcrow = endcrow
        
    def get_row(self):
        return self.row
    
    def get_startrow(self):
        return self.startrow
    
    def get_endrow(self):
        return self.endcrow
    
#Create a Table out of a pandas Dataframe 
#input  document
#       dataframe: a pandas dataframe
#       coloumnnames
#       headline colors
#       coloumncolor
#       coloumwithcolor list of integer #Paint the number of the coloumn in the 
#                               coloumncolor default value all coloumns
#       multirows # List of Triple [Rownumber,start coloumn, end coloumn]
#                   rownumber start at 0 the row of the table
#                   coloumnnumber start at 0 the start coloumn to connect coloumns
#       multicoloumns # # List of MultiCol elements to add multicoloumns          
def createtable(doc, df, multcoloumn = None):
    df = df.copy(deep=False) #create a copy of the dataframe
    c = '|' # create latex table coloumn string
    inverti = 0 # Invert counter to add data from a multicoloumn
    multicoloumnfound = False # termination condition
    multirowfound = False # termination condition
    mrowstart = None
    k = 0 # multicoloumncounter
    j = 0 # rowcounter
    i = 0 # coloumncounter
    l = 0
    m = 0
    datamultcol = '' # data of a multicoloumn
    datamultrow = '' # data of a multicoloumn
    #add package
    
    #test
    multcoloumn = None
    #multcoloumn = [Multicol(0,1,2)]
    multrow = None
    multrow = [Multirow(0,0,1)]
    
    
    doc.preamble.append(Package("multirow"))
     
    #Create coloumns      
    while i in range(0,len(df.columns)):
        c = c + 'c|'
        i = i+1
    
    table3 = Tabular(c)
    #Iterate through rows
    for i in df.index:
        values = []
        j = 0
        mclmlength = 0
        multicoloumnfound = True
        multirowfound = True
        datamultcol = ''
        datamultrow = ''
        k = 0
        #Multirow
        mrowstart = None
        
        l = 0
        m = 0
        #check if there is a multicoloumn in the coloumnv
        while ((l < len(multrow)) and (multirowfound)):
            mrow = multrow[l]
            l = l+1;
            #if there is a multirow set the length, the start coloumn and a boolean
            if (mrow.get_row() == i):
               mrowlength = mrow.get_endrow()+1 - mrow.get_startrow()
               mrowstart = mrow.get_startrow()
               multirowfound = False
               
        #Multicloumn
        #iterate through coloumns
        for coloumn in df.columns:
            k = 0
            #check if there is a multicoloumn in the coloumn # todo so geht nur eine multrow
            if multcoloumn is not None:
                while ((k < len(multcoloumn)) and (multicoloumnfound)):
                    mclm = multcoloumn[k]
                    k = k+1;
                    #if there is a multicoloumn set the length and a invertcounter
                    if ((mclm.get_col() == i) and (mclm.get_startcoloumn() == j)):
                       mclmlength = mclm.get_endcoloumn()+1 - mclm.get_startcoloumn()
                       inverti = mclmlength
                       multicoloumnfound = False
            
            #Mulirow found and multicoloumn not found
            if (not(multirowfound)) and (multicoloumnfound) and (mrowstart == j):
                #add data
                print("test")
                for m in range(0,mrowlength):
                    datamultrow = datamultrow + str(df[coloumn][i+m])
                    df.loc[j+m,coloumn] = ''
                values.append(MultiRow(mrowlength, data=datamultrow))
                multirowfound = True
                mrowlength = 0
            else:     
                #Multirowend
                #connect the rows for a multicoloumn or write the value to the table
                if inverti > 0:
                    inverti = inverti - 1
                    datamultcol = datamultcol + str((df[coloumn][i]))
                    if inverti == 0:
                        values.append(MultiColumn(mclmlength, align='|c|', data=datamultcol))
                        datamultcol = ''
                        mclmlength = 0
                else:
                    values.append(df[coloumn][i])
                    print(str(values))
                    
            j=j+1
        print('Data:' + str(values))
        table3.add_row(values)
        table3.add_hline()

    
    #table3.add_row((MultiColumn(2, align='|c|',
    #                           data=MultiRow(2, data='multi-col-row')), 'X'))
    #table3.add_row((MultiColumn(2, align='|c|', data=''), 'X'))

          
    doc.append(table3)

def latexmainstart(doc):
    
    doc.preamble.append(Package("babel", "ngerman"))
    doc.preamble.append(Package("helvet"))
    doc.preamble.append(Package("natbib"))
    doc.preamble.append(Package("graphicx"))
    doc.preamble.append(Package("float"))
    doc.preamble.append(NoEscape(r"\counterwithout{figure}{chapter}"))
    doc.preamble.append(NoEscape(r"\counterwithout{table}{chapter}"))
    doc.preamble.append(Package("caption", "justification=justified,singlelinecheck=false"))
    doc.preamble.append(NoEscape(r"\renewcommand*{\thefigure}{\textbf{\arabic{figure}}}"))
    doc.preamble.append(NoEscape(r"\renewcommand*{\thetable}{\textbf{\arabic{table}}}"))
    doc.preamble.append(Package("adjustbox", "export"))
    doc.preamble.append(Package("xcolor"))
    doc.preamble.append(Package("fancyhdr"))
    doc.preamble.append(Package("marginnote"))
    doc.preamble.append(Package("multirow"))
    doc.preamble.append(NoEscape(r"\let\oldfbox\fbox"))
    doc.preamble.append(NoEscape(r"\renewcommand\fbox[1]{\savenotes\oldfbox{#1}\spewnotes}"))
    doc.preamble.append(Package("geometry", "outer=6cm,inner=1.5cm,top=5cm, bottom=3cm"))
    doc.preamble.append(Package("footnote"))
    doc.preamble.append(Package("pdfpages"))
    doc.preamble.append(Package("tocloft"))
    doc.preamble.append(NoEscape(r"\renewcommand{\cftfigpresnum}{\textbf{Abb. }}"))
    doc.preamble.append(NoEscape(r"\renewcommand{\cfttabpresnum}{\textbf{Tab. }}"))
    doc.preamble.append(NoEscape(r"\renewcommand{\cftfigaftersnum}{:}"))
    doc.preamble.append(NoEscape(r"\renewcommand{\cfttabaftersnum}{:}"))
    doc.preamble.append(NoEscape(r"\setlength{\cftfignumwidth}{2cm}"))
    doc.preamble.append(NoEscape(r"\setlength{\cfttabnumwidth}{2cm}"))
    doc.preamble.append(NoEscape(r"\setlength{\cftfigindent}{0cm}"))
    doc.preamble.append(NoEscape(r"\setlength{\cfttabindent}{0cm}"))
    doc.preamble.append(NoEscape(r"\pagestyle{fancy}"))
    doc.preamble.append(NoEscape(r"\addtolength{\headwidth}{\marginparsep}"))
    doc.preamble.append(NoEscape(r"\addtolength{\headwidth}{\marginparwidth}"))
    doc.preamble.append(NoEscape(r"\fancyhead[RE,LO]{\textbf{Sozialatlas 2019}}"))
    doc.preamble.append(NoEscape(r"\fancyhead[LE,RO]{\includegraphics[scale=0.5]{ABBILDUNGEN/Abbildungen/Logo Stadt Flensburg.png}}"))
    doc.preamble.append(NoEscape(r"\fancyfoot[CE,CO]{}"))
    doc.preamble.append(NoEscape(r"\fancyfoot[LE,RO]{\thepage}"))
    doc.preamble.append(NoEscape(r"\renewcommand{\footrulewidth}{1pt}"))
    doc.preamble.append(NoEscape(r"\marginparwidth=5cm"))
    doc.preamble.append(NoEscape(r"\setlength{\parindent}{0pt}"))
#    doc.append(NoEscape(r"\includepdf[pages={1}]{ABBILDUNGEN/Deckblatt Sozialatlas 2019.pdf}"))
    doc.append(NoEscape(r"\renewcommand{\figurename}{\textbf{Abb.}}"))
    doc.append(NoEscape(r"\renewcommand{\tablename}{\textbf{Tab.}}")) 
    doc.append(NoEscape(r"\thispagestyle{empty}")) 
#    doc.append(NoEscape(r"\include{0_Impressum}")) 
    doc.append(NoEscape(r"\newpage")) 
    doc.append(NoEscape(r"\pagestyle{fancy}")) 
    doc.append(NoEscape(r"\pagestyle{fancy}")) 
#    doc.append(NoEscape(r"\include{0_Vorwort}")) 
#    doc.append(NoEscape(r"\include{0_Zusammenfassung}"))
#    modularization with include didn't work 
#    doc.append(NoEscape(r"\include{1Bevoelkerung}")) 

#    doc.append(NoEscape(r"\include{2_ArbeitsmarktundBeschäftigung}")) 
#    doc.append(NoEscape(r"\include{3_Wohnen}")) 
#    doc.append(NoEscape(r"\include{4_SozialeSicherung}")) 
#    doc.append(NoEscape(r"\include{5_HilfenzurErziehung}")) 
    doc.append(NoEscape(r"\newgeometry{outer=2cm,inner=1.5cm}")) 
#    doc.append(NoEscape(r"\include{6_ÜbersichtüberdieStadtteile}")) 

def latexmainend(doc):
    
    doc.append(NoEscape(r"\addcontentsline{toc}{chapter}{Übersicht über die Stadtteile}")) 
    doc.append(NoEscape(r"\begingroup")) 
    doc.append(NoEscape(r"	\renewcommand*{\addvspace}[1]{}")) 
    doc.append(NoEscape(r"	\phantomsection")) 
    doc.append(NoEscape(r"	\addcontentsline{toc}{chapter}{\listfigurename}")) 
    doc.append(NoEscape(r"	\listoffigures")) 
    doc.append(NoEscape(r"	\newpage")) 
    doc.append(NoEscape(r"	\phantomsection")) 
    doc.append(NoEscape(r"	\addcontentsline{toc}{chapter}{\listtablename}")) 
    doc.append(NoEscape(r"	\listoftables")) 
    doc.append(NoEscape(r"\endgroup"))    
    
        

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
    ma = Document('main', documentclass="scrreprt",document_options="a4paper, 10.5pt, twoside", lmodern=False, textcomp=False,page_numbers=False )
    latexmainstart(ma)
    bevoelkerung(ma,year,populationofelevenyears)
    latexmainend(ma)
    ma.generate_tex()  
    
    # modularization with include didn't work
    # Bevoelkerung document
    #bev = Document('bevölkerung', documentclass= False,fontenc=None,inputenc=None, font_size=None, lmodern=False, textcomp=False,page_numbers=False )
    #bevoelkerung(bev)
    #bev.generate_tex()    
    
    
    #doc.generate_pdf(clean_tex=False)
    

    # Document with `\maketitle` command activated


    #doc.preamble.append(Command('title', 'Sozialatlas' + str(jahr)))
    #doc.preamble.append(Command('author', 'Anonymous author'))
    #doc.preamble.append(Command('date', NoEscape(r'\today')))
    #doc.append(NoEscape(r'\maketitle'))
