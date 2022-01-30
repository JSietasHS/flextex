# -*- coding: utf-8 -*-
"""
Created on Mon Oct  4 09:44:52 2021

@author: SIETAS
"""
from pylatex.utils import NoEscape
from pylatex import Tabular, MultiColumn, MultiRow
from Flextexdataset import Flextexdataset
import os


#import  altair  as  alt 
#from altair_saver import save

import plotly.express as px


#from selenium import webdriver

#class Content:
class Content:
    def toTex(self, doc):
        """create a Latex document."""
        pass
    
    #def toStreamlit(self, content: str):
     #   """ display content to streamlit"""
     #   pass

    #def save(self, full_file_name: str) -> dict:
     #   """Extract text from the currently loaded file."""
     #   pass

#var data 


#func save

#func toTex

#func print

#compare?

#Helperclass
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
            self.data = MultiColumn(self.get_length(), align='|c|', data=NoEscape(data))
        else:
            self.data = MultiColumn(self.get_length(), align='|c|', data=data)
             
    def checkMulticolInDF(self, row,column):
        return ((row == self.get_row()) and (column == self.get_startcolumn()))
    
#Helperclass Multirow
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
        if color != None:
            self.color = color
        else:
            self.color = "white"
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
        rs = "white"
        if self.color == None:
            rs = "white"
        else:
            rs = self.color
        return rs
    
    def get_bordersamecolor(self):
        return self.bordersamecolor
        
    def set_size_minusone(self):
        self.size = self.size -1

def getColorfromCommandString(command):
    res = None
    if ("{" in command ) and ("}" in command):
        res = command[command.index('{')+1:command.index('}')]
    return res



    
class Table(Content): 
    df = Flextexdataset
    framewidth = 2
    framecolor = None
    fontsize = None
    rowstretch = None
    align = None
    addcoloumnnames = True
    fontweightrowlist = []
    fontweightcoloumnlist = []
    rowcolor = None
    columncolor = None
    label = None 
    multrow = None
    multcolumn = None

    
    def __init__(self, df, addcoloumnnames = True, multcolumn=None, multrow=None, columncolor = None, rowcolor=None, fontsize = 'normalsize', framecolor = "black", framewidth = 2, rowstretch = None, align = None, fontweightcoloumnlist = [], fontweightrowlist = [], label = None):
        self.df = df
        self.framewidth = framewidth
        self.framecolor = framecolor
        self.fontsize = fontsize
        self.rowstretch = rowstretch
        self.align = align
        self.addcoloumnnames = addcoloumnnames
        self.fontweightrowlist = fontweightrowlist
        self.fontweightcoloumnlist = fontweightcoloumnlist
        self.rowcolor = rowcolor
        self.columncolor = columncolor
        self.label = label
        self.multrow = multrow
        self.multcolumn = multcolumn
        
        
    def __getColorfromCommandString(command):
        res = None
        if ("{" in command ) and ("}" in command):
            res = command[command.index('{')+1:command.index('}')]
        return res       
        
        
    def toTex(self,doc):
        dataframe = self.df.getDataset()
        dfcopy = dataframe.copy(deep=True) #copy of the dataframe
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
        
        doc.append(NoEscape(r"\begin{table}[h!]"))  
        
        doc.append(NoEscape("\\setlength{\\arrayrulewidth}{"+ str(self.framewidth) +"pt}"))
        #create table coloumns
        if self.framecolor is not None: 
            doc.append(NoEscape(r'\arrayrulecolor{'+self.framecolor+'}'))
        if self.fontsize is not None:
            doc.append(NoEscape(r'\begin{'+self.fontsize+'}'))
        if self.rowstretch is not None:
            doc.append(NoEscape(r'\renewcommand{\arraystretch}{'+str(self.rowstretch)+'}'))
            
        #align
        while i in range(0,len(dataframe.columns)): 
            if self.align is not None:
                setalign = False
                for a in self.align:
                    if ((a[0] is not None) and (a[1] is not None) and (a[0] == i)):
                        c = c + str(a[1]) + '|'
                        setalign = True
            if not setalign:
                c = c + 'c|' 
            i = i+1 
                   
        table3 = Tabular(c)
        table3.add_hline()
        #add header coloumn of pd.dataframe
        if self.addcoloumnnames:
            coloumnnames = dataframe.columns.values.tolist()
            if (self.fontweightrowlist is not None) and (-1 in self.fontweightrowlist):
                for i in range(0,len(dataframe.columns)):
                    coloumnnames[i] = r"\textbf{"+coloumnnames[i]+"}"
            if self.fontweightcoloumnlist is not None:
                  for i in range(0,len(dataframe.columns)):
                      if i in self.fontweightcoloumnlist:
                          coloumnnames[i] = r"\textbf{"+coloumnnames[i]+"}"
            if self.rowcolor is not None:
                headercolor = ""
                headercolor = getCommandstringFromList(self.columncolor,0,"cellcolor")
                for i in range(0,len(dataframe.columns)):
                    coloumnnames[i] = NoEscape(headercolor + " " + coloumnnames[i])
            if self.columncolor is not None:
                for i in range(0,len(dataframe.columns)):
                    headercolor = ""
                    headercolor = getCommandstringFromList(self.columncolor,i,"cellcolor")
                    coloumnnames[i] = NoEscape(headercolor + " " + coloumnnames[i])
            table3.add_row(coloumnnames)
            table3.add_hline()
            
        #crete hhline
        for cid, c in enumerate(dataframe.columns):
            cellcolor= ""
            columncellcolor = ""
            columncellcolor = getCommandstringFromList(self.columncolor,cid,"cellcolor")    
            #columncolor before rowcolor     
            if columncellcolor == "" :
                cellcolor = rowcellcolor
            else:
                cellcolor = columncellcolor
            if (cellcolor != "") and (cellcolor != None):
                cellcolorstring = getColorfromCommandString(cellcolor)
            else:
                cellcolorstring = "white"          
            #create List for hhline
            hhline.append(Hhlinevalue(0,cellcolorstring))
                
        
        #iterate through rows
        for row in dfcopy.index:
            
            #print(str(row))
            
            #add font weight
            fontweightrow= False
            if row in self.fontweightrowlist:
                fontweightrow = True
            
            rowcellcolor = ""
            #print('row:' + str(row))
            values = []
            #iterate through coloumns
            rowcellcolor = getCommandstringFromList(self.rowcolor,row,"cellcolor")
            for columnind, column in enumerate(dataframe.columns):
                #add fontweight
                fontweightcolumn = False
                if i in self.fontweightcoloumnlist:
                    fontweightcolumn = True
                    
                if fontweightrow or fontweightcolumn:
                    dfcopy.iat[row,columnind] = "\\textbf{" + str(dfcopy.iat[row,columnind]) + "}"
                
                addmcol = None
                addmrow = None
                cellcolor= ""
                columncellcolor = ""
                columncellcolor = getCommandstringFromList(self.columncolor,columnind,"cellcolor")
                
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
                if self.multcolumn is not None: 
                    for mcol in self.multcolumn:
                        data = ''
                        if mcol.checkMulticolInDF(row,columnind):
                            addmcol = mcol
                #check multirow is in cell
                if self.multrow is not None: 
                    for mrow in self.multrow:
                        if mrow.checkMultirowInDF(row,columnind):
                            addmrow = mrow
                #append multicolrow
                if ((addmcol is not None) and (addmrow is not None)):
                    for i in range(0,addmrow.get_length()):
                        for j in range(0,addmcol.get_length()):
                            data = data + str(dfcopy.iat[row+i,columnind+j])
                            #change hhline color and size
                            hhline[columnind+j].set_size(addmrow.get_length()-1)
                            addmrow.get_length()-1-j
                            hhline[columnind+j].set_color(getColorfromCommandString(cellcolor))
                            hhline[columnind+j].set_bordersamecolor(True)
                            if (i > 0)  and (i != addmrow.get_endrow()) :
                                dfcopy.iat[row+i,columnind+j] = "" 
                            else:
                                dfcopy.iat[row+i,columnind+j] = None 
                        hhline[columnind+j].set_bordersamecolor(False)
                        #append more multicoloumns
                        if (i > 0) and (i < (addmrow.get_length()-1)):
                            self.multcolumn.append(Multicol((addmcol.get_row()+i),addmcol.get_startcolumn(), addmcol.get_endcolumn()))
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
                    print("1:" + hhline[columnind].get_color())
                    
                    
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
                if ((hhlinevalue.get_size()) > 0):
                    print(hhlinevalue.get_bordersamecolor())
                    if (hhlinevalue.get_bordersamecolor()):
                        
                        hhlinelatexcommand = hhlinelatexcommand + r">{\arrayrulecolor{" + hhlinevalue.get_color() + r"}}->{\arrayrulecolor{"+ hhlinevalue.get_color() +"}}|"
                    else:
                        print("2:" + hhlinevalue.get_color())
                        hhlinelatexcommand = hhlinelatexcommand + r">{\arrayrulecolor{" + hhlinevalue.get_color() + r"}}->{\arrayrulecolor{"+ self.framecolor +"}}|"
                    hhlinevalue.set_size_minusone()
                    if ((hhlinevalue.get_size()) == 0):
                        hhlinevalue.set_bordersamecolor(False)
                        hhlinevalue.set_color(""+ self.framecolor +"")
                    
                else:
                    hhlinelatexcommand = hhlinelatexcommand + r">{\arrayrulecolor{"+ self.framecolor +"}}-|"
            hhlinelatexcommand = hhlinelatexcommand + r"}"
            table3.append(NoEscape(hhlinelatexcommand))
            hhlinelatexcommand = r"\hhline{|"
    
        doc.append(table3)
        if not (self.label is None):
            doc.append(NoEscape("\label{"+self.label+"}"))
            

            
        if self.fontsize is not None:
            doc.append(NoEscape(r'\end{'+self.fontsize+'}'))
            
            
        doc.append(NoEscape(r"\end{table}"))     
        return doc  


class Text(Content):
    text =str

    def __init__(self,content):
        self.text = content
    
    def setContent(self, content : str):
        self.text = content
    
    def toTex(self, doc):
        doc.append(NoEscape(r""+self.text))
        return doc
    

        
        
#class Dialog:
    
    
    
    

    
    
class Diagram(Content):
    df = Flextexdataset
    x = str
    y = str
    color_discrete_sequence = str
    title =str
    filepath = str
    label = str
    opacity = float
    range_x = None
    range_y = None
    width = float
    height = float
    caption = str
    
    
    def __init__(self, df , x:str, y:str, title:str, color_discrete_sequence: str, text:str, opacity:float, range_x , range_y, width:int, height:int, caption:str = "", label: str = "" ):
        self.df = df
        self.x = x
        self.y = y
        self.color_discrete_sequence = color_discrete_sequence
        self.title = title
        self.caption = caption
        self.label = label
        self.opacity = opacity
        self.range_x = range_x
        self.range_y = range_y
        self.width = width
        self.height = height

        
    def getDiagram(self):
        return px.bar(self.df.getDataset(), x=self.x , y=self.y, title=self.title, color_discrete_sequence = self.color_discrete_sequence,  opacity = self.opacity, range_x = self.range_x , range_y = self.range_y , width = self.width, height = self.height)
    
    def setFilePath(self,filepath:str):
        self.filepath = filepath
        
    def getFilePath(self) -> str:
        return self.filepath
        
        
    def __savefig__(self):
        fig = self.getDiagram()
        f =  self.getFilePath()
        fig.write_image(f)
        
    def toTex(self,doc):
        self.__savefig__()
        doc.append(NoEscape(r'\begin{figure}[H]'))
        doc.append(NoEscape(r' \includegraphics[width=\textwidth]{'+os.path.basename(self.getFilePath())+ "}" ))
        doc.append(NoEscape(r' \caption{'+self.caption+'}'))
        if (self.label != ""):
            doc.append(NoEscape(r' \label{fig:'+ self.label+'}')) 
        doc.append(NoEscape(r'\end{figure}'))
        return doc
    
#    
#class Diagram(Content):
#    """A class that saves Values to define a Multicoloumn."""
#    
#    x = ""
#    y = ""
#    xtype = "Q"
#    ytype = "Q"
#    xlabel = ""
#    ylabel = ""
#    xtitle = ""
#    ytitle = ""
#    xscalemin = 0
#    yscalemin = 0
#    xscalemax = 0
#    yscalemax = 0     
#    align = "center"
#    baseline = "bottom"
#    dx = 0
#    dy = 0
#    width = 500
#    height = 300
#    color = '#4F81BD'
#    size = 20
#    textlabe = None
#    ds = None
#    caption = ""
#           
#
#    def __init__(self, ds, x, y, xtype="Q", ytype="Q", xlabel=False, ylabel=False, xtitle="", ytitle="", xscalemin=0, yscalemin=0, xscalemax=0, yscalemax=0, align="center", baseline="bottom", dx=0, dy=0, width=500, height=300, color='#4F81BD', size=20, textlabel=None, caption="EmptyCaption"):
#        self.x = x#+":"+xtype
#        self.y = y#+":"+ytype
#        self.xtype = xtype
#        self.ytype = ytype
#        self.xlabel = xlabel
#        self.ylabel = ylabel
#        self.xtitle = xtitle
#        self.xscalemin = xscalemin
#        self.yscalemin = yscalemin
#        self.xscalemax = xscalemax
#        self.yscalemax = yscalemax
#        self.align = align
#        self.baseline = baseline
#        self.dx = dx
#        self.dy = dy
#       self.width = width
#        self.color = color
#        self.size = size
#        self.textlabel = textlabel
#        self.ds = ds
#        self.caption = caption
#        
#    def toTex(self, doc):
#        
#        chrome_driver_path= '\\Obelix\home$\SIETAS\Documents\Code\sozialatlasgenerator\\chromedriver.exe'
#   
#        #set chromedriver Path
#        try:
#            driver = webdriver.Chrome(r'C:\Program Files\Google\Chromedriver\chromedriver.exe')
#            driver.close()
#        except:
#            driver = webdriver.Chrome() 
#        
#        print("label1")
#        print(str(self.xlabel))
#        print("labelbool")
#        print(str(self.xtitle))
#        ################################################Population Chart####################################
#        bars = alt.Chart(self.ds).mark_bar(size = self.size, color=self.color).encode(
#                alt.X(self.x,
#                      axis=alt.Axis(labels=self.xlabel, title=self.xtitle),
#                     #scale=alt.Scale(domain=(year-11, year-1))
#                     ),
#               alt.Y(self.y,
#                    axis=alt.Axis(labels=self.ylabel, title=self.ytitle),
#                    #scale=alt.Scale(domain=(round(minpopulation*0.97), round(maxpopulation*1.03))),
#                    ),
#                )
#        if self.textlabel is not None: 
#            text = bars.mark_text(
#                align=self.align,
#                baseline=self.baseline,
#                dx = self.dx,
#                dy = self.dy
#            ).encode(
#                text= self.textlabel
#            )
#        
#            d =(bars + text).properties(width=self.width)
#        else:
#            d =(bars).properties(width=self.width)
#        
#        d.properties(width=self.width).save(r'C:\Temp\newdoc\abb1.png')
#                
#    
#        ################################################Population Chart####################################
#        
#        doc.append(NoEscape(r'\begin{figure}[H]'))
#       doc.append(NoEscape(r' \includegraphics[width=\textwidth]{'+r'C:\Temp\newdoc\abb1.png}'))
#       doc.append(NoEscape(r' \caption{\textbf{Bevölkerungsentwicklung ... bis ... (ohne Berücksichtigung Zensus 2011).}}'))
#        doc.append(NoEscape(r' \label{fig:Abbildung_1}'))
#        doc.append(NoEscape(r'\end{figure}'))
#    
#    #def toStreamlit(self, doc):
#    #    """Load in the file for extracting text."""
#    #    pass
       
        