# -*- coding: utf-8 -*-
"""
Created on Tue Jan 25 11:49:01 2022

@author: SIETAS
"""
from pylatex import Document, Package
from Content import Content
from Flextexdataset import Flextexdataset
from pylatex.utils import NoEscape

class FlexTexDocument():
    content = []
    datasets = []
    dirfolder = str
    filename = str
    documentclass = str
    papersizeformat = str
    singleordoubleside = str
    fontsize = str
    listofentrypraefix = str
    lmodern  = bool
    textcomp = bool
    pagenumbers = bool
    dirname = str
    filename = str

    
    """A class that saves Values to define a Multicoloumn."""
    
    def __init__(self, dirname : str, filename : str, documentclass:str = "scrreprt", fontsize : str = "10.5pt", papersizeformat: str = "a4paper" , singleordoubleside:str = "twoside" , listofentrypraefix: bool = True, lmodern: bool = False, textcomp: bool = False, pagenumbers:bool =False):
       
        self.dirname = dirname
        self.filename = filename
        self.fontsize = fontsize
        self.papersizeformat = papersizeformat
        self.singleordoubleside = singleordoubleside
        self.listofentrypraefix = listofentrypraefix
        self.documentclass = documentclass
        self.lmodern = lmodern
        self.textcomp = textcomp
        self.pagenumbers = pagenumbers
        
         
    def setDirname(self,dirname : str):
        self.dirname = dirname
        
    def getDirname(self) -> str:
        return self.dirname
    
    def setFilename(self,filename : str):
        self.filename = filename
        
    def getFilename(self) -> str:
        return self.filename
    
    def getfullPath(self) -> str:
        return (self.dirname + "/" + self.filename)
    
    def setDocumentclass(self,documentclass):
        self.documentclass = documentclass
        
    def getDocumentclass(self) -> str:
        return self.documentclass
    
    def setFontsize(self,fontsize):
        self.fontsize = fontsize
        
    def getFontsize(self) -> str:
        return self.fontsize
    
    def setPapersizeformat(self,papersizeformat):
        self.papersizeformat = papersizeformat
        
    def getPapersizeformat(self) -> str:
        return self.papersizeformat
    
    def setSingleordoubleside(self,singleordoubleside):
        self.singleordoubleside = singleordoubleside
        
    def getSingleordoubleside(self) -> str:
        return self.singleordoubleside
    
    def setListofentrypraefix(self,listofentrypraefix):
        self.listofentrypraefix = listofentrypraefix
        
    def getListofentrypraefix(self) -> bool:
        return self.listofentrypraefix
        
    def getDocumentOptions(self) -> str:
        #document_options = "a4paper, 10pt, twoside, listof=entryprefix"
        document_options = self.papersizeformat + ', ' + self.fontsize + ', '  + self.singleordoubleside 
        if self.listofentrypraefix: 
            document_options= document_options + ', '+ "listof=entryprefix"
        return document_options
    
    def setLmodern(self,lmodern : bool):
        self.lmodern = lmodern
        
    def getLmodern(self) -> bool:
        return self.lmodern
         
    def setTextcomp(self,textcomp : bool):
        self.textcomp = textcomp
        
    def getTextcomp(self) -> bool:
        return self.textcomp
    
    def setPagenumbers(self,pagenumbers : bool):
        self.pagenumbers = pagenumbers
        
    def getPagenumbers(self) -> bool:
        return self.pagenumbers
    
    def appendContent(self,  content: Content):
        self.content.append(content)
        
    def getContent(self) -> content:
        return self.content
    
    def appendDataset(self,  dataset: Flextexdataset):
        self.datasets.append(dataset)
        
    def getDatasets(self) -> datasets:
        return self.datasets
    
    def getFlextexdatasetByIndex(self,index) -> Flextexdataset:
        return self.datasets[index].getFlextexdataset()
    
    def getDatasetnames(self) -> list:
        res = []
        for ds in self.datasets:
            res.append(ds.getDatasetname())
        return res
    
    def updateAllDatasets(self):
        i = 0
        while i < len(self.datasets):
            self.updateDatasetByIndex(i)
            i = i + 1
    
    def updateDatasetByIndex(self,index):
        self.datasets[index].loadData()
        
    def getDatasetByIndex(self,index):
        return self.datasets[index].getDataset()
        
    def generateFlexTex(self):
        doc = Document(self.getfullPath(), documentclass=self.documentclass,document_options=self.getDocumentOptions(), lmodern=self.lmodern, textcomp=self.lmodern,page_numbers=self.pagenumbers )
        
        #Die Packete müssen standardmßig rein bzw eigentlich nicht aber das kann ich noch nicht abfragen
        #\usepackage{multirow}%
        #\usepackage{xcolor}%
        #\usepackage{hhline}%
        #\usepackage{colortbl}%
        #\usepackage{float}
        #\usepackage{graphicx}
        doc.preamble.append(Package("multirow"))
        doc.preamble.append(Package("xcolor"))
        doc.preamble.append(Package("hhline"))
        doc.preamble.append(Package("colortbl"))
        doc.preamble.append(Package("float"))
        doc.preamble.append(Package("graphicx"))
        
        counter = 0
        for c in self.content :
            counter = counter + 1
            if ((type(c).__name__) == "Diagram"):
                c.setFilePath(self.getDirname() + "/" + "Diagram " + str(counter) + ".png" )
    
                
            doc = c.toTex(doc)
            
        
            doc.generate_tex() 
        
        