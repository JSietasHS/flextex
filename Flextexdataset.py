# -*- coding: utf-8 -*-
"""
Created on Tue Jan 25 16:47:07 2022

@author: SIETAS
"""
from sqlalchemy import create_engine
import pandas as pd
import os


class Flextexdataset:
    
    def getDataset(self):
        """Load in the file for extracting text."""
        pass
        
    def loadData(self):
        """Load in the file for extracting text."""
        pass
    
    def getDatasetname(self):
        """get the name of a dataset. At the CSV File its the filename"""
        pass
    
    
class Csv(Flextexdataset):
    df =  None
    sep = str
    data_file = str

    def __init__(self, data_file:str, sep:str):
        self.data_file = data_file
        self.sep = sep
        self.loadData()
    
    def loadData(self):
        if os.path.isfile(self.data_file):
            self.df = pd.read_csv(self.data_file, sep = self.sep )

    def getDataset(self):
        return self.df
    
    def getDatasetname(self) -> str:
        return os.path.basename(self.data_file)
    
    def getFlextexdataset(self):
        return self
    
    
        
class MariaDB(Flextexdataset):
    hostname = str
    port = str
    database = str
    user = str
    password = str
    df = None
    query = str
    datasetname = str

    
    def __init__(self, hostname:str, database:str, user:str, password:str, query:str, datasetname:str):
        
        print(hostname)
        print(user)
        
        self.hostname = hostname
        self.database = database
        self.user = user
        self.password = password
        self.query = query
        self.datasetname = datasetname
        self.loadData()
    
    def loadData(self):
        
        engine = create_engine("mysql+pymysql://"+self.user+":"+self.password+"@"+self.hostname+"/"+self.database+"?charset=utf8mb4")
        self.df = pd.read_sql(self.query, engine)
        
    def getDataset(self):
        return self.df
    
    def getDatasetname(self) -> str:
        return self.datasetname
    
    def getFlextexdataset(self):
        return self
    
    