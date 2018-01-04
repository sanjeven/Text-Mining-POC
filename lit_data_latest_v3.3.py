import os
import time
import sys, getopt
import csv
import pandas
import urllib
import re
import optparse

from cgi import logfile
from itertools import izip
from _ast import Div
from win32com.client import Dispatch
from bs4 import BeautifulSoup 


_START_YEAR = 0
_END_YEAR = 0
_FILE_NAME= os.path.dirname(os.path.abspath(__file__))+'\\report_'+str(_START_YEAR)+'.csv'
_FILE_NAME_CLEAN= os.path.dirname(os.path.abspath(__file__))+'\\report_'+str(_START_YEAR)+'_clean.csv'
_FILE_NAME_FINAL= os.path.dirname(os.path.abspath(__file__))+'\\report_'+str(_START_YEAR)+'_final.csv'
_TITLE_NAME= os.path.dirname(os.path.abspath(__file__))+'\\title.csv'
_URL = 'https://en.wikipedia.org/wiki/'+str(_START_YEAR)+'_in_literature'

class Lit(object):
   
    def arg_parser(self):
        global _START_YEAR
        global _END_YEAR
        parser = optparse.OptionParser()

        parser.add_option('-s', '--start',
        action="store", dest="start",
        help="start year", default=1900, type = int)
        
        parser.add_option('-e', '--end',
        action="store", dest="end",
        help="end year", default=2000, type = int)
        
        options, args = parser.parse_args()
        
        _START_YEAR = options.start
        _END_YEAR = options.end
        print 'Start year:', _START_YEAR
        print 'End year:', _END_YEAR
        
    def wiki_scrapping(self):
        global _FILE_NAME
        global _URL
        r = urllib.urlopen(_URL).read()
        soup = BeautifulSoup(r,'html.parser')
        line =1      
          
        for div in soup.find_all('li'):
            try:
                FIND_AUTHOR = div.find('a').text
            except AttributeError:
                FIND_AUTHOR = 'none'
                pass
            try:
                FIND_BOOK = div.find('i').text
            except AttributeError:
                FIND_BOOK = 'none'
                pass
            try:
                FIND_AUTHOR_1 = div.find('a')
                FIND_AUTHOR_LINK = 'https://en.wikipedia.org'+str(FIND_AUTHOR_1.get('href').encode('ascii', 'replace').decode('ascii'))
            except AttributeError:
                FIND_AUTHOR_LINK = 'none'
                pass
            
            with open (_FILE_NAME,"ab") as logFile:
                    logFileWriter = csv.writer(logFile)
                    logFileWriter.writerow([FIND_AUTHOR.encode('ascii', 'replace').decode('ascii')]
                                           +[FIND_BOOK.encode('ascii', 'replace').decode('ascii')]+[FIND_AUTHOR_LINK]) 
              
    def Cell_cleaning(self):
        global _FILE_NAME
        global _FILE_NAME_CLEAN
        word_cleaner = ['none','^','+']

        with open(_FILE_NAME) as oldfile, open(_FILE_NAME_CLEAN, 'w') as newfile:
            for line in oldfile:
                if not any(bad_word in line for bad_word in word_cleaner):
                    newfile.write(line)
        if os.path.exists(_FILE_NAME):
            os.remove(_FILE_NAME) 
                        
    def write_authors(self):   
        global _FILE_NAME_CLEAN
        global _FILE_NAME_CLEAN_AUT_1
        global _FILE_NAME_CLEAN_AUT_2
        global _START_YEAR
        
        colnames = ['Author', 'Book', 'Link', ]
        data = pandas.read_csv(_FILE_NAME_CLEAN, names=colnames)
        link = data.Link.tolist()
        
        _LENGTH = 0
        
        while (_LENGTH< (len(link))):
            print 'Year: '+ str(_START_YEAR) + ' Author Line: '+str(_LENGTH)
            try:
                r = urllib.urlopen(link[_LENGTH]).read()
            except IOError:
                r = urllib.urlopen('http://google.com').read()
                with open ('io_error.csv',"ab") as logFile2:
                    logFileWriter = csv.writer(logFile2)
                    logFileWriter.writerow([_START_YEAR])  
                pass
            
            soup = BeautifulSoup(r,'html.parser')
            
            try:
                test = soup.find('span',{'class':'birthplace'}).text
                with open (_FILE_NAME_CLEAN_AUT_1,"ab") as logFile:
                    logFileWriter = csv.writer(logFile)
                    logFileWriter.writerow([test.encode('ascii', 'replace').decode('ascii')]) 
            except AttributeError:
                with open (_FILE_NAME_CLEAN_AUT_1,"ab") as logFile:
                    logFileWriter = csv.writer(logFile)
                    logFileWriter.writerow(['none']) 
                pass
            
            try:
                test2 = soup.find('td',{'style':'line-height:1.4em;'}).text
                with open (_FILE_NAME_CLEAN_AUT_2,"ab") as logFile:
                    logFileWriter = csv.writer(logFile)
                    logFileWriter.writerow([test2.encode('ascii', 'replace').decode('ascii')]) 
            except AttributeError:
                with open (_FILE_NAME_CLEAN_AUT_2,"ab") as logFile:
                    logFileWriter = csv.writer(logFile)
                    logFileWriter.writerow(['none']) 
                pass
            
            _LENGTH = _LENGTH +1
      
    def append_lit_and_authors(self):
        global _FILE_NAME
        global _FILE_NAME_CLEAN
        global _FILE_NAME_FINAL
        global _FILE_NAME_CLEAN_AUT_1
        global _FILE_NAME_CLEAN_AUT_2
        
        colnamesmain = ['Author','Title','Link']
        collocation=['Location1','Location2']
        colnames1 = ['AuthorLocation1']
        colnames2 = ['AuthorLocation2'] 
        
        data1 = pandas.read_csv(_FILE_NAME_CLEAN_AUT_1, names=colnames1)
        data2 = pandas.read_csv(_FILE_NAME_CLEAN_AUT_2, names=colnames2)
        datamain = pandas.read_csv(_FILE_NAME_CLEAN, names=colnamesmain)
        
        author_location_1 = data1.AuthorLocation1.tolist()
        author_location_2 = data2.AuthorLocation2.tolist()     
        author = datamain.Author.tolist()
        title = datamain.Title.tolist()
        #link = datamain.Link.tolist()
        
        with open(_FILE_NAME,'wb') as resultFile:
            wr = csv.writer(resultFile)
            wr.writerows(izip(author, title,author_location_1,author_location_2))
        
        if os.path.exists(_FILE_NAME_CLEAN_AUT_1):
            os.remove(_FILE_NAME_CLEAN_AUT_1)
        if os.path.exists(_FILE_NAME_CLEAN_AUT_2):
            os.remove(_FILE_NAME_CLEAN_AUT_2)  
        if os.path.exists(_FILE_NAME_CLEAN):
            os.remove(_FILE_NAME_CLEAN)
        
    def run_author_macro(self):
        global _FILE_NAME
        
        myExcel = Dispatch('Excel.Application')
        myExcel.Visible = 0
        myExcel.Workbooks.Open(_FILE_NAME)
        myExcel.Run("Book1.xlsb!Macro3")
        myExcel.DisplayAlerts = 0
        myExcel.Quit() 
            
if __name__ == "__main__":
    global _FILE_NAME_CLEAN_AUT_1
    global _FILE_NAME_CLEAN_AUT_2
    global _FILE_NAME_CLEAN_AUT_3
    
    server=Lit()
    server.arg_parser()
    
    while (_START_YEAR < _END_YEAR+1):
        print 'Year:', _START_YEAR    
        _FILE_NAME= os.path.dirname(os.path.abspath(__file__))+'\\report_'+str(_START_YEAR)+'.csv'
        _FILE_NAME_CLEAN= os.path.dirname(os.path.abspath(__file__))+'\\report_'+str(_START_YEAR)+'_clean.csv'
        _FILE_NAME_CLEAN_AUT_1= os.path.dirname(os.path.abspath(__file__))+'\\report_'+str(_START_YEAR)+'_clean1.csv'
        _FILE_NAME_CLEAN_AUT_2= os.path.dirname(os.path.abspath(__file__))+'\\report_'+str(_START_YEAR)+'_clean2.csv'
        _FILE_NAME_FINAL= os.path.dirname(os.path.abspath(__file__))+'\\report_'+str(_START_YEAR)+'_final.csv'
        _URL = 'https://en.wikipedia.org/wiki/'+str(_START_YEAR)+'_in_literature'
        
        server.wiki_scrapping()
        server.Cell_cleaning()
        server.write_authors()
        server.append_lit_and_authors()
        server.run_author_macro()
        _START_YEAR = _START_YEAR + 1
       