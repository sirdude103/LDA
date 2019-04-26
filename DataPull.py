import nltk   # Dependencies needed for stopwords removal function
import os # Dependency needed to manipulate external files
import json # Depndency needed to spefically manipulate the attached json files
import xlrd # Dependency needed to open and pull the classification and testing data from MX Excel
import re # Dependency needed to strip whitespace from text during processing
from textblob import TextBlob
from textblob import Blobber  # Dependencies needed for stopwords removal function
# from nltk.corpus.util import LazyCorpusLoader # Dependency needed to support corpus associated with stopwords
# nltk.download('stopwords')   # Dependencies needed for stopwords removal function only use if stopwords are not up to date in NLTK distribution
from nltk.corpus import stopwords  # Dependencies needed for stopwords removal function
from textblob.classifiers import NaiveBayesClassifier # Dependency needed for classification

###################### Functions #############################

### Character Position Finder###
### Returns the string position of the occurrence of an indicated character
def charposition(string, char):
    pos = [] #list to store positions for each 'char' in 'string'
    for n in range(len(string)):
        if string[n] == char:
            pos.append(n)
    return pos
### END Character Position Finder###

### Attribution Space Catcher ###
### Returns either the string (data version 1) or textblob (data version 2)that precedes and follows quotes in a string
def attrib_catcher(s, start, end, d):
     if d >= start:       #if the tolerance is larger than the distance to 1st occurence
        return str(s[0:start]), str(s[end:end+d]) #data version 1)
     else:
        return str(s[start-d:start]), str(s[end:end+d]) #data version 1)
     # return (s[start-d:start]), (s[end:end+d]) #data version 2
### END Attribution Space Catcher ###

### PairCounter ###
### The loop that catches the quotes in pairs of two ###
def quote_count_pairdown(quote, quotevect, tol):
     pair_count = 0
     while pair_count<len(quotevect):
          print (quotevect[pair_count], quotevect[pair_count+1])
          print (attrib_catcher(quote, quotevect[pair_count], quotevect[pair_count+1], tol)) 
          pair_count +=2
     #return (pair_count, quotevect[pair_count], quotevect[pair_count+1])
### END Pair Counter ###

### Text Occasion Counter ###
### Counts and returns the number of occurences of a specific character in a string ###
def occCounter (t):
    alttex = t
    return alttex.count('"')
### END Occasion Counter ###

### Stopwords Extractor ###
### Removes stopwords from feature extracted text ###
stop_words = set(['Body', 'the', 'and', 'a' ,'an', 'is', 'thus', 'there', 'I', 'my', 'that', 'what', 'to', 'for', 'of', 'in', 's', 'on', 'it'])
def stopwordcatcher (s):
    sub_stopwords = set(s.words) - stop_words
    final_string=' '
    for word in sub_stopwords:
       final_string=final_string + word +' '
    return final_string
### END stopwordcatcher m###

#################### END FUNCTIONS  ###########################

### Stopwords Extractor ###
### Removes stopwords from feature extracted text ###
stop_words = set(['the', 'and', 'a' ,'an', 'is', 'thus', 'there', 'I', 'my', 'that', 'what', 'as', 'he', 'to', 'his', 'her'])
def stopwordcatcher (s):
    sub_stopwords = set(s.words) - stop_words
    final_string=' '
    for word in sub_stopwords:
       final_string = final_string + word +' '
    return final_string
### END stopwordcatcher m###

book = xlrd.open_workbook("/home/terryt/Desktop/Machine Learning Research/Fake News Project Database (Responses).xlsx") #open our xls file, there's lots of extra default options in this call, for logging etc. take a look at the docs'
sheet = book.sheets()[0] #book.sheets() returns a list of sheet objects... alternatively...
train1 = []
#for i in xrange(sheet.nrows):
for i in range (0, 168):
   t1 = sheet.row(i)[5]
   tstripped = str(t1)
   tstripped = tstripped.replace('\n', '').replace('\r', '')
   wikit1 = TextBlob(tstripped.strip())
   wikit1 = stopwordcatcher(wikit1) 
   t2 = sheet.row(i)[7]
   #data.append(sheet.row_values(i)) #drop all the values in the rows into data
   #print (t1.value)
   #print (t2.value)
   datatuple = (t1.value, t2.value)
   train1.append(datatuple)
print (len(train1))
test1 = []
for i in range (169, 218):
   t1 = sheet.row(i)[5]
   t2 = sheet.row(i)[7]
   #data.append(sheet.row_values(i)) #drop all the values in the rows into data
   #print (t1.value)
   #print (t2.value)
   datatuple = (t1.value, t2.value)
   test1.append(datatuple)
print (len(test1))  

train2 = []
#for i in xrange(sheet.nrows):
for i in range (0, 168):
   t1 = sheet.row(i)[5]
   wikit1 = TextBlob(str(t1))
   #wikit1 = attrib_catcher(wikit1,) 
   wikit1 = stopwordcatcher(wikit1)
   t2 = sheet.row(i)[7]
   #data.append(sheet.row_values(i)) #drop all the values in the rows into data
   #print (t1.value)
   #print (t2.value)
   datatuple = (t1.value, t2.value)
   train2.append(datatuple)
print (len(train2))
test2 = []
for i in range (169, 218):
   t1 = sheet.row(i)[5]
   t2 = sheet.row(i)[7]
   #data.append(sheet.row_values(i)) #drop all the values in the rows into data
   #print (t1.value)
   #print (t2.value)
   datatuple = (t1.value, t2.value)
   test2.append(datatuple)
print (len(test2))  
print (train1[1][0])