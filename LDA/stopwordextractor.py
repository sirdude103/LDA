import nltk   # Dependencies needed for stopwords removal function
import os # Dependency needed to manipulate external files
import json # Depndency needed to spefically manipulate the attached json files
import xlrd # Dependency needed to open and pull the classification and testing data from MX Excel
from textblob import TextBlob
from textblob import Blobber  # Dependencies needed for stopwords removal function
# from nltk.corpus.util import LazyCorpusLoader # Dependency needed to support corpus associated with stopwords
# nltk.download('stopwords')   # Dependencies needed for stopwords removal function only use if stopwords are not up to date in NLTK distribution
from nltk.corpus import stopwords  # Dependencies needed for stopwords removal function
from textblob.classifiers import NaiveBayesClassifier # Dependency needed for classification

stop_words = set(['the', 'and', 'a' ,'an', 'is', 'thus', 'there', 'I', 'my', 'that', 'what', '\\n\\nThe'])

book = xlrd.open_workbook("/home/terryt/Desktop/Machine Learning Research/Fake News Project Database (Responses).xlsx") #open our xls file, there's lots of extra default options in this call, for logging etc. take a look at the docs'
sheet = book.sheets()[0] #book.sheets() returns a list of sheet objects... alternatively...
train1 = []
#for i in xrange(sheet.nrows):
for i in range (0, 3):
   t1 = sheet.row(i)[5]
   wikit1 = TextBlob(str(t1))
   sub_stopwords = set(wikit1.words) - stop_words
   final_string=' '
   for word in sub_stopwords:
      final_string = final_string + word +' '
   t2 = sheet.row(i)[7]
   #train1.append(sheet.row_values(i)) #drop all the values in the rows into data
   #print (t1.value)
   #print (t2.value)
   datatuple = (final_string, t2.value)
   train1.append(datatuple)
print (len(train1))


b= TextBlob("the quick brown fox jumped over a lazy dog. ")

sub_stopwords = set(b.words) - stop_words

final_string=' '
for word in sub_stopwords:
    final_string = final_string + word +' '
print (final_string)
print (train1[1])