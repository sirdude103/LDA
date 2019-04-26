################## Attribution Assessment ####################

################## Dependencies ##############################
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
from textblob.classifiers import MaxEntClassifier #Test Dependency
from textblob.classifiers import DecisionTreeClassifier #Test Dependency
import numpy as np # Dependency needed to calculate Confusion Matrix, enhanced classifier analysis scores, and the confusion plot
import matplotlib.pyplot as plt # Dependency needed to generate the Confusion Matrix from the classifier analysis data
import matplotlib.pyplot as plt2  # Dependency needed to generate the second Confusion Matrix
import xlsxwriter # Dependency needed to write to the Excel file

################## End Dependencies ##########################

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

######################## DATA #################################

### Training, Development, and Test Data ###

### Manual Training Data for initial function testing ###
#train = [
 #       ('Jonathon commented that, "love this sandwich.""', 'real'),
  #      ('"I think food is great," was what John stated after tasting food', 'real'),
   #     ('Karen argued that "all money is great money"', 'fake'),
    #    ('" is my best work," commented by Erik', 'real'),
     #   ('Xavier stated. "what an awesome view"', 'real'),
      #  ('After party, someone mentioned, "I hate party."', 'fake'),
       # ('"I am tired" of stuff.', 'fake'),
 #       ('I "cannot deal with him"', 'fake'),
  #      ('"This is fake news"', 'fake'),
   #     ('my boss is horrible "and it is fake"', 'fake')]
#test =  [
 #       ('Terry stated, "beer was good."', 'real'),
  #      ('I do not "enjoy my job"', 'fake'),
   #     ('Noone stated, "aint feeling dandy today."', 'fake'),
    #    ('"I feel amazing!" was commented by Torance', 'real'),
     #   ('The Gods argued that. "Gary is a friend of mine."', 'real'),
      #  ('It is a fact that "I am not doing this."', 'fake')]
### End Manual Training Data ###

### Program Directed Training and Test Data ###
### Modify this section of the code to vary what gets put into the training, developmental, and test set
### This custom training, dev, and test set function exists because textblob and nltk can't feature extract 
### on quaotation markes which is a major extraction feature for this custom classifier.  Straight bag of words 
### approaches raises errors when handling raw data with single and double quotes
book = xlrd.open_workbook("C:/Users/gandr/Desktop/College Classes/Spring 2019 Classes/AI/LDA/Fake News Project Database (Responses).xlsx") #open our xls file, there's lots of extra default options in this call, for logging etc. take a look at the docs'
sheet = book.sheets()[0] #book.sheets() returns a list of sheet objects... alternatively...
train1 = []
#for i in xrange(sheet.nrows):
for i in range (0, 180):
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
print ('Training Set Size: ', len(train1))
test1 = []
for i in range (181, 218):
   t1 = sheet.row(i)[5]
   wikit1 = TextBlob(str(t1))
   sub_stopwords = set(wikit1.words) - stop_words
   final_string=' '
   for word in sub_stopwords:
      final_string = final_string + word +' '
   t2 = sheet.row(i)[7]
   #test1.append(sheet.row_values(i)) #drop all the values in the rows into data
   #print (t1.value)
   #print (t2.value)
   datatuple = (final_string, t2.value)
   test1.append(datatuple)
print ('Test Set Size: ', len(test1))
### End Program Directed Training Data ###

### Wiki data for initial testing ###
wiki = TextBlob('Python is an okay "high-level", general-purpose "programming" language.')

###################### END DATA ###############################

### Begin Main Program ###
### Main program now classifies on the entire body of text as opposed to just the quotes ###
### The program takes about 35-40 seconds longer to process and classify the entire corpus ###

### Troubleshooting Code ###
### Only needed when testing the code or the training mechanism

print (charposition (train1[15][0], '"'))
L1= charposition (train1[15][0], '"')
print (len(L1))
print (train1[15][1])
# testtex='Python is an okay high-level, "general-purpose" programming" langua"ge.'
# testtex=str(wiki)
# quote_count_pairdown(wiki, L1, 30)
# print (occCounter(testtex))
# print (attrib_catcher(wiki, 18, 29, 30)) 
# wiki2 = TextBlob(str(attrib_catcher(wiki, 18, 29, 30)))
# print (stopwordcatcher (wiki2))
# print (cl.classify(stopwordcatcher(wiki2)))
### End Troubleshooting Code ###

cl = NaiveBayesClassifier(train1)
# cl2 = DecisionTreeClassifier(train1)

print (cl.accuracy(test1))
# print (cl2.accuracy(test1))

# cl.show_informative_features(5

def BuildTestPerfRaw(testvec):
   TestPerformanceRaw = []
   for i in range (1, len(testvec)):
      TestVector = [testvec[i][0], testvec[i][1], cl.classify(testvec[i][0])]
      TestPerformanceRaw.append(TestVector)
   return(TestPerformanceRaw)

TestPerformance = BuildTestPerfRaw(test1)

### Function that builds the analysis data for classifier performance calculations ###
### Builds the True positive and additional datapoints to create the confusion matrix ###
### Input:  Vector with [<<testdata>>, <<actual or observed label>>, <<classifier predicted label>>] ###
### Output:  [TPF, TNF,  and etc] ###

def ClassAnalysis(dVect):
   # Initialize all outgoing variable data
   # Fakedocument Variables
   TPF = 0  # True Positive Fake Docs reference
   TNF = 0  # True Negative Fake Docs reference
   FPF = 0  # False Positive Fake Docs reference
   FNF = 0  # False Negative Fake Docs reference
   # Realdocument Variables
   TPR = 0
   TNR = 0
   FPR = 0
   FNR = 0

   #Main loop to build the analysis
   for i in range(0, len(dVect)):
      # Fakedocument Accuracy Calculations
      if dVect[i][1] == 'Fake':  #if the document is supposed to be fake
         if dVect[i][1] == dVect[i][2]:
            TPF += 1
         else:
            FNF += 1
      else:
         if dVect[i][1] == dVect[i][2]:
            TNF += 1
         else:
            FPF += 1
      # Realdocument Accuracy Calculations
      if dVect[i][1] == 'Real':  #if the document is supposed to be real
         if dVect[i][1] == dVect[i][2]:
            TPR += 1
         else:
            FNR += 1
      else:
         if dVect[i][1] == dVect[i][2]:
            TNR += 1
         else:
            FPR += 1
   
   return([TPF, TNF, FPF, FNF, TPR, TNR, FPR, FNR])

Analysis_Test = ClassAnalysis(TestPerformance)
# ConfusionData = ClassAnalyzer(Analysis_Test)

### Precision ###
### Function that calculates the precision for a classification problem ###
### Input: True Positive = TP, False Positive = FP ###
### Output: Precision = float ###
def precision(TP, FP):
    PR = (TP/(TP+FP))
    return(PR)

### True Positive Rate, Recall, or Sensitivity ###
### Function that calulcates the recall or sensitivity for a classification problem ###
### Input: True Positive = TP, False Negative = FN ###
### Output: RE = Float ###
def recall(TP, FN):
    RE = (TP/(TP+FN))
    return(RE)

### False Positive Rate ###
### Function that calulcates the FPR for a classification problem ###
### Input: False Positive = FP, True Negative = TN ###
### Output: FPR = Float ###
def falsePosRate(FP, TN):
    FPR = (FP/(FP+TN))
    return(FPR)

### True Negative Rate ###
### Function that calulcates the TNR for a classification problem ###
### Input: False Positive = FP, True Negative = TN ###
### Output: FPR = Float ###
def trueNegRate(FP, TN):
    TNR = (TN/(FP+TN))
    return(TNR)

### False Negative Rate ###
### Function that calulcates the TNR for a classification problem ###
### Input: False Positive = FP, True Negative = TN ###
### Output: FPR = Float ###
def falseNegRate(FN, TP):
    FNR = (FN/(FN+TP))
    return(FNR)

### F1 or F-Score ###
### Function that is the harmonic mean of the precision and recall or sensitivity ###
### Input: Precision = PRE, Recall = REC ###
### Output: F1 = Float ###
def fScore(PRE,REC):
    F1 = 2*((PRE*REC)/(PRE+REC))
    return(F1)

### Accuracy ###
### Function that calculates the accuracy of a classifier ###
### Input: True Positive = TP, True Negative = TN, False Positive = FP, False Negative = FN ###
### Output: Accuravy = Float ###
def ClassAccuracy(TP, TN, FP, FN):
   CA =((TP + TN)/(TN + TP + FP + FN))
   return(CA)

### Error ###
### Function that calculates the error associated with a classifier ###
### Input: Accuracy ###
### Output: Error = Float ###
def ClassError(CA):
    CE = 1-CA
    return(CE)

### Class Analyzer ###
### Functions that takes in the vector with all of the analysis and outputs the calculations for
### follow-on processing
### The vector with the TP, and etc ###
### Output: The calculations as floats

def ClassAnalyzer(AVect):

    # Precision Calculations
    pCalcFake = precision(AVect[0],AVect[2])
    pCalcReal = precision(AVect[4],AVect[6])

    # True Positive, Recall, or Sensitivity Rate Calculations
    reCalcFake = recall(AVect[2], AVect[1])
    reCalcReal = recall(AVect[6], AVect[5])

    # False Positive Rate Calculations
    fPosCalcFake = falsePosRate(AVect[0], AVect[1])
    fPosCalcReal = falsePosRate(AVect[4], AVect[5])

    # True Negative Rate Calculations
    trueNegCalcFake = trueNegRate(AVect[2], AVect[1])
    trueNegCalcReal = trueNegRate(AVect[6], AVect[5])

    # False Negative Rate Calculations
    falseNegCalcFake = falseNegRate(AVect[3], AVect[1])
    falseNegCalcReal = falseNegRate(AVect[7], AVect[5])

    # F1 Score
    fScoreFake = fScore(pCalcFake, reCalcFake)
    fScoreReal = fScore(pCalcReal, reCalcReal)

    # Accuracy Score for Classifier
    ClassAccuracyFake = ClassAccuracy(AVect[0], AVect[1], AVect[2], AVect[3])
    ClassAccuracyReal = ClassAccuracy(AVect[4], AVect[5], AVect[6], AVect[7])

    # Classifier Error
    ClassErrorFake = ClassError(ClassAccuracyFake)
    ClassErrorReal = ClassError(ClassAccuracyReal)

    return([pCalcFake, reCalcFake, fPosCalcFake, trueNegCalcFake, falseNegCalcFake, fScoreFake, ClassAccuracyFake, ClassErrorFake, 
            pCalcReal, reCalcReal, fPosCalcReal, trueNegCalcReal, falseNegCalcReal, fScoreReal, ClassAccuracyReal, ClassErrorReal])

ConfusionData = ClassAnalyzer(Analysis_Test)    
#ConfusionData = ClassAnalyzer([11, 14, 2, 9, 14, 11, 9, 2])
print('Summarized Classifier Performance')
print(' ')
print('Overal Classifier Accuracy')
print('Classifier Accuracy: %.2f' %ConfusionData[6])
print('Classifier Error: %.2f' %ConfusionData[7])
print(' ')
print('Fake Document Classification Data')
print('Fake Document Precision: %.2f' %ConfusionData[0])
print('Fake Document Recall or True Positive Rate: %.2f' %ConfusionData[1])
print('Fake Document False Positive Rate: %.2f' %ConfusionData[2])
print('Fake Document True Negative Rate: %.2f' %ConfusionData[3])
print('Fake Document False Negative Rate: %.2f' %ConfusionData[4])
print('Fake Document F1 Score: %.2f' %ConfusionData[5])
print(' ')
print('Real Document Classification Data')
print('Real Document Precision: %.2f' %ConfusionData[8])
print('Real Document Recall or True Positive Rate: %.2f' %ConfusionData[9])
print('Real Document False Positive Rate: %.2f' %ConfusionData[10])
print('Real Document True Negative Rate: %.2f' %ConfusionData[11])
print('Real Document False Negative Rate: %.2f' %ConfusionData[12])
print('Real Document F1 Score: %.2f' %ConfusionData[13])

# Confusion Matrix Figure Code
def GenConfusionMatrix(ConfData):

    m = np.array([[ConfData[1], ConfData[2], ConfData[1] ],
                  [ConfData[4], ConfData[3], ConfData[3]],
                  [ConfData[0], ConfData[8], 0]]) 

    n = np.array([[ConfData[9], ConfData[10], ConfData[9] ],
                  [ConfData[12], ConfData[11], ConfData[11]],
                  [ConfData[8], ConfData[0], 0]]) 


    x = ('Fake', 'Real', 'Recall')
    y = ('Fake', 'Real', 'Precision')

    x1 = ('Real', 'Fake', 'Recall')
    y1 = ('Real', 'Fake', 'Precision')

    plt.rcParams['xtick.bottom'] = plt.rcParams['xtick.labelbottom'] = False
    plt.rcParams['xtick.top'] = plt.rcParams['xtick.labeltop'] = True

    # plt.bar(x, y)

    plt.imshow(m, cmap='Blues')
    plt.colorbar()
    plt.ylabel('Actual Classifications')
    plt.yticks(rotation=0)
    plt.yticks(range(len(y)), y, size = 'small')
    plt.xlabel('Predicted Classifications')
    plt.xticks(rotation=90)
    plt.xticks(range(len(x)), x, size = 'small')
    plt.title('Fake Document Detection Confusion Matrix\n', y = 1.08)
    plt.show() 

    plt2.rcParams['xtick.bottom'] = plt.rcParams['xtick.labelbottom'] = False
    plt2.rcParams['xtick.top'] = plt.rcParams['xtick.labeltop'] = True

    # plt.bar(x, y)

    plt2.imshow(n, cmap='Blues')
    plt2.colorbar()
    plt2.ylabel('Actual Classifications')
    plt2.yticks(rotation=0)
    plt2.yticks(range(len(y1)), y1, size = 'small')
    plt2.xlabel('Predicted Classifications')
    plt2.xticks(rotation=90)
    plt2.xticks(range(len(x1)), x1, size = 'small')
    plt2.title('Real Document Detection Confusion Matrix\n', y = 1.08)
    plt2.show()

GenConfusionMatrix([0.85, 0.55, 0.125, 0.875, 0.45, 0.67, 0.69444, 0.30556, 0.6087, 0.86, 0.45, 0.55, 0.125, 0.71795, 0.69444, 0.30556])

def exportTestData(filename, vect):
    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    # Widen the first column to make the text clearer.
    worksheet.set_column('A:A', 20)

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    # Write some simple text.
    # worksheet.write('A1', 'Hello')

    # Text with formatting.
    # worksheet.write('A2', 'World', bold)

    # Write some numbers, with row/column notation.
    # worksheet.write(2, 0, 123)
    # worksheet.write(3, 0, 123.456)

    # testdata = [[1,2,3], [4,5,6], [7,8,9]]

    for i in range(0, len(vect)):
        for j in range(0,len(vect[0])):
            worksheet.write(i, j, vect[i][j])

    workbook.close()

exportTestData('demo8.xls', TestPerformance)
### End Main Program