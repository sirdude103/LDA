from gensim import corpora, models # dependency for use of bag of words
import json # depedency to work with JSON input
import xlrd # dependency used to access Excel documents

# python processing files
from tokenization import tokenize
from stopping import remove_stops
from stemming import stem_words

# make sure to pip install the following:
## gensim
## nltk
## stop_words

###################### FUNCTIONS #########################

# Finds a column given its header
# header_row : xlrd.sheet.row : the row where the column headers are located
# column_name : string : the value of the column header located in row 0
#
# returns : int : the int of the column it's located in
def find_column_by_value(header, column_value):

    body_column = 0
    for column_title in header:
        if column_title.value == column_value:
            break
        body_column += 1
    else: 
        raise Exception('"{}" was not found in the header given.'.format(column_value))
    
    return body_column

    
# Retrieves documents from an Excel document column
# file_name : string : given location of an excel file and column
#
# returns : xlrd.sheet : the first xlrd worksheet in the given file
def retrieve_sheet_from_excel(file_name):
    workbook = xlrd.open_workbook(file_name)
    return workbook.sheet_by_index(0)
    
    

# Retrieves JSON from its file and returns as a set
# file_name : string : given location of .json file
#
# returns : dictionary : the json values
def retrieve_json(file_name):
    with open(file_name, "r") as json_file:
        return json.load(json_file)

        
# Extracts topics from the doc set
# doc_set : list of strings : the document bodies where topics should be read
# num_of_topics : int : the amount of different topics expected out of all the documents
# num_of_words : int : the amount of words returned to represent a topics
# 
# returns : list of tuples : the major topics and and representing words as a list of tuples
def run_extractor(doc_set, num_of_topics, num_of_words):
    raw = []
    
    if type(doc_set) == str:
        doc_set = [doc_set]
        
    for doc in doc_set:
    
        # odd scenario where our document text is given in a single element list
        if type(doc) == list:
            doc = doc[0]
        raw_doc = doc.lower()
        
        # convert to tokens
        tokens = tokenize( raw_doc )
        
        # remove stop words from tokens
        stopped_tokens = remove_stops( tokens )
        
        # stem token
        stemmed_tokens = stem_words( stopped_tokens )
        
        # add to our list
        raw.append( stemmed_tokens )
        
    dictionary = corpora.Dictionary( raw )
    
    # convert to bag of words 
    # tuples where tuple[0] is word key and tuple[1] is word occurence in given document
    corpus = [dictionary.doc2bow(text) for text in raw]
    
    ldamodel = models.ldamodel.LdaModel(corpus, num_topics=num_of_topics, id2word = dictionary, passes=20)
    
    return ldamodel.print_topics(num_topics=num_of_topics, num_words=num_of_words)


# Splits a single document into smaller blocks in order to be extracted
# doc : string : text content of a document
# blocks : int : desired amount of splitted blocks
# 
# returns : dictionary : the blocks of the document
def split_doc(doc, blocks):
    raise Exception("Function: 'split_doc' is currently not implemented.")
    
    
def topic_match(topic1, topic2):
    match_count = 0
    for word1 in topic1[1]:
        for word2 in topic2[1]:
            if word1 == word2:
                match_count += 1
    return match_count
    
    
# Takes a tuple in ldamodel bag of words format and unpacks it
# given_tuple : tuple : tuple to unpack in form of (int, 'float*"string"' + ...)
# 
# returns : list of 2 lists : first list is frequency of words. second list is the words
def unpack_tuple(given_tuple):
    freqs = []
    words = []
    output = [freqs, words]
    
    # We ignore the ID of the given tuple
    freq_word_combos = given_tuple[1].split(" + ")
    
    for combo in freq_word_combos:
        combo = combo.split("*")
        freqs.append(combo[0])
        words.append(combo[1])
    
    return output
    
###################### END FUNCTIONS ######################

###################### DEPRECATED CODE ####################

# # plaintext test documents
from test_docs import test_documents
from full_test_doc import full_doc

# retrieve text from JSON
#test_json = retrieve_json("FormattedJSONTest.json")
#document_texts = [doc["text"] for doc in test_json["documents"]]

#print("Type of document_texts: {}".format(type(document_texts)))
#topics = run_extractor(document_texts, 5, 3)

#for topic in topics:
#    print(topic)

#print("Type of full_doc: {}".format(type(full_doc)))
#print("Content: {}".format(full_doc))
#topics = run_extractor(full_doc, 5, 3)

#for topic in topics:
#    print(topic)   

#################### END DEPRECATED CODE ##################


def main():
    # Declare constants
    SINGLE_DOC_TOPICS = 1
    TOTAL_TOPICS = 5
    TOTAL_WORDS = 5
    
    # Retrieve documents from the excel file
    worksheet = retrieve_sheet_from_excel("Column_Setup.xlsx")
    column_titles = worksheet.row(0)
    body_column_id = find_column_by_value(column_titles, "Body")
    document_texts = [ worksheet.cell(row_id, body_column_id).value for row_id in range(worksheet.nrows) ]
    documents_topics = []
    
    # Find the topics of each document
    for text in document_texts: 
        print("Document texts: {}".format(text))
        print("Type of the variable: {}".format(type(text)))
        print("Topics: ")
        
        topics = run_extractor(text, SINGLE_DOC_TOPICS, TOTAL_WORDS)
        document_topics = []
        
        for topic in topics:
            print(topic)
            document_topics.append(unpack_tuple(topic))
        
        documents_topics.append(document_topics)
    
    print("*****All the docs together")
    
    # Find the topics of all the documents together
    print("Type of the variable: {}".format(type(document_texts)))
    print("Topics: ")
    
    topics = run_extractor(document_texts, TOTAL_TOPICS, TOTAL_WORDS)
    
    overarching_topics = []
    for topic in topics:
        print(topic)
        overarching_topics.append(unpack_tuple(topic))
    
    # Match document with its most accurate topic
    print(overarching_topics)
    
    # current nested for loop is processing len(documents_topics) * SINGLE_DOC_TOPICS times.
    documents_most_related_topic_id = []
    
    for document_topics in documents_topics:
        doc_topic_id = -1
        
        doc_best_topic = 0 # ID of the document's best topic
        overaching_best_topic = 0 # ID of the connected overarching topic
        doc_best_match = 0 # Best matches out of every doc_topic-overarching_topic combo
    
        for document_topic in document_topics: # loops SINGLE_DOC_TOPICS times
            doc_topic_id += 1
            
            overarching_topic_id = -1
            overarching_best_id = 0
            overarching_best_match = 0
            
            for overarching_topic in overarching_topics:
                overarching_topic_id += 1
                current_match = topic_match(overarching_topic, document_topic)
                if current_match > overarching_best_match:
                    overarching_best_match = current_match
                    overarching_best_id = overarching_topic_id
                    if overarching_best_match == TOTAL_WORDS: 
                        break
                        
            if overarching_best_match > doc_best_match:
                doc_best_topic = doc_topic_id
                overarching_best_topic = overarching_best_id
                doc_best_match = overarching_best_match
                
                if doc_best_match == TOTAL_WORDS:
                    break
            
        documents_most_related_topic_id.append(overarching_best_topic)   
            
    print("BEST MATCHING TOPICS: ")
    for i in range(len(documents_most_related_topic_id)):
        print(documents_most_related_topic_id[i])
        
main()
