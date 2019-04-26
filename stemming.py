from nltk.stem.porter import PorterStemmer # Dependency for reducing words to stems

# Create p_stemmer of class PorterStemmer
p_stemmer = PorterStemmer()

# stem tokens
def stem_words( stopped_tokens ):
    return [p_stemmer.stem(i) for i in stopped_tokens]