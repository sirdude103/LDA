from stop_words import get_stop_words # dependency for a list of English stop words

# create English stop words list
en_stop = get_stop_words('en')

# remove stop words from tokens
def remove_stops( tokened_input ):
    return [i for i in tokened_input if not i in en_stop]