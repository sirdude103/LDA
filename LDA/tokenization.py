from nltk.tokenize import RegexpTokenizer # Dependency for reducing text to tokens
tokenizer = RegexpTokenizer(r"(\w+('\w+)?)")
# tokenizer = RegexpTokenizer(r"\w+") # tokenizer without contractions

def tokenize( raw_input ):
    # tuple[0] is important because RegexpTokenizer returns a tuple of all regexp in parenthesis,
    # which is very odd since parenthesis are also used in regexp for grouping
    return [ tuple[0] for tuple in tokenizer.tokenize( raw_input ) ] 