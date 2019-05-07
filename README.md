This natural language processing was created as part of a larger project that tracks the change of topics over time. 

This LDA software takes in an Excel document where the top row contains multiple headers. In order for the program to work, the last header should be the string "Body". The Body column should contain the document bodies to run the LDA against. Following the Body column, leave (TOTAL_TOPICS * TOTAL_WORDS * 2) + 2 empty columns for the program to output text. At the time of writing this README.md, there should be 5 TOTAL_TOPICS, 3 TOTAL_WORDS, and 32 required empty columns.

Also note that the LDA currently works only on the first sheet of the input Excel file.

How to run the program:

1) setup input file as described above.
2) include your input file in the directory of the program.
3) modify main.py file and change the INPUT_FILE variable at the start of the main definition to be the name of the Excel file you want to run.
4) modify main.py file and change the OUTPUT_FILE variable to the name you want the output Excel file to be.
5) Run main.py. 

When the program completes, the program will create an Excel file in the project directory with the OUTPUT_FILE title. 

Output:

(TOTAL_TOPICS * TOTAL_WORDS * 2) + 2 additional columns added onto the Excel document.
For each topic, there will be TOTAL_WORDS * 2 columns.
For each word, there will two columns: one for the frequency of the word appearing in the document, and one for word itself.

The remaining two columns will be the program's estimated best topic number that connects to the overarching documents topics, and the overarching topic number that it connects with. 

The overarching documents topics will be present on the second sheet of the output Excel document. There will be TOTAL_TOPICS rows with TOTAL_WORDS * 2 columns, presented in the same way as described above.

Important Constants:
    INPUT_FILE : string : name of the input Excel document.
    OUTPUT_FILE : string : name of the output Excel document.
    HEADER_ROW = : int : row index of the header row for the program (typically 0).
    SINGLE_DOC_TOPICS : int : total topics produced for each document.
    TOTAL_TOPICS : int : total topics produced for combining all documents together.
    TOTAL_WORDS : int : amount of words present in each topic.
    VERBOSE : boolean : if true, the program will print out console messages for test use.
