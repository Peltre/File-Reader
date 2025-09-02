# Code by Pedro Sotelo Arce | Student ID: 2025849263 | Last Modified: 02/09/2025 at 3:58pm
# Import libraries to read pdf / docx / doc files
import PyPDF2
import docx
import textract
import os

# This function asks the user to input a file name, then depending on the file extension
# it reads the file counts the total words, and using a dictionary it stores the most
# used words
# Works with .txt, .pdf, .docx & .doc
def readAndCountAll():
    filename = input("Enter file name: ")
    name, ext = os.path.splitext(filename)

    # No library needed
    if ext.lower() == ".txt":
        handle = open(filename, 'r')
        data = handle.read()
        splitAndCountWords(data)

    # Using PyPDF 2 library to handle pdf file opening & reading
    elif ext.lower() == ".pdf":
        reader = PyPDF2.PdfReader(filename)
        text = ""
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            text += page.extract_text() + "\n"
        splitAndCountWords(text)

    # Uses python-docx library to handle docx opening & reading
    elif ext.lower() == ".docx":
        doc = docx.Document(filename)
        fullText = []
        for para in doc.paragraphs:
            fullText.append(para.text)
        text = '\n'.join(fullText)
        splitAndCountWords(text)

    # Uses textract to handle doc opening (requires pip version less than 24.1) & antiword
    elif ext.lower() == ".doc":
        file = textract.process(filename)
        text = file.decode('utf-8')
        splitAndCountWords(text)
    
    # Handle empty & undesired file extensions
    else:
        print("File extension not supported by this function")

# Function to separately count the words in the file to avoid code repetition
def splitAndCountWords(text):
    words = text.split()
    TotalWords = len(words)
    print("Total words:", TotalWords)
    MostUsedWords(words)

# Function to find the most used words using a dictionary and saving them
# in a list, it also says how many times those words appear.
def MostUsedWords(words):
    counts =  dict()
    for word in words:
        counts[word] = counts.get(word, 0) + 1
    wordCount = None
    MostUsed = []
    for word, count in list(counts.items()):
        if wordCount is None or count > wordCount:
            MostUsed = [word]
            wordCount = count
        elif count == wordCount:
            MostUsed.append(word)
    print("Most used word:", MostUsed, "used a total of:", wordCount, "times")


# Call function
readAndCountAll()

# Expected output
# doc file: Total Words 84, Most used "I" and "elif" 5 times each
# docx file: Total Words 79, Most used "I"
# txt file: Total Words 85, Most used "test" 6 times each
# pdf file: Total Words 917, Most used "sit" and "ac"









