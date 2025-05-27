import nltk

nltk.download('punkt_tab')

text = "Hello! This is a test sentence."
sentences = nltk.sent_tokenize(text)
print(sentences)
