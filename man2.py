from googlesearch import search


def google_search(query):
    for result in search(query):
        print(result)


query = "what is java"
google_search(query)
