import requests
import pandas as pd
import re

def personPatten(name):
    regex = r'[^a-zA-Z]'
    patten = re.sub(regex, '.', name, count=0)
    return patten

def inputExcel():
    excel = input("請輸入Excel: ")
    return excel

def readExcel():
    excelPath = inputExcel()
    df = pd.read_excel(excelPath)
    studentName = df['姓名'].to_list()
    
    return studentName

API_KEY = ""
SEARCH_ENGINE_ID = ""

def googleSearch(name):
    # the search query you want
    # search key word
    query = name + "linkedin"
    # using the first page
    page = 1
    # constructing the URL
    # doc: https://developers.google.com/custom-search/v1/using_rest
    # calculating start, (page=2) => (start=11), (page=3) => (start=21)
    start = (page - 1) * 10 + 1
    print("Search: ", query)
    url = f"https://www.googleapis.com/customsearch/v1?key={API_KEY}&cx={SEARCH_ENGINE_ID}&q={query}&start={start}"

    # make the API request
    data = requests.get(url).json()

    # get the result items
    search_items = data.get("items")
    # iterate over 10 results found
    try:
        for i, search_item in enumerate(search_items, start=1):
            try:
                long_description = search_item["pagemap"]["metatags"][0]["og:description"]
            except KeyError:
                long_description = "N/A"
            # get the page title
            title = search_item.get("title")
            # page snippet
            snippet = search_item.get("snippet")
            # alternatively, you can get the HTML snippet (bolded keywords)
            html_snippet = search_item.get("htmlSnippet")
            # extract the page url
            link = search_item.get("link")
            # print the results
            print("="*10, f"Result #{i+start-1}", "="*10)
            print("Title:", title)
            # print("Description:", snippet)
            # print("Long description:", long_description)
            print("URL:", link, "\n")
            nameCheck = re.search(personPatten(name), title, flags=re.I)
            urlCheck = re.search("linkedin", link)
            if urlCheck != None and nameCheck != None:
                print("get")
                return name
                # return True
            else:
                print("not get")
                return None
    except:
        print("error")
        return None

if __name__ == '__main__':
    name = readExcel()
    
    recodeArray = list()
    for n in name:
        if googleSearch(n) != None:
            recodeArray.append(n)

    path = 'output.txt'
    with open(path, 'w') as f:
        for i in recodeArray:
            f.write(f"{i}\n")
    f.close()
    # print(recodeArray)

    
