#!/usr/bin/env python

__author__ = "Martin Rydén"
__copyright__ = "Copyright 2016, Martin Rydén"
__license__ = "MIT"
__version__ = "1.0.1"
__email__ = "pemryd@gmail.com"

import re
import requests
from bs4 import BeautifulSoup
from GoogleScraper import scrape_with_config, GoogleSearchError
import collections

tries = 0
# Input Excel formula
while True:
    formula = input('\nFunction: ')
    if(formula[0] != '=') or (len(formula) > 150):
        if(tries == 0):
            print('Wrong input format. Try again.')
            tries += 1
        elif(tries == 1):
            print('The formula must begin with "=" and cannot exceed 150 characters.')
            tries += 1
        else:
            break
    else:
        break
#formula = '=INDEX(A1:Z10,MATCH("product",A1:A10,0),MATCH("month",A1:Z1,0))'

# Split by and keep any non-alphanumeric delimiter, filter blanks
def split_formula(f):
    split = list(filter(None, re.split('([^\\w.":!$])', f)))
    return split
    
dlformula = split_formula(formula)


# List of functions
functions = []

def find_functions(formula, flist):
    with open('excel_functions.txt') as f:
        lines = f.read().splitlines()
        for fc in formula:
            for f in lines:
                if(fc == f):
                    flist.append(fc)

find_functions(dlformula, functions)

# =INDEX(A1:Z10,MATCH("product",A1:A10,0),MATCH("month",A1:Z1,0))

# Set regex pattern:
# variables: any non-alpha-numeric char, with exceptions
# separators: any non-alpha-numeric char
# tbd: take another look at these and figure out wtf is actually going on here
variables = re.compile('([\w\.":!$]+)',re.I)
separators = re.compile(r'^\W+',re.I)

# Dictionary to keep track of which formula element is const, var, or sep
dfl = collections.defaultdict(list)

wc = [] # Same as dlformula but vars substituted for wilcards
tfunction = [] # Store elements of formula which are known functions
sep = [] # List of separators in formula

# Appends formula elements according to above description
for f in dlformula:
    if f in functions:
        wc.append(f)
        tfunction.append(f)
    else:
        a = variables.sub("*", f)
        wc.append(a)
        if(not re.search(variables, f)):
            sep.append(f)

# Appends element type to dict with appropriate element type key
for f in dlformula: 
    if(f in tfunction):
        dfl['const'].append(f)
    elif(f in sep):
        dfl['sep'].append(f)
    else:
        dfl['var'].append(f)

# Join the wilcard formula, add an extra wildcard for good luck
wcf = ''.join(wc)+"*"

print('\nSearching for substituted formula:\n%s' % wcf)

#### Scrape google for top hits ####
"""
config = {
    'use_own_ip': 'True',
    'keyword': ("%s excel formula -youtube") % wcf,
    'search_engines': ['duckduckgo'],
    'num_pages_for_keyword': 1,
    'scrape_method': 'http',
    'do_caching': 'True',
    'print_results': 'summarize'
}

urls = [] # List of result URLs

# Begin scraping
try:
    search = scrape_with_config(config)

except GoogleSearchError as e:
    pass

# Manually set max urls generated, since the built-in function is a bit wonky
maxresults = 6
r = 0
# Results - append URL for each hit to urls list
for serp in search.serps:
    for link in serp.links:
        if(r < maxresults):
            urls.append(link.link)
            r += 1
        
#### Parse into BS4 ####

# This dict will be used to store each hit with an id as key
# and chosen web element, its URL, and a total score (sum of found elements)
ranking = collections.defaultdict(dict)

# Web elements to look for
elements = ['pre', 'p', 'ul', 'td']

# Searches a web page for an element, stores matches in dict
def find_elements(element):
        for p in (soup.find_all(element)):
            if(all(x in p.getText() for x in tfunction)):
                matches[element] += 1

webid = 0 

for url in urls:
    webid += 1 # Gives an id to each web hit
    matches = collections.defaultdict(int) # New dict for each url

    r = requests.get(url)
    soup = BeautifulSoup(r.content, "html.parser") # Parses the page

    # Iterate through each chosen element, which are counted
    # using the find_elements function
    for e in elements:
        find_elements(e)
    
    ## tbd: Move to later stage ?
   
    try:
        stitle = soup.title.string
    except:
        stitle = "No title"
 
    print('\nFound matches in "%s".\n \
    URL: %s' % (stitle,url))

    ##
    
    # Adds each element an its number of matches to ranking dict
    # Also adds the URL for reference
    ranking[webid] = (matches)
    ranking[webid]['url'] = (url)

# Sums total count of elements per web hit into a score
# This score will used in decided which page is more likely
# to contain useful data
for k,v in ranking.items():
    score = 0
    for e in elements:
        score += v[e]
        ranking[k]['xscore'] = (score)

# Now, the ranking dict should look something like this:

# defaultdict(dict,
#             {1: defaultdict(int,
#                          {'p': 0,
#                           'pre': 0,
#                           'ul': 0,
#                           'url': 'http://www.example_url_1.com/',
#                           'xscore': 0}),
#              2: defaultdict(int,
#                          {'p': 1,
#                           'pre': 5,
#                           'ul': 0,
#                           'url': 'http://www.example_url_2.com/',
#                           'xscore': 6}),
#                           ...


# The top ranked url
tophiturl = ''
# The current highest ranking score
curmax = 0

for k,v in ranking.items():
    if v['xscore'] > curmax:
        curmax = v['xscore']
        tophiturl = v['url']
    
print('\nSelected web page: %s' % tophiturl)
print('\n')
"""

tophiturl = "https://www.ablebits.com/office-addins-blog/2014/08/13/excel-index-match-function-vlookup/"
# tbd: Browse selected wp, look for interesting portion p, add p-1, p to list.
#      Substitute variables with original input variables oip
#      Search p-1, for variables and replace with corresponding oip

# The fetched formula
# tbd: rewrite to allow for multiple ones
newf = ''

# Parses data of input url and checks hits for each specified element
# If a paragraph contains every function in the original formula, set newf to p
def get_data(url):
    r = requests.get(url)
    soup = BeautifulSoup(r.content, "html.parser") # Parses the page
    for p in (soup.find_all('code')):
        if(all(x in p.getText() for x in tfunction)):
            nfunctions = []
            #print(p.findPrevious('p').contents[0])
            #print(p.getText())
            global newf
            newf = p.getText()
            newf_split = split_formula(newf)
            find_functions(newf_split, nfunctions)
            if(all(y in tfunction for y in nfunctions)):
                for k in newf_split:
                    
#                print(p.findNext().contents[0])
            
get_data(tophiturl)



#newf = split_formula(newf)

#find_functions(newf, nfunctions)

# Replaces variables from formula
def varsub(newfl):
    vcount = 0
    for f in newfl:
        if(f not in tfunction) and (f not in sep):
            x = dfl['var'][vcount]
            #print(f)
            vcount += 1
            y = ''.join(newf)
            y = y.replace(f, x)
            #print(f)
            #print(y)
            #print("Replacing %s with %s." % (f, x))
            return x

repvar = varsub(newf)
