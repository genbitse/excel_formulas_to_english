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

with open('excel_functions.txt') as f:
    lines = f.read().splitlines()
    for fc in dlformula:
        for f in lines:
            if(fc == f):
                functions.append(fc)

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
    for element in elements:
        for p in (soup.find_all('body')):
            if(all(x in p.getText() for x in tfunction)):
                #print(p.findPrevious('p').contents[0])
                print(p.getText())
                global newf
                newf = p.getText()
#                print(p.findNext().contents[0])
            
get_data(tophiturl)

newf = split_formula(newf)

# Replaces variables from formula
def varsub(newfl):
    vcount = 0
    for f in newfl:
        if(f not in tfunction) and (f not in sep):
            x = dfl['var'][vcount]
            vcount += 1
            print("Replacing %s with %s." % (f, x))
            return x

repvar = varsub(newf)

"""   
for x in ranking.values():
    try:
        for k,v in x.items():
            print(k,v)        
    except:
        pass


####

"""
"""

### TEMP TEMP TEMP TEMP

for fc in spfunction:
    for f in all_functions:
        if(fc == f):
            print(fc)
    


# Find all tables in selected languages
url = "http://www.piuha.fi/excel-function-name-translation/index.php?page=%s-%s.html" % (avail_lang[langf][0], avail_lang[langt][1])
r = requests.get(url)
soup = BeautifulSoup(r.content, "html.parser")

# Create dict and enumerate over table values
tdict = dict((i,t) for i,t in enumerate(soup.find_all('td'))) 

# Get rid of '=' for now
function = function.replace('=','')

# Split by and keep any non-alphanumeric delimiter (except .)
# the full string including delimters is added at the end
spfunction = re.split('([^\\w.])', function)

# Split by and remove delimiters, filter out empty elements
function = list(filter(None, re.split(r'[\W]+', function)))

# Only keep elements longer than 2 chars, in order to limit matching
function = [f for f in function if len(f) > 2]

# Iterate over table values, function parts
# Add original and translated values to dict
trdict = {}
#for i, t in tdict.items():
for i, t in tdict.items():
    for x in function:
         if(str(x) in str(t.getText())):
              if(langf == "en"):
                   if(t.getText() not in trdict.keys()): # Prevent duplicates
                       fr = (tdict[i+1].getText().split(','))[0]
                       to = (t.getText().split(','))[0]
                       trdict[to] = fr # If translating from English, set key to English
              else: 
                   fr = (t.getText().split(','))[0]
                   to = (tdict[i+1].getText().split(','))[0]
                   trdict[fr] = to # If translating from non-English, set key to non-English


new = [trdict.get(x,x) for x in spfunction] # Get translated value from dict

if(langf != "en"):
     new = '=' + ''.join(new).replace(';',',')
elif(langf == "en"):
     new = '=' + ''.join(new).replace(',',';')

print("\n%s"%new)
"""
