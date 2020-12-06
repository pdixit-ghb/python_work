# Library for opening url and creating  
# requests 
import urllib.request 
  
# pretty-print python data structures 
from pprint import pprint 
  
# for parsing all the tables present  
# on the website 
from html_table_parser import HTMLTableParser 
  
# for converting the parsed data in a 
# pandas dataframe 
import pandas as pd 
  
  
# Opens a website and read its 
# binary contents (HTTP Response Body) 
def url_get_contents(url): 
  
    # Opens a website and read its 
    # binary contents (HTTP Response Body) 
  
    #making request to the website 
    req = urllib.request.Request(url=url) 
    f = urllib.request.urlopen(req) 
  
    #reading contents of the website 
    return f.read() 
  
# defining the html contents of a URL. 
xhtml = url_get_contents('https://clearquest.alstom.hub/cqweb/restapi/CQat/atvcm/QUERY/47363098?format=HTML&noframes=true').decode('utf-8') 
  
# Defining the HTMLTableParser object 
#p = HTMLTableParser() 
  
# feeding the html contents in the 
# HTMLTableParser object 
#p.feed(xhtml) 
  
# Now finally obtaining the data of 
# the table required 
#pprint(p.tables[1]) 
  
# converting the parsed data to 
# datframe 
print("\n\nPANDAS DATAFRAME\n") 
df_list = pd.read_html('C:/PD/MooN/CR&CQ/Automation/cq-moon-cr.html')
df = df_list[-1]
print(df_list)
#print(pd.DataFrame(p.tables[1]))