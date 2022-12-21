# google-results-scraper
input of excel file with the following format:
column D - searchterm.
colmn E - the inurl argument.
first five results will be scraped to F, G, H, I, J.
number of links can be changed as well.
the script actually opens the pages, so it takes the entire page not just the little google snippet.
needs a rotating proxy to work beyond a dozen results.
added email rocognition, but anything like phone numbers and names can also be extracted with minimal alteration to the code.
needs chromedriver either added to path or you can just specify location.
