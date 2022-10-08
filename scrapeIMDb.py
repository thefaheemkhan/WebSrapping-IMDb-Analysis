from bs4 import BeautifulSoup
import requests , openpyxl

# converting data into Excel sheetwith openpyxl library
excel = openpyxl.Workbook()
# print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'
print(excel.sheetnames)
sheet.append(['Movie Rank','Movie name','Year of Release','IMDb Rating'])  #defining column header



# using IMDb website we are going to extrace data and loaded to the CSV file 
try:
    source = requests.get('https://www.imdb.com/chart/top/') # response is going to be stored in source variable and response is in HTML format
    source.raise_for_status() # this code will give you error when website is not reachebale or link is broken , using try and catch to prevent to crash app
    soup = BeautifulSoup(source.text,'html.parser')
    soupencode = soup.encode('utf-8')  # encoding utf-8 to encode text and print
    # print(soupencode)
    movies = soup.find('tbody',class_="lister-list").find_all('tr')  #beautifilsoup having method (.find) (.find_all) to fetch data from tag
    
    #in order to iterate each tr tag we are going to apply loop
    for movie in movies:

        name = movie.find('td', class_='titleColumn').a.text
        rank = movie.find('td', class_='titleColumn').get_text(strip=True).split('.')[0]
        year = movie.find('td', class_='titleColumn').span.text.strip('()')
        rating = movie.find('td', class_='ratingColumn imdbRating').strong.text
        print(rank, name, year, rating)
        sheet.append([rank,name,year,rating]) # loading data in excel file 
        


except Exception as e:
    print(e)

excel.save('IMDb Movie Ratings.xlsx')
