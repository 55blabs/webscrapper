from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'
print(excel.sheetnames)
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'Imdb Rating'])


try:
    source = requests.get('https://www.imdb.com/chart/top/')
    #holds the source of each html website
    source.raise_for_status()
    
    soup = BeautifulSoup(source.text, 'html.parser')
    
    movies = soup.find('tbody', class_="lister-list").find_all('tr')
    
    for movie in movies:
        #loops through our source for information to fetch
        name = movie.find('td', class_ = "titleColumn").a.text
        ##above allows you to extract the title (clean)
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        #above code allows us to finalize the rank with all inline characters, and spaces stripped away
        year = movie.find('td', class_="titleColumn").span.text.strip('()')
        #strips away the return so it is clean, and doesn't have any added parenthesis, commas
        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text
        
        print(rank, name, rating, year)
        #breaks our program so it doesn't continue to loop through our website
        sheet.append([rank, name, year, rating])
        #above will have all my data from my IMDB database and pass it in my excel sheet
    
except Exception as e:
    print(e)
excel.save('IMDB Movie Ratings.xlsx')
#excel.save will save it with the title
