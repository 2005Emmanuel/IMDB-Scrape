#Importing the modules
from bs4 import BeautifulSoup
import requests,openpyxl
Workbook = openpyxl.Workbook()
print(Workbook.sheetnames)
Sheet = Workbook.active
Sheet.title = "Imdb Most Ppular Movies"
print(Workbook.sheetnames)
Sheet.append(["Movie_Name","Movie_Rating"])

try:
#Getting the url of the page i want to scrape
    url = requests.get("https://www.imdb.com/chart/moviemeter/?ref_=nv_mv_mpm")
    url.raise_for_status()  #Checking if the website is accessible for scraping
    print(url)

    Result = BeautifulSoup(url.text,"html.parser") #bringing out all htnl tags in the page
    Movies = Result.find("tbody",class_="lister-list").find_all("tr") #finding the parent container for all tags


    for Movie in Movies:  #looping Through all tags stated here
        Name = Movie.find("td",class_="titleColumn").get_text(strip=True).split('.')[0]
        Rating = Movie.find("td", class_="ratingColumn imdbRating").get_text(strip=True).split(',')[0]
        print(Name,Rating)
        Sheet.append([Name,Rating]) #appending the looped data to an excel worksheet

except Exception as e:
 print(e)
# Workbook.save("Jumia_web_scraping.xlsx")