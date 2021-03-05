import requests
from bs4 import BeautifulSoup
import pandas as pd



URL = 'https://www.kijiji.ca/b-canada/iphone-11/k0l0?rb=true&sort=relevancyDesc&dc=true'
#headers = {"User-Agent":'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.87 Safari/537.36'}
"""
User Agent string contains information about your web browser name, 
operating system, device type and lots of other useful bits of information.
"""
headers = {"User-Agent":'Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36'}
page = requests.get(URL, headers=headers)


soup = BeautifulSoup(page.content, 'html.parser')
products=[] #List to store name of the product
prices=[] #List to store price of the product
locations = [] #list to locate the seller
descriptions = []
#print(soup.prettify())

for a in soup.find_all("div", class_="info-container"):
        price = a.find('div', attrs={'class': 'price'})
        product = a.find('div', attrs = {'class':'title'})
        location = a.find('div', attrs = {'class':'location'})
        description = a.find('div', attrs={'description'})
        print(description)
        products.append(product.text)
        prices.append(price.text)
        locations.append(location.text)
        descriptions.append(description.text)
        df = pd.DataFrame({'Product Name': products, 'Price': prices, 'Locations': locations,'description' : descriptions})
df.to_excel('products.xlsx', index=False, encoding='utf-8')
        