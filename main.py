from openpyxl import Workbook
from openpyxl.drawing.image \
import Image
import requests
import io


wb = Workbook()
ws = wb.active

num = input("How many pokemons do you want in the Excel? (1-151)\n> ")
num = int(num)
    

api_url = \
'https://pokeapi.co/api/'+\
f'v2/pokemon/?limit={num}'


res = requests.get(api_url)\
.json()['results']

ws.column_dimensions['A']\
.width=15
ws.column_dimensions['B']\
.width=15

for i in range(len(res)):
    name = res[i]['name']
    
    img_url = requests.get(res[i]['url'])\
    .json()['sprites']\
    ['front_default']
    
    img_res = requests.get(img_url)
    img_file = io.BytesIO(img_res.content)
    img = Image(img_file)
    
    ws.row_dimensions[i+1]\
    .height=70
    
    ws['A'+str(i+1)] = name
    ws.add_image(img, 'B'+str(i+1))
wb.save('pokemons.xlsx')