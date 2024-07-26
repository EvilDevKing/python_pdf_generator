import requests
from requests.adapters import HTTPAdapter
from bs4 import BeautifulSoup

s = requests.Session()
adapter = HTTPAdapter(max_retries=3)
s.mount("http://", adapter)
s.mount("https://", adapter)

resp = s.get("https://beta.allbreedpedigree.com/login").content
soup = BeautifulSoup(resp, 'html.parser')
meta_tags = soup.find_all('meta')
token = ""
for tag in meta_tags:
    if 'name' in tag.attrs:
        name = tag.attrs['name']
        if name == "csrf-token":
            token = tag.attrs['content']
            break
s.post("https://beta.allbreedpedigree.com/login", data={ "_token": token, "email": "brittany.holy@gmail.com", "password": "7f2uwvm5e4sD5PH" })
resp = s.get(f"https://beta.allbreedpedigree.com/search?query_type=check&search_bar=horse&g=5&inbred=Standard&breed=&query=Mistys Money N Fame")
soup = BeautifulSoup(resp.content, 'html.parser')
table = soup.select_one("table.pedigree-table tbody")
gender = soup.select_one("span#pedigree-animal-info span[title='Sex']").text
birth = soup.select_one("span#pedigree-animal-info span[title='Date of Birth']").text
result = []
## Extract the values will be input in spreadsheet ##
ids = ["MMM", "MMMM", "MM", "MMF", "MMFM", "M", "MFM", "MFMM", "MF", "MFF", "MFFM", "FMM", "FMMM", "FM", "FMF", "FMFM", "F", "FFM", "FFMM", "FF", "FFF", "FFFM"]
for id in ids:
    td_elem = table.select_one(f"td#{id}")
    next_td_elem = table.select_one(f"td#{id} + td")
    if next_td_elem.get("class") == "pedigree-cell-highlight":
        label = td_elem.select_one("div.block-name").get("title").title()
        label += "*"
        result.append(label)
    else:
        result.append(td_elem.select_one("div.block-name").get("title").title())
        
print(gender)
print(birth)
print(result)
with open("111.txt", "wb") as file:
    file.write(resp.content)
    file.close()