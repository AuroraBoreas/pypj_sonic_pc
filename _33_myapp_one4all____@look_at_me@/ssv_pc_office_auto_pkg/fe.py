import bs4, requests, sys
# send request to WaiHui GuanLi Ju
# link: http://www.safe.gov.cn/

url = r"http://www.safe.gov.cn/"

response = requests.get(url)
if response.status_code != 200:
    sys.exit(f"{url} Not Respond")

text: str = response.text
with open("dom", "w", encoding="utf-8") as f:
    f.write(text);

# bs4 parse
soup = bs4.BeautifulSoup(text, "xml")
rmb_query = soup.find_all("li", {"class":"white"})
# rmb_query = soup.find("li", class_="white")
print(rmb_query)
