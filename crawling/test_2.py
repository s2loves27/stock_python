import requests

from bs4 import BeautifulSoup

## HTTP GET Request
req = requests.get('https://search.naver.com/search.naver?where=news&sm=tab_jum&query=%ED%9D%AC%ED%86%A0%EB%A5%98')

##HTML 소스 가져오기
html = req.text

## BeautifulSoup으로 html소스를 python객체로 변환하기
## 첫 인자는 html소스코드, 두 번째 인자는 어떤 parser를 이용할지 명시.
## 이 글에서는 Python 내장 html.parser를 이용했다.

soup = BeautifulSoup(html, 'html.parser')

my_titles = soup.select('sp_nws1 > dl > dt > a')

#sp_nws1 > dl > dt > a
print(len(my_titles))
for title in my_titles:
    print(title.text)
    print(title.get('href'))