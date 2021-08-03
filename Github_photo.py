import requests
from bs4 import BeautifulSoup

aa=str(input("Enter the GitHub username:"))

try:
    r=requests.get('https://github.com/'+str(aa))
    soup=BeautifulSoup(r.content,'html.parser')

    a=soup.find_all(class_="avatar avatar-user width-full border color-bg-primary")
    try:
        print("Link to Dev's Profile Photo->\t",a[0].get('src'))

        r=requests.get(a[0].get('src'))
        f=open(str(aa)+str('.png'),'wb')
        f.write(r.content)

        print("Image Downloaded!")
    except:
        print("Account Not Found.")
except:
    print("Error in Connection!😔")

