#Sam Tech OP
#Ye karke bada Maza aaya😀😎

from bs4 import BeautifulSoup
import requests
import pprint

r=requests.get("https://www.google.com/search?q=1+dollars+in+rupees+live&rlz=1C1CHBD_enIN928IN928&oq=1+dollars+in+rupees+realtime&aqs=chrome.1.69i57j0i22i30j0i390l3.29864j1j7&sourceid=chrome&ie=UTF-8")
soup = BeautifulSoup(r.content, 'html.parser')
a=soup.find_all(class_="BNeawe iBp4i AP7Wnd")
d=a[0].text
v=float(d.split()[0])

opt=int(input("Enter:-\n1 for Rupees to Dollar Conversion.\n2 for Dollar to Rupees Conversion.\n"))

def r2d(n):
    ans=float(n/v)
    return ans

def d2r(n):
    ans=float(n*v)
    return ans

if opt==1:
    n=float(input("\nEnter Rupees (in Numbers)="))
    print("Dollars=",r2d(n))
elif opt==2:
    n=float(input("\nEnter Dollars (in Numbers)="))
    print("Rupees=",d2r(n))
else:
    print("Invalid Input")
