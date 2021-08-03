t=input("Enter hour minutes, for ex:-  3 45\n")
temp=t.split(" ")
hr,mi= int(temp[0]), int(temp[1])

if hr==0:
    thr=0
    tmi=60-mi
elif mi==0:
    thr=12-hr
    tmi=0
elif hr<=11 and mi<=59:
    thr=12-hr
    tmi=60-mi
else:
    print("Invalid Input!")

try:
    print(thr, tmi)
except:
    pass