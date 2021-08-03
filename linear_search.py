n="0123456789"
key=input("Enter Number between 0 to 10:")
c=0
for i in n:
    if int(i)==int(key):
        print("Element Found at position ",c)
    c=c+1

if c==len(n):
    print("Not in string")