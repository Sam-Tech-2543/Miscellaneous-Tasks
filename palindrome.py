lst=[]
def palindrome(w):
    if w==w[::-1]:  lst.append(w)

a=(input("Enter a Word:")).lower()

for i in range(len(a)):
    for j in range(i,len(a)+1):
        if len(a[i:j])>=2:  palindrome(a[i:j])

if lst==[]:
    print(-1)
else:
    lst.sort()
    print("Shortest Palindrome:",lst[0])
    print("Longest Palindrome:",lst[-1])