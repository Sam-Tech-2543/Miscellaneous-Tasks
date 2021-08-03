def center(s):
    if s[0] == 1:
        x = (s[1] + s[5]) / 2
        y = (s[2] + s[6]) / 2
        return [x, y]
    elif s[0]==0:
        return [s[2], s[3]]
    else:
        print("Invalid Inputs!")


n = int(input("Enter Number of Targets:"))
lst = []
if n >= 1 and n <= 100000:
    for i in range(n):
        j = input()
        temp = j.split(" ")
        temp = list(map(float, temp))
        lst.append(temp)
print(lst)

try:
    if len(lst) != 1:
        a = []
        for i in lst:
            a.append(center(i))
        slope = (a[1][1] - a[0][1]) / (a[1][0] - a[0][0])
        flag = "Yes"
        print(slope)
        for i in range(len(a) - 1):
            x1, x2 = a[i][0], a[i + 1][0]
            y1, y2 = a[i][1], a[i + 1][1]
            s = (y2 - y1) / (x2 - x1)
            print(s)
            if s != slope:
                flag = "No"
        print(flag)
    else:
        print("Yes")
except:
    pass
