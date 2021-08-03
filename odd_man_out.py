def iq_test(numbers):
    lst=numbers.split(" ")
    odd=0
    even = 0
    for i in lst:
        if float(i)%2==0:
            even=even+1
        else:
            odd+=1
    u="e" if even>odd else "o"
    if u=="e":
        for i in lst:
            if float(i)%2!=0:
                indexval=lst.index(i)
                print(indexval+1)
                break
    else:
        for i in lst:
            if float(i)%2==0:
                indexval=lst.index(i)
                print(indexval+1)
                break

#iq_test("2 4 6 3 10")
iq_test("1 2 2")
