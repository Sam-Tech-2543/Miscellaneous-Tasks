def isomorphic(s1, s2):
    s1=s1.lower()
    s2=s2.lower()
    if len(s1)==len(s2):
        for i in s1:
            a1=s2[s1.index(i)]
            n=0
            for j in s1:
                if j==i:
                    if s2[n]==a1:
                        a=0
                    else:
                        return False
                n+=1
    else:
        return False
    return True

print(isomorphic("Eggcc","addas"))