def tribonacci(signature, n):
    if n<len(signature):
        return signature[:n]
    else:
        for i in range(0,n-3):
            b=sum(signature[-3:])
            signature.append(b)
        return signature
