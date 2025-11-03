Public Function gcdOf3つの整数(ByVal u, ByVal v, ByVal w)
    Dim ans
    ans = gcd(u, v)
    
    gcdOf3つの整数 = gcd(ans, w)
End Function

Private Function gcd(ByVal u, ByVal v)
    Dim temp
    Do Until u = 0
        If u < v Then
            temp = u
            u = v
            v = temp
        End If
        
        u = u Mod v
    Loop
    
    gcd = v
End Function
