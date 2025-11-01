Function gcdOf3(ByVal u As Long, ByVal v As Long, ByVal w As Long) As Long
    Dim ans As Long
    ans = gcd(u, v)
    
    gcdOf3 = gcd(ans, w)
End Function
