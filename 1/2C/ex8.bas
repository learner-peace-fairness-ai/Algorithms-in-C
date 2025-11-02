Function gcdOf3(ByVal u As Long, ByVal v As Long, ByVal w As Long) As Long
    Dim ans As Long
    ans = gcd(u, v)
    
    gcdOf3 = gcd(ans, w)
End Function

Function gcd(ByVal u As Long, ByVal v As Long) As Long
    Dim temp As Long
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
