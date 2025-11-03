Function gcd(ByVal u, ByVal v)
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
