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
