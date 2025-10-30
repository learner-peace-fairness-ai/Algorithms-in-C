Function gcd(ByVal u As Long, ByVal v As Long) As Long
    Do While u > 0
        If u < v Then
            Dim temp As Long
            
            temp = u
            u = v
            v = temp
        End If
        
        u = u Mod v
    Loop
    
    gcd = v
End Function
