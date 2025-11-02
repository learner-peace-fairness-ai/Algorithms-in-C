Public 分子 As Long
Public 分母 As Long


Public Sub Simplify()
    If 分子 <= 0 Or 分母 <= 0 Then Exit Sub

    Dim 最大公約数 As Long
    最大公約数 = gcd(分子, 分子)
    
    分子 = 分子 / 最大公約数
    分母 = 分母 / 最大公約数
End Sub


Private Function gcd(ByVal u As Long, ByVal v As Long) As Long
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
