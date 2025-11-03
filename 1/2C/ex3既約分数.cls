Public 分子
Public 分母

Public Sub Simplify()
    Dim 最大公約数
    最大公約数 = gcd(分子, 分子)
    
    分子 = 分子 / 最大公約数
    分母 = 分母 / 最大公約数
End Sub

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
