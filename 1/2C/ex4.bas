Function ConvertTo10進数(ByVal str)
    Dim is負の数
    If Left(str, 1) = "-" Then
        is負の数 = True
        
        str = Mid(str, 2)
    Else
        is負の数 = False
    End If
    
    Dim decimalNumber
    decimalNumber = 0
    
    Dim ch
    Dim num
    Dim i
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        
        If ch = " " Then Exit For
        
        num = Asc(ch) - Asc("0")
        
        decimalNumber = decimalNumber * 10 + num
    Next
    
    If is負の数 Then
        ConvertTo10進数 = decimalNumber * (-1)
    Else
        ConvertTo10進数 = decimalNumber
    End If
End Function
