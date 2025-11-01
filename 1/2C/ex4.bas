Function ConvertTo10進数(ByVal str As String) As Long
    Dim is負の数 As Boolean
    If Left(str, 1) = "-" Then
        is負の数 = True
        
        str = Mid(str, 2)
    Else
        is負の数 = False
    End If
    
    Dim decimalNumber As Long
    decimalNumber = 0

    Dim i As Long
    For i = 1 To Len(str)
        Dim ch As String
        ch = Mid(str, i, 1)
        
        If ch = " " Then Exit For
        
        Dim num As Integer
        num = Asc(ch) - Asc("0")
        
        decimalNumber = decimalNumber * 10 + num
    Next
    
    If is負の数 Then
        ConvertTo10進数 = decimalNumber * (-1)
    Else
        ConvertTo10進数 = decimalNumber
    End If
End Function
