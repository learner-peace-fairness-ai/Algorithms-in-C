Sub Print2進数(ByVal num As Long)
    Dim is負の数 As Boolean
    If num < 0 Then
        is負の数 = True
        
        num = Abs(num)
    Else
        is負の数 = False
    End If

    Dim arr() As Integer
    Dim i As Long
    i = 0
    Do Until num = 0
        i = i + 1
        ReDim Preserve arr(1 To i)
        
        arr(i) = num Mod 2
        
        num = num \ 2
    Loop
    
    If i = 0 Then
        ReDim Preserve arr(1 To 1)
    End If
    
    Dim i_最下位桁 As Long
    Dim i_最上位桁 As Long
    Dim i_符号ビット As Long
    i_最下位桁 = LBound(arr)
    i_最上位桁 = UBound(arr)
    i_符号ビット = i_最上位桁 + 1
    ReDim Preserve arr(1 To i_符号ビット)
    
    arr(i_符号ビット) = 0
    
    If is負の数 Then
        For i = i_最下位桁 To i_符号ビット
            If arr(i) = 1 Then
                arr(i) = 0
            Else
                arr(i) = 1
            End If
        Next
        
        For i = i_最下位桁 To i_符号ビット
            If arr(i) = 0 Then
                arr(i) = 1
                
                Exit For
            Else
                arr(i) = 0
            End If
        Next
    End If
    
    For i = i_符号ビット To i_最下位桁 Step -1
        Dim ch As String
        ch = CStr(arr(i))
        
        Debug.Print ch;
    Next
    
    Debug.Print
End Sub
