Const MAX_NUMBER As Long = 1000

Sub エラトステネスの篩()
    Dim i As Long
    Dim j As Long
    Dim arr(1 To MAX_NUMBER) As Boolean
    
    arr(1) = False
    
    For i = 2 To MAX_NUMBER
        arr(i) = True
    Next
    
    For i = 2 To Int(MAX_NUMBER / 2)
        If Not arr(i) Then GoTo CONTINUE
    
        For j = 2 To Int(MAX_NUMBER / i)
            arr(i * j) = False
        Next
        
CONTINUE:
    Next
    
    For i = 1 To MAX_NUMBER
        If arr(i) Then
            Debug.Print i
        End If
    Next
End Sub
