Function エラトステネスの篩(ByVal maxNumber)
    Dim arr()
    ReDim arr(1 To maxNumber)
    
    arr(1) = False
    
    Dim i
    For i = 2 To maxNumber
        arr(i) = True
    Next
    
    Dim j
    For i = 2 To Int(maxNumber / 2)
        If Not arr(i) Then GoTo CONTINUE
    
        For j = 2 To Int(maxNumber / i)
            arr(i * j) = False
        Next
        
CONTINUE:
    Next
    
    エラトステネスの篩 = arr
End Function
