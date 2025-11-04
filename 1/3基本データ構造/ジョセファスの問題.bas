Function ジョセファスの問題(人数, m番目)
    Dim head As Node
    Set head = New Node
    head.key = 1
    
    Dim ノード As Node
    Set ノード = head
    
    Dim key
    For key = 2 To 人数
        Set ノード.nextNode = New Node
        
        Set ノード = ノード.nextNode
        ノード.key = key
    Next
    
    Set ノード.nextNode = head
    
    Dim arr()
    Dim i
    ReDim arr(1 To 人数)
    i = 1
    
    Dim cnt
    Dim m番目のノード As Node
    Do Until ノード.nextNode Is ノード
        For cnt = 1 To m番目 - 1
            Set ノード = ノード.nextNode
        Next
        
        Set m番目のノード = ノード.nextNode
        arr(i) = m番目のノード.key
        i = i + 1
        
        Set ノード.nextNode = m番目のノード.nextNode
    Loop
    
    arr(i) = ノード.key
    
    ジョセファスの問題 = arr
End Function
