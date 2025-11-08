Private Const HEAD = 1
Private Const TAIL = 2

Private 未使用
Private key()
Private 次()


Sub Initialize(listSize)
    ReDim key(1 To listSize + 2)
    ReDim 次(1 To listSize + 2)
    
    未使用 = 3
    
    次(HEAD) = TAIL
    次(TAIL) = TAIL
End Sub

Sub DeleteNext(node)
    Dim nextNode
    nextNode = 次(node)
    
    次(node) = 次(nextNode)
End Sub

Function InsertAfter(val, node)
    Dim newNode
    newNode = 未使用
    未使用 = newNode + 1
    
    key(newNode) = val
    次(newNode) = 次(node)
    
    次(node) = newNode
    
    InsertAfter = newNode
End Function

Sub PrintAll()
    Dim node
    node = 次(HEAD)
    
    Do Until node = TAIL
        Debug.Print key(node)
        node = 次(node)
    Loop
End Sub
