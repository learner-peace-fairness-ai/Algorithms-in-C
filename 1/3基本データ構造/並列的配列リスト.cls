Private Const HEAD_ = 1
Private Const TAIL_ = 2

Private 未使用の先頭_
Private key_()
Private next_()

Property Get HEAD()
    HEAD = HEAD_
End Property

Sub Initialize(listSize)
    ReDim key_(1 To listSize + 2)
    ReDim next_(1 To listSize + 2)
    
    未使用の先頭_ = 3
    
    next_(HEAD_) = TAIL_
    next_(TAIL_) = TAIL_
End Sub

Sub DeleteNext(node)
    Dim 次
    次 = next_(node)
    
    Dim 次の次
    次の次 = next_(次)
    
    next_(node) = 次の次
End Sub

Function InsertAfter(val, node)
    Dim 新
    新 = 未使用の先頭_
    未使用の先頭_ = 新 + 1
    
    key_(新) = val
    next_(新) = next_(node)
    
    next_(node) = 新
    
    InsertAfter = 新
End Function
