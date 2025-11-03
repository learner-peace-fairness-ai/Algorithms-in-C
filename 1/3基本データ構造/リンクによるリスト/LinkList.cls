Private head_ As Node
Private tail_ As Node

Public Sub Init()
    Set head_ = New Node
    Set tail_ = New Node
    
    Set head_.nextNode = tail_
    Set tail_.nextNode = tail_
End Sub

Public Sub DeleteNext(nd As Node)
    Set nd.nextNode = nd.nextNode.nextNode
End Sub

Public Function InsertAfter(ByVal val, Optional nd As Node = Nothing)
    If nd Is Nothing Then
        Set nd = head_
    End If

    Dim newNode As New Node
    newNode.key = val
    Set newNode.nextNode = nd.nextNode
    
    Set nd.nextNode = newNode
    
    Set InsertAfter = newNode
End Function
