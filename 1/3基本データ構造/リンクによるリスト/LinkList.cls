Public head As Node
Public tail As Node

Public Sub Initialize()
    Set head = New Node
    Set tail = New Node

    Set head.nextNode = tail
    Set tail.nextNode = tail
End Sub

Public Sub DeleteNext(nd As Node)
    Set nd.nextNode = nd.nextNode.nextNode
End Sub

Public Function InsertAfter(ByVal val, nd As Node) As Node
    Dim newNode As New Node
    newNode.key = val
    Set newNode.nextNode = nd.nextNode

    Set nd.nextNode = newNode

    Set InsertAfter = newNode
End Function
