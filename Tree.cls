VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'Its a small class but if you understand it, its going to clear lot
'of your concepts about objects of Visual Basic and of Data structures.
Option Explicit
Private t As spec
Private Parent As Node, Tmp As Node
Private m_List As ListBox

Private treeSize As Integer

Private list As Collection

Public Function init()

    Set list = New Collection


End Function

Public Function getCollection()

    Set getCollection = list

End Function

Public Sub Insert(X As spec)
    Set Tmp = New Node
    Set Tmp.d = X
    Set Tmp.Left = Nothing
    Set Tmp.Right = Nothing
    Set t = X
    If Parent Is Nothing Then
        Set Parent = Tmp
    Else
        Recurse Parent
    End If
End Sub

'used to remove a node from the tree
Public Function removeNode(rootNode As Node, id As Integer)

        
    If rootNode Is Nothing Then
        Set removeNode = New Node
    
    
    ElseIf id = rootNode.d.spec_id Then
    
        If rootNode.Left Is Nothing Then
        
            If rootNode.Right Is Nothing Then
            
                Set rootNode = Nothing
            
            Else
            
                Set rootNode = rootNode.Right
                
            End If
        
        ElseIf rootNode.Right Is Nothing Then
        
            Set rootNode = rootNode.Left
        
        Else
        
            Dim leftMostNode As Node
            
            Set leftMostNode = getLeftmostNode(rootNode.Right)
            
            Call removeNode(Parent, leftMostNode.d.spec_id)
        
            Set rootNode.d = leftMostNode.d
            
        End If
    
    ElseIf id < rootNode.d.spec_id Then
        Set rootNode.Left = removeNode(rootNode.Left, id)
    ElseIf id > rootNode.d.spec_id Then
        Set rootNode.Right = removeNode(rootNode.Right, id)
    End If
    Set removeNode = rootNode
End Function



'if the given data is less then current node data then it'll become its left
'node, else it'll become its right node

Private Sub Recurse(tmpx As Node)
    If t.spec_id >= tmpx.d.spec_id Then
        If Not (tmpx.Right Is Nothing) Then
            Recurse tmpx.Right
        Else
            Set tmpx.Right = Tmp
        End If
    Else
        If Not (tmpx.Left Is Nothing) Then
            Recurse tmpx.Left
        Else
            Set tmpx.Left = Tmp
        End If
    End If
End Sub
Private Function getLeftmostNode(tmpNode As Node)

If Not tmpNode.Left Is Nothing Then
    Set getLeftmostNode = getLeftmostNode(tmpNode.Left)
Else
    Set getLeftmostNode = tmpNode
End If


End Function

Private Sub RecurseRemove(tmpx As Node)
    If t.spec_id >= tmpx.d.spec_id Then
        If Not (tmpx.Right Is Nothing) Then
            Recurse tmpx.Right
        Else
            Set tmpx.Right = Tmp
        End If
    Else
        If Not (tmpx.Left Is Nothing) Then
            Recurse tmpx.Left
        Else
            Set tmpx.Left = Tmp
        End If
    End If
End Sub
Public Sub EnumThem()
    m_List.Clear
    RecEnum Parent
End Sub

'first visit left then head then right
Private Sub RecEnum(tmpx As Node)
    If Not (tmpx.Left Is Nothing) Then RecEnum tmpx.Left
    m_List.AddItem tmpx.d
    If Not (tmpx.Right Is Nothing) Then RecEnum tmpx.Right
End Sub

Public Property Get SetListBox() As ListBox
    Set SetListBox = m_List
End Property

Public Property Let SetListBox(ByVal vNewValue As ListBox)
    Set m_List = vNewValue
End Property
Public Sub DeleteAll()
    RecursiveDelete Parent
End Sub
'While deleting, first delete left then right then head
Private Sub RecursiveDelete(tmpx As Node)
    If Not (tmpx.Left Is Nothing) Then RecursiveDelete tmpx.Left
    If Not (tmpx.Right Is Nothing) Then RecursiveDelete tmpx.Right
    Set tmpx = Nothing
End Sub
Public Sub WalkInOrder()
    clearList
    Call inOrder(Parent)
End Sub
Public Function getRootNode()

    Set getRootNode = Parent

End Function
Public Function setRootNode(setNode As Node)

    Set Parent = setNode

End Function

Public Function CountNodes()
        clearList
        treeSize = 0
        Call inOrderCount(Parent)
        CountNodes = treeSize
End Function
Public Sub WalkPreOrder()
    clearList
    Call PreOrder(Parent)
End Sub
Public Sub WalkPostOrder()
    clearList
    Call PostOrder(Parent)
End Sub
Private Function clearList()
    If Not list Is Nothing Then
        Set list = Nothing
    End If
    Set list = New Collection
End Function


Private Sub inOrder(ti As Node)
    
    If Not ti Is Nothing Then
    
        Dim size As Integer, i As Integer
        Call inOrder(ti.Left)
        Call list.Add(ti.d)
        'Debug.Print ti.d; " "
        Call inOrder(ti.Right)
        
    End If
    
End Sub
Private Function inOrderCount(ti As Node)
    
    If Not ti Is Nothing Then
    
        Dim size As Integer, i As Integer
        Call inOrderCount(ti.Left)
        treeSize = treeSize + 1
        Call inOrderCount(ti.Right)
        
    End If
    
End Function
Private Sub PreOrder(ti As Node)
    If Not ti Is Nothing Then
        Debug.Print ti.d.spec_id; " ";
        Call PreOrder(ti.Left)
        Call PreOrder(ti.Right)
    End If
End Sub
Private Sub PostOrder(ti As Node)
    If Not ti Is Nothing Then
        Call PostOrder(ti.Left)
        Call PostOrder(ti.Right)
        Debug.Print ti.d.spec_id; " ";
    End If
End Sub
