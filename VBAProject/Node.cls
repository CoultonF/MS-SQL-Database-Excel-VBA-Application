VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is serving as a "Structure" for the Tree class
'I have made this tree to store only integers for the pupose of
'understanding only, but its truely easy if you want to add names
'strings in it, just add a string variable to this Node class below.
'Similarly you can add as many details( datatypes ) in this Node
'as you want to making a quite complex and flexible data storage
'mechanism.

Option Explicit

Public d As spec, Left As Node, Right As Node
'Left & Right are self-referencing there own Node class.


Public Function MinValue()
    If Left Is Nothing Then
        MinValue = Me.d.spec_id
    Else
        MinValue = Left.MinValue()
    End If
End Function
