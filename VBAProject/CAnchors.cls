VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAnchors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colAnchors As Collection
Private m_frmParent As Object
Public MinimumWidth As Single
Public MinimumHeight As Single
Public OrigLeft As Single
Public OrigTop As Single
Public OrigWidth As Single
Public OrigHeight As Single
Public UpdateWhilstDragging As Boolean

Private Const MRESIZEHANDLE = "ResizeGrabHandle"
Private m_sngLeftResizePos As Single
Private m_sngTopResizePos As Single
Private m_blnResizing As Boolean
Public WithEvents ResizeHandle As MSForms.Label
Attribute ResizeHandle.VB_VarHelpID = -1

Public Property Set Add(RHS As MSForms.Control)

    Dim clsTemp As CAnchor
    
    Set clsTemp = New CAnchor
    Set clsTemp.Cnt = RHS
    With clsTemp
        .AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyletop
        .MinimumWidth = .Cnt.Width
        .MinimumHeight = .Cnt.Height
        .OrigLeft = .Cnt.Left
        .OrigTop = .Cnt.Top
        .OrigWidth = .Cnt.Width
        .OrigHeight = .Cnt.Height
    End With
    m_colAnchors.Add clsTemp, clsTemp.Cnt.Name
    
End Property
Public Sub Remove(Index As Variant)

    On Error Resume Next
    m_colAnchors.Remove Index
    
End Sub


Private Sub m_AddResizer(Parent As Object)
'
' add resizing control to bottom righthand corner of userform
'
    Set ResizeHandle = Parent.Controls.Add("Forms.label.1", MRESIZEHANDLE, True)
    With ResizeHandle
        With .Font
            .Name = "Marlett"
            .Charset = 2
            .size = 14
            .Bold = True
        End With
        .BackStyle = fmBackStyleTransparent
        .AutoSize = True
        .BorderStyle = fmBorderStyleNone
        .Caption = "o"
        .MousePointer = fmMousePointerSizeNWSE
        .ForeColor = RGB(100, 100, 100)
        .ZOrder
        .Top = Parent.InsideHeight - .Height
        .Left = Parent.InsideWidth - .Width
    End With
    
End Sub

Private Sub ResizeHandle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Button = 1 Then
        m_sngLeftResizePos = X
        m_sngTopResizePos = Y
        m_blnResizing = True
    End If
    
End Sub
Private Sub ResizeHandle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim sngSize As Single
    
    If Button = 1 Then
        With ResizeHandle
            .Move .Left + X - m_sngLeftResizePos, .Top + Y - m_sngTopResizePos
            sngSize = m_frmParent.Width + X - m_sngLeftResizePos
            If sngSize < Me.MinimumWidth Then sngSize = MinimumWidth
            m_frmParent.Width = sngSize
            
            sngSize = m_frmParent.Height + Y - m_sngTopResizePos
            If sngSize < MinimumHeight Then sngSize = MinimumHeight
            m_frmParent.Height = sngSize
            .Left = m_frmParent.InsideWidth - .Width
            .Top = m_frmParent.InsideHeight - .Height
            
            If UpdateWhilstDragging Then
                m_UpdateControls
            End If
        
        End With
    End If
    
End Sub
Private Sub ResizeHandle_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        If Not UpdateWhilstDragging Then
            m_UpdateControls
        End If
        m_blnResizing = False
    End If
End Sub

Public Function Anchor(Index As Variant) As CAnchor
' access to specific anchored control
    On Error Resume Next
    Set Anchor = m_colAnchors(Index)
End Function

Public Function Anchors() As Collection
' access to the collection of anchored controls
    Set Anchors = m_colAnchors
End Function


Public Property Set Parent(RHS As Object)
'
' Use this to assign all default properties
'
    Dim clsTemp As CAnchor
    Dim cntTemp As MSForms.Control
    
    Set m_frmParent = RHS
        
    UpdateWhilstDragging = False
    With RHS
        MinimumWidth = .Width
        MinimumHeight = .Height
        OrigLeft = 1
        OrigTop = 1
        OrigWidth = .InsideWidth
        OrigHeight = .InsideHeight
    End With
    
    For Each cntTemp In m_frmParent.Controls
        Set clsTemp = New CAnchor
        Set clsTemp.Cnt = cntTemp
        With clsTemp
            .AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyletop
            .MinimumWidth = cntTemp.Width
            .MinimumHeight = cntTemp.Height
            .OrigLeft = cntTemp.Left
            .OrigTop = cntTemp.Top
            .OrigWidth = cntTemp.Width
            .OrigHeight = cntTemp.Height
        End With
        m_colAnchors.Add clsTemp, clsTemp.Cnt.Name
    Next
    
    m_AddResizer RHS
    
End Property

Private Sub Class_Initialize()

    Set m_colAnchors = New Collection
    
End Sub

Private Sub Class_Terminate()
    
    Do While m_colAnchors.count > 0
        m_colAnchors.Remove m_colAnchors.count
    Loop
    
    Set m_colAnchors = Nothing
        
    m_frmParent.Controls.Remove MRESIZEHANDLE
    Set ResizeHandle = Nothing
    
End Sub
Private Sub m_UpdateControls()
'
' Calculate New position of all controls
'
    Dim clsAnchor As CAnchor
    Dim cntTemp As MSForms.Control
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim sngHeight As Single
    Dim sngWidth As Single
    
    For Each clsAnchor In m_colAnchors
        Set cntTemp = clsAnchor.Cnt
        If clsAnchor.AnchorStyle = enumAnchorStyleNone Then
            ' do nothing with this control
        Else
            If ((clsAnchor.AnchorStyle And enumAnchorStyletop) = enumAnchorStyletop) And _
               ((clsAnchor.AnchorStyle And enumAnchorStyleBottom) = enumAnchorStyleBottom) Then
                ' maintain gap between top and bottom edges by adjusting height
                sngHeight = m_frmParent.InsideHeight - (OrigHeight - clsAnchor.OrigTop - clsAnchor.OrigHeight) - clsAnchor.OrigTop
                If sngHeight < clsAnchor.MinimumHeight Then sngHeight = clsAnchor.MinimumHeight
                If sngHeight < 0 Then sngHeight = 0
                cntTemp.Height = sngHeight
            
            ElseIf (clsAnchor.AnchorStyle And enumAnchorStyletop) = enumAnchorStyletop Then
                ' maintain gap between top leave height alone
                ' does not require code
            ElseIf (clsAnchor.AnchorStyle And enumAnchorStyleBottom) = enumAnchorStyleBottom Then
                ' maintain gap between bottom leave height alone
                sngTop = m_frmParent.InsideHeight - (OrigHeight - clsAnchor.OrigTop - clsAnchor.OrigHeight) - clsAnchor.OrigHeight
                If sngTop < clsAnchor.MinimumTop Then sngTop = clsAnchor.MinimumTop
                If sngTop < 0 Then sngTop = 0
                cntTemp.Top = sngTop
            End If
            
            If ((clsAnchor.AnchorStyle And enumAnchorStyleLeft) = enumAnchorStyleLeft) And _
               ((clsAnchor.AnchorStyle And enumAnchorStyleRight) = enumAnchorStyleRight) Then
                ' maintain gap between left and right edges by adjusting Width
                sngWidth = m_frmParent.InsideWidth - (OrigWidth - clsAnchor.OrigLeft - clsAnchor.OrigWidth) - clsAnchor.OrigLeft
                If sngWidth < clsAnchor.MinimumWidth Then sngWidth = clsAnchor.MinimumWidth
                If sngWidth < 0 Then sngWidth = 0
                cntTemp.Width = sngWidth
            ElseIf (clsAnchor.AnchorStyle And enumAnchorStyleLeft) = enumAnchorStyleLeft Then
                ' maintain gap between left leave Width alone
                ' does not require code
            ElseIf (clsAnchor.AnchorStyle And enumAnchorStyleRight) = enumAnchorStyleRight Then
                ' maintain gap between Right leave Width alone
                sngLeft = m_frmParent.InsideWidth - (OrigWidth - clsAnchor.OrigLeft - clsAnchor.OrigWidth) - clsAnchor.OrigWidth
                If sngLeft < clsAnchor.MinimumLeft Then sngLeft = clsAnchor.MinimumLeft
                If sngLeft < 0 Then sngLeft = 0
                cntTemp.Left = sngLeft
            End If
        End If
    Next
    DoEvents
    
End Sub


