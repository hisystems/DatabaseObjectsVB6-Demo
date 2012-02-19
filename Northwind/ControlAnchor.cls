VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlAnchor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'these are the values for the side to be anchored
Private Const pcintLeftSide As Integer = 1
Private Const pcintTopSide As Integer = 2
Private Const pcintRightSide As Integer = 4
Private Const pcintBottomSide As Integer = 8

Public Enum ControlAnchorEnum
    ccAnchorNone = 0
    'ccAnchorLeft = 1
    'ccAnchorTop = 2
    'ccAnchorLeftTop = pcintLeft Or pcintTop
    'ccAnchorTopRight = pcintTopSide Or pcintRightSide
    'ccAnchorLeftTopRight = pcintLeftSide Or pcintTopSide Or pcintRightSide
    'ccAnchorLeftBottom = pcintLeftSide Or pcintBottomSide
    'ccAnchorLeftTopBottom = pcintLeftSide Or pcintTopSide Or pcintBottomSide
    ccAnchorRight = pcintRightSide
    ccAnchorLeftRight = pcintLeftSide Or pcintRightSide
    ccAnchorBottom = pcintBottomSide
    ccAnchorTopBottom = pcintTopSide Or pcintBottomSide
    ccAnchorRightBottom = pcintRightSide Or pcintBottomSide
    ccAnchorLeftRightBottom = pcintLeftSide Or pcintRightSide Or pcintBottomSide
    ccAnchorTopRightBottom = pcintTopSide Or pcintRightSide Or pcintBottomSide
    ccAnchorAll = pcintLeftSide Or pcintTopSide Or pcintRightSide Or pcintBottomSide
End Enum

Private Type typControl
    Control As Control
    Anchor As ControlAnchorEnum
    'distance from window to control
    'ie. lngRight = distance from the right side of the window
    'to the right side of the control
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private WithEvents pfrmForm As Form
Attribute pfrmForm.VB_VarHelpID = -1
Private pvControls() As typControl

Public Sub Initialize(ByVal frmForm As Form)
'
' Purpose:      x
' Arguments:    None
' Returns:      None
' Date:         18/07/2001
' Author:       Toby Wicks
'

    Set pfrmForm = frmForm
    ReDim pvControls(0 To 0) As typControl
    
End Sub

Public Sub AddControl(ByVal objControl As Control, ByVal eAnchor As ControlAnchorEnum)
'
' Purpose:      Adds a control to the collection of controls to anchor
' Arguments:    objControl - The control to anchor
'               eAnchor - The anchor type for the control
' Returns:      None
' Date:         18/07/2001
' Author:       Toby Wicks
'
    Dim tControl As typControl
    
    If UBound(pvControls) = 0 Then
        ReDim pvControls(1 To 1) As typControl
    Else
        ReDim Preserve pvControls(1 To UBound(pvControls) + 1) As typControl
    End If
    
    With tControl
        Set .Control = objControl
        .Anchor = eAnchor
        .Top = objControl.Top
        .Left = objControl.Left
        .Bottom = pfrmForm.Height - objControl.Top - objControl.Height
        .Right = pfrmForm.Width - objControl.Left - objControl.Width
    End With

    pvControls(UBound(pvControls)) = tControl

End Sub

Private Sub pfrmForm_Resize()
'
' Purpose:      Resizes all of the controls in the collection
' Arguments:    None
' Returns:      None
' Date:         18/07/2001
' Author:       Toby Wicks
'

    Dim intIndex As Integer

    For intIndex = 1 To UBound(pvControls)
        RepositionControl pvControls(intIndex)
    Next
    
End Sub

Private Sub RepositionControl(ByRef tControl As typControl)
'
' Purpose:      Repositions the control based on the anchor type
'               and it's initial position
' Arguments:    tControl - The type containing the control data
' Returns:      None
' Date:         18/07/2001
' Author:       Toby Wicks
'

    Dim eAnchor As ControlAnchorEnum
    Dim objControl As Control

    Set objControl = tControl.Control
    eAnchor = tControl.Anchor

    On Error Resume Next

    If eAnchor And pcintRightSide Then
        If eAnchor And pcintLeftSide Then
            'Debug.Print tControl.Control.Name & " change width: " & objControl.Width
            objControl.Width = pfrmForm.Width - tControl.Left - tControl.Right
        Else
            objControl.Left = pfrmForm.Width - objControl.Width - tControl.Right
        End If
    End If

    If eAnchor And pcintBottomSide Then
        If eAnchor And pcintTopSide Then
            objControl.Height = pfrmForm.Height - tControl.Top - tControl.Bottom
        Else
            objControl.Top = pfrmForm.Height - objControl.Height - tControl.Bottom
        End If
    End If

End Sub
