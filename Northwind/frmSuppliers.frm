VERSION 5.00
Begin VB.Form frmSuppliers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suppliers"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   Icon            =   "frmSuppliers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   435
      Left            =   2820
      TabIndex        =   4
      Top             =   5820
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   435
      Left            =   180
      TabIndex        =   1
      Top             =   5820
      Width           =   1215
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "&Rename"
      Height          =   435
      Left            =   1500
      TabIndex        =   2
      Top             =   5820
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   3
      Top             =   5820
      Width           =   1215
   End
   Begin VB.ListBox lstSuppliers 
      Height          =   5520
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5955
   End
End
Attribute VB_Name = "frmSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadSuppliers()

    Dim intIndex As Integer
    Dim objSupplier As Supplier

    With Me.lstSuppliers
        .Clear
        
        'To use the For Each enumerator see the Enumerator function in the Suppliers class
        For Each objSupplier In northwinddb.Suppliers
            .AddItem objSupplier.Name
        Next
        
        'Alternatively you can interate using the
        'For intIndex = 1 To northwinddb.Suppliers.Count
        '    .AddItem northwinddb.Suppliers(intIndex).Name
        'Next
        '.AddItem northwinddb.Suppliers(20).Name
        '.AddItem northwinddb.Suppliers(2).Name
        '.AddItem northwinddb.Suppliers(4).Name
    End With

End Sub

Private Sub cmdNew_Click()

    Dim objSupplier As Supplier
    Dim strNewSupplierName As String
    
    strNewSupplierName = Trim$(InputBox("Please enter the name of the new supplier:"))
        
    If strNewSupplierName <> vbNullString Then
        If northwinddb.Suppliers.Exists(strNewSupplierName) Then
            MsgBox "Supplier '" & strNewSupplierName & "' already exists.", vbOKOnly + vbInformation
        Else
            Set objSupplier = New Supplier
            objSupplier.Name = strNewSupplierName
            objSupplier.Save
            LoadSuppliers
        End If
    End If

End Sub

Private Sub cmdRename_Click()

    Dim strNewSupplierName As String
    Dim objSupplier As Supplier

    If Me.lstSuppliers.ListIndex >= 0 Then
        Set objSupplier = northwinddb.Suppliers(Me.lstSuppliers.Text)
        strNewSupplierName = Trim$(InputBox("Please enter the new supplier name for '" & objSupplier.Name & "'.", , objSupplier.Name))
        If strNewSupplierName <> vbNullString Then
            'if the new name is the same as the old name then do nothing
            If StrComp(strNewSupplierName, objSupplier.Name, vbTextCompare) = 0 Then
                
            ElseIf northwinddb.Suppliers.Exists(strNewSupplierName) Then
                MsgBox "This supplier name already exists.", vbOKOnly + vbInformation
            Else
                objSupplier.Name = strNewSupplierName
                objSupplier.Save
                LoadSuppliers
            End If
        End If
    Else
        MsgBox "Please select a supplier to rename.", vbOKOnly + vbInformation
    End If

End Sub

Private Sub cmdDelete_Click()

    Dim objSupplier As Supplier

    If Me.lstSuppliers.ListIndex >= 0 Then
        Set objSupplier = northwinddb.Suppliers(Me.lstSuppliers.Text)
        If objSupplier.IsDeletable Then
            If MsgBox("Are you sure you want to delete supplier '" & objSupplier.Name & "'.", vbYesNo + vbQuestion) = vbYes Then
                northwinddb.Suppliers.Delete objSupplier
                LoadSuppliers
            End If
        Else
            MsgBox "Supplier '" & objSupplier.Name & "' cannot be deleted, because it is a supplier of an existing product.", vbOKOnly + vbInformation
        End If
    Else
        MsgBox "Please select a supplier to delete.", vbOKOnly + vbInformation
    End If

End Sub

Private Sub Form_Load()

    LoadSuppliers
    
End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub
