VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   Icon            =   "frmOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEditQuantity 
      Caption         =   "Edit Quantity"
      Height          =   435
      Left            =   360
      TabIndex        =   24
      Top             =   6480
      Width           =   1515
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   435
      Left            =   180
      TabIndex        =   23
      Top             =   7260
      Width           =   1335
   End
   Begin VB.TextBox txtCost 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   600
      Width           =   1455
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   180
      TabIndex        =   19
      Top             =   3540
      Width           =   7635
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   2535
         Left            =   180
         TabIndex        =   20
         Top             =   300
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Quantity"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Unit Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cost"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraShipTo 
      Caption         =   "Ship To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   180
      TabIndex        =   8
      Top             =   1080
      Width           =   7635
      Begin VB.TextBox txtShipName 
         Height          =   315
         Left            =   1320
         TabIndex        =   17
         Top             =   300
         Width           =   3315
      End
      Begin VB.TextBox txtPostalCode 
         Height          =   315
         Left            =   1320
         TabIndex        =   15
         Top             =   1740
         Width           =   1035
      End
      Begin VB.TextBox txtRegion 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   1380
         Width           =   1035
      End
      Begin VB.TextBox txtCity 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   1020
         Width           =   1695
      End
      Begin VB.TextBox txtAddress 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   660
         Width           =   3315
      End
      Begin VB.Label lblShipName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblPostalCode 
         Caption         =   "Postal Code:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label lblRegion 
         Caption         =   "Region:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblCity 
         Caption         =   "City:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.TextBox txtOrdered 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   180
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpRequired 
      Height          =   315
      Left            =   1500
      TabIndex        =   4
      Top             =   540
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Format          =   50855937
      CurrentDate     =   38180
   End
   Begin VB.TextBox txtOrderNumber 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   180
      Width           =   1755
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   5040
      TabIndex        =   1
      Top             =   7260
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   6480
      TabIndex        =   0
      Top             =   7260
      Width           =   1335
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Cost:"
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   660
      Width           =   1155
   End
   Begin VB.Label lblOrdered 
      Alignment       =   1  'Right Justify
      Caption         =   "Ordered:"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label lblRequired 
      Caption         =   "Required:"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label lblOrderNumber 
      Caption         =   "Order Number:"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   240
      Width           =   1155
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pobjOrder As Order

Private Sub LoadOrder( _
    ByVal objOrder As Order)

    Dim intIndex As Integer
    Dim objOrderDetail As OrderDetail

    Me.txtOrderNumber.Text = objOrder.Number
    Me.dtpRequired.Value = objOrder.RequiredDate
    Me.txtOrdered.Text = FormatDateTime(objOrder.TheDate, vbShortDate)
    
    Me.txtShipName.Text = objOrder.ShipName
    
    Me.txtAddress.Text = objOrder.ShipAddress
    Me.txtCity.Text = objOrder.ShipCity
    Me.txtPostalCode.Text = objOrder.ShipPostalCode
    Me.txtRegion.Text = objOrder.ShipRegion
    
    With Me.lvwDetails.ListItems
        For intIndex = 1 To objOrder.Details.Count
            Set objOrderDetail = objOrder.Details(intIndex)
            LoadOrderDetail objOrderDetail, .Add(, , objOrderDetail.Product.Name)
        Next
    End With
    
    RefreshTotals
    
End Sub

Private Sub LoadOrderDetail( _
    ByVal objOrderDetail As OrderDetail, _
    ByVal objItem As ListItem)
    
    objItem.SubItems(1) = objOrderDetail.Quantity
    objItem.SubItems(2) = FormatCurrency(objOrderDetail.UnitPrice)
    objItem.SubItems(3) = FormatCurrency(objOrderDetail.Cost)
    
End Sub

Private Sub cmdDelete_Click()

    If MsgBox("Are you sure you want to delete this order.", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        northwinddb.Orders.Delete pobjOrder
        Unload Me
    End If

End Sub

Private Sub cmdEditQuantity_Click()
    
    EditQuantity

End Sub

Private Sub cmdOK_Click()

    pobjOrder.RequiredDate = Me.dtpRequired.Value
    
    pobjOrder.ShipName = Me.txtShipName.Text
    pobjOrder.ShipAddress = Me.txtAddress.Text
    pobjOrder.ShipCity = Me.txtCity.Text
    pobjOrder.ShipPostalCode = Me.txtPostalCode.Text
    pobjOrder.ShipRegion = Me.txtRegion.Text
    
    pobjOrder.Save
    Unload Me

End Sub

Private Sub Form_Load()

    'load the first order
    Set pobjOrder = northwinddb.Orders(5)
    LoadOrder pobjOrder

End Sub

Private Sub LoadProducts( _
    ByVal cboBox As ComboBox)
    
    Dim intIndex As Integer
    
    With cboBox
        .Clear
        For intIndex = 1 To northwinddb.Products.Count
            .AddItem northwinddb.Products(intIndex).Name
        Next
        .ListIndex = 0
    End With
    
End Sub

Private Sub EditQuantity()

    Dim objItem As ListItem
    Dim objOrderDetail As OrderDetail
    Dim strQuantity As String
    
    Set objItem = Me.lvwDetails.SelectedItem
    
    If Not objItem Is Nothing Then
        Set objOrderDetail = pobjOrder.Details(objItem.Index)
        
        strQuantity = InputBox("Please enter the new quantity for the '" & objOrderDetail.Product.Name & "' product.", , objOrderDetail.Quantity)
        strQuantity = Trim$(strQuantity)
        
        If strQuantity <> vbNullString Then
            If Val(strQuantity) > 0 And Val(strQuantity) < 32000 Then
                'update the detail lines on the fly
                objOrderDetail.Quantity = Val(strQuantity)
                LoadOrderDetail objOrderDetail, objItem
                RefreshTotals
            End If
        End If
    Else
        MsgBox "Please select a detail line to edit.", vbOKOnly + vbInformation
    End If

End Sub

Private Sub RefreshTotals()

    Me.txtCost.Text = FormatCurrency(pobjOrder.Cost)

End Sub

Private Sub lvwDetails_DblClick()

    EditQuantity

End Sub

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub
