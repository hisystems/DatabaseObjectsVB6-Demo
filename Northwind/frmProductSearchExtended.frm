VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProductSearchExtended 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Search Extended"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9525
   Icon            =   "frmProductSearchExtended.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   435
      Left            =   7800
      TabIndex        =   11
      Top             =   1500
      Width           =   1395
   End
   Begin VB.ComboBox cboOnOrder 
      Height          =   315
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   180
      TabIndex        =   10
      Top             =   120
      Width           =   9195
      Begin VB.ComboBox cboDiscontinued 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cboInStock 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtProductName 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblDiscontinued 
         Caption         =   "Discontinued"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label lblOnOrder 
         Caption         =   "On Order"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblInStock 
         Caption         =   "In Stock"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label lblProductName 
         Caption         =   "Product Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   0
         Top             =   420
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lvwProducts 
      Height          =   4575
      Left            =   180
      TabIndex        =   8
      Top             =   2280
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   4366
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Supplier"
         Object.Width           =   2699
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Unit Price"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Quantity Per Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "In Stock"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "On Order"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Discontinued"
         Object.Width           =   2028
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   435
      Left            =   8040
      TabIndex        =   9
      Top             =   7020
      Width           =   1335
   End
End
Attribute VB_Name = "frmProductSearchExtended"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadProducts( _
    ByVal colProducts As Collection)

    Dim objProduct As Product
    
    With Me.lvwProducts.ListItems
        .Clear
        For Each objProduct In colProducts
            With .Add(, , objProduct.Name)
                .SubItems(1) = objProduct.Supplier.Name
                .SubItems(2) = FormatCurrency(objProduct.UnitPrice)
                .SubItems(3) = objProduct.QuantityPerUnit
                .SubItems(4) = ConvertBooleanToYesNo(objProduct.IsInStock)
                .SubItems(5) = ConvertBooleanToYesNo(objProduct.IsOnOrder)
                .SubItems(6) = ConvertBooleanToYesNo(objProduct.Discontinued)
            End With
        Next
    End With

End Sub

Private Sub cmdSearch_Click()
    
    Dim colProducts As Collection
    Dim objSearch As ProductSearch
    Set objSearch = New ProductSearch
    
    objSearch.Name = Me.txtProductName.Text
    
    If Me.cboInStock.Text <> vbNullString Then
        objSearch.InStock = ConvertYesNoToBoolean(Me.cboInStock.Text)
    End If
    
    If Me.cboOnOrder.Text <> vbNullString Then
        objSearch.OnOrder = ConvertYesNoToBoolean(Me.cboOnOrder.Text)
    End If
    
    If Me.cboDiscontinued.Text <> vbNullString Then
        objSearch.Discontinued = ConvertYesNoToBoolean(Me.cboDiscontinued.Text)
    End If
    
    'The Search function returns a collection of Product objects that match the search criteria
    Set colProducts = objSearch.Search
        
    If colProducts.Count = 0 Then
        Me.lvwProducts.ListItems.Clear
        MsgBox "There were no products that matched the specified criteria.", vbOKOnly + vbInformation
    Else
        LoadProducts colProducts
    End If
    
End Sub

Private Sub LoadYesNoComboBox( _
    ByVal cboBox As ComboBox)

    With cboBox
        .AddItem vbNullString
        .AddItem ConvertBooleanToYesNo(True)
        .AddItem ConvertBooleanToYesNo(False)
        .ListIndex = 0
    End With

End Sub

Private Function ConvertBooleanToYesNo( _
    ByVal bValue As Boolean) As String

    If bValue Then
        ConvertBooleanToYesNo = "Yes"
    Else
        ConvertBooleanToYesNo = "No"
    End If
    
End Function

Private Function ConvertYesNoToBoolean( _
    ByVal strText As String) As Boolean
    
    If StrComp(strText, "yes", vbTextCompare) = 0 Then
        ConvertYesNoToBoolean = True
    Else
        ConvertYesNoToBoolean = False
    End If
    
End Function

Private Sub Form_Load()
    
    LoadYesNoComboBox Me.cboInStock
    LoadYesNoComboBox Me.cboOnOrder
    LoadYesNoComboBox Me.cboDiscontinued

End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub
