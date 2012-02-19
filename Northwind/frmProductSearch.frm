VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProductSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Search"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   Icon            =   "frmProductSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   435
      Left            =   5400
      TabIndex        =   7
      Top             =   1140
      Width           =   1395
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
      Height          =   1635
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   6855
      Begin VB.PictureBox picOptionsContainer 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   4275
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   840
         Width           =   4275
         Begin VB.OptionButton optProductDescription 
            Caption         =   "Product Name"
            Height          =   195
            Left            =   60
            TabIndex        =   6
            Top             =   360
            Width           =   2715
         End
         Begin VB.OptionButton optProductDescriptionPrefix 
            Caption         =   "Product Name Prefix"
            Height          =   195
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Value           =   -1  'True
            Width           =   2775
         End
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   300
         TabIndex        =   0
         Top             =   360
         Width           =   6315
      End
   End
   Begin MSComctlLib.ListView lvwProducts 
      Height          =   4935
      Left            =   180
      TabIndex        =   1
      Top             =   1920
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8705
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
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   435
      Left            =   5700
      TabIndex        =   2
      Top             =   7020
      Width           =   1335
   End
End
Attribute VB_Name = "frmProductSearch"
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
            End With
        Next
    End With

End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdSearch_Click()
    
    Dim colProducts As Collection
    
    'The Search function returns a collection of Product objects that match the search criteria
    Set colProducts = northwinddb.Products.Search(Me.txtSearch.Text, IIf(Me.optProductDescriptionPrefix.Value, dbProductSearchDescriptionPrefix, dbProductSearchDescription))
        
    If colProducts.Count = 0 Then
        Me.lvwProducts.ListItems.Clear
        MsgBox "There were no products that matched the specified criteria.", vbOKOnly + vbInformation
    Else
        LoadProducts colProducts
    End If
    
End Sub
