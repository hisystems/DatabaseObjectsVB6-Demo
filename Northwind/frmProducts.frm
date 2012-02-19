VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProducts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Products"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   Icon            =   "frmProducts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwProducts 
      Height          =   4875
      Left            =   2940
      TabIndex        =   1
      Top             =   120
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   8599
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
   Begin MSComctlLib.TreeView tvwCategories 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   8493
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   435
      Left            =   8520
      TabIndex        =   2
      Top             =   5100
      Width           =   1335
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadCategories()

    Dim intIndex As Integer
    Dim objRoot As Node
    Dim objCategories As Categories

    With Me.tvwCategories.Nodes
        Set objRoot = .Add(, , , "All")
        objRoot.Expanded = True
        Set objCategories = NorthwindDB.Categories
        For intIndex = 1 To objCategories.Count
            .Add objRoot, tvwChild, , objCategories(intIndex).Name
        Next
    End With

End Sub

Private Sub LoadProducts( _
    ByVal objProducts As Products)

    Dim objProduct As Product
    
    With Me.lvwProducts.ListItems
        .Clear
        For Each objProduct In objProducts
            With .Add(, , objProduct.Name)
                .SubItems(1) = objProduct.Supplier.Name
                .SubItems(2) = FormatCurrency(objProduct.UnitPrice)
                .SubItems(3) = objProduct.QuantityPerUnit
                'Debug.Print objProduct.Category2.Name
            End With
        Next
    End With

End Sub

Private Sub Form_Load()

    LoadCategories
    LoadProducts NorthwindDB.Products
    
End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub tvwCategories_NodeClick(ByVal Node As MSComctlLib.Node)

    If Node.Parent Is Nothing Then
        LoadProducts NorthwindDB.Products
    Else
        LoadProducts NorthwindDB.Categories(Node.Text).Products
    End If

End Sub
