VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   Caption         =   "Database Objects Demonstration"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdProductSearchExtended 
      Caption         =   "Search Products (Extended)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8940
      TabIndex        =   4
      Top             =   3720
      Width           =   1695
   End
   Begin VB.PictureBox picContainer 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7605
      Left            =   0
      ScaleHeight     =   7605
      ScaleWidth      =   8580
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   8580
      Begin SHDocVwCtl.WebBrowser wbMain 
         Height          =   6195
         Left            =   0
         TabIndex        =   11
         Top             =   1380
         Width           =   8595
         ExtentX         =   15161
         ExtentY         =   10927
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.Label lblGeneralIntroduction 
         BackStyle       =   0  'Transparent
         Caption         =   "Quick Start"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2340
         MouseIcon       =   "frmMain.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   1080
         X2              =   1080
         Y1              =   1020
         Y2              =   1320
      End
      Begin VB.Label lblBack 
         BackStyle       =   0  'Transparent
         Caption         =   "< Back"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         MouseIcon       =   "frmMain.frx":015E
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblDocumentation 
         BackStyle       =   0  'Transparent
         Caption         =   "Documentation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3600
         MouseIcon       =   "frmMain.frx":02B0
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblWelcome 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1320
         MouseIcon       =   "frmMain.frx":0402
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.Image imgLogo 
         Height          =   930
         Left            =   7500
         Picture         =   "frmMain.frx":0554
         Top             =   0
         Width           =   915
      End
      Begin VB.Line lnTitleUnderline 
         BorderColor     =   &H00C0C0C0&
         X1              =   240
         X2              =   4380
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblDemonstration 
         BackStyle       =   0  'Transparent
         Caption         =   "Demonstration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   2940
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Database Objects"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   5355
      End
   End
   Begin VB.CommandButton cmdProductSearch 
      Caption         =   "Search Products"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8940
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "View An Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8940
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdProducts 
      Caption         =   "View Products"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8940
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdSuppliers 
      Caption         =   "View Suppliers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8940
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Frame fraBorder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7995
      Left            =   8235
      TabIndex        =   8
      Top             =   -300
      Width           =   390
   End
   Begin VB.Label lblExamples 
      Caption         =   "Examples"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8940
      TabIndex        =   14
      Top             =   360
      Width           =   1515
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pobjControlAnchor As ControlAnchor

Private Sub Form_Load()

    Const cstrAccessFilePath As String = "C:\Program Files\Microsoft Visual Studio\VB98\nwind.mdb"

    If Dir$(cstrAccessFilePath) = vbNullString Then
        If MsgBox("Could not connect to the Northwind Access database at '" & cstrAccessFilePath & "', would you like to download the database from the Microsoft website?", vbYesNo + vbQuestion) = vbYes Then
            Me.wbMain.Navigate "http://www.microsoft.com/downloads/details.aspx?FamilyID=C6661372-8DBE-422B-8676-C632D66C529C&displaylang=EN"
        End If
    Else
        northwinddb.Connect_MicrosoftAccess cstrAccessFilePath
        'northwinddb.Connect_SQLServer "(local)", "Northwind"
        'northwinddb.Connect_MySQL "localhost", "northwind"
        lblWelcome_Click
    End If

    Set pobjControlAnchor = New ControlAnchor
    With pobjControlAnchor
        .Initialize Me
        .AddControl Me.cmdSuppliers, ccAnchorRight
        .AddControl Me.cmdProducts, ccAnchorRight
        .AddControl Me.cmdProductSearch, ccAnchorRight
        .AddControl Me.cmdProductSearchExtended, ccAnchorRight
        .AddControl Me.cmdOrder, ccAnchorRight
        .AddControl Me.fraBorder, ccAnchorTopRightBottom
        .AddControl Me.imgLogo, ccAnchorRight
        .AddControl Me.picContainer, ccAnchorAll
        .AddControl Me.wbMain, ccAnchorAll
        .AddControl Me.lblExamples, ccAnchorRight
    End With

End Sub

Private Sub cmdOrder_Click()
    
    frmOrder.Show vbModal, Me
    
End Sub

Private Sub cmdProducts_Click()
    
    frmProducts.Show vbModal, Me
    
End Sub

Private Sub cmdProductSearch_Click()

    frmProductSearch.Show vbModal, Me

End Sub

Private Sub cmdProductSearchExtended_Click()

    frmProductSearchExtended.Show vbModal, Me

End Sub

Private Sub cmdSuppliers_Click()

    frmSuppliers.Show vbModal, Me

End Sub

Private Sub lblBack_Click()

    Me.wbMain.GoBack
    
End Sub

Private Sub lblDocumentation_Click()

    Me.wbMain.Navigate "http://www.hisystems.com.au/databaseobjects/reference_vb6.htm"

End Sub

Private Sub lblGeneralIntroduction_Click()

    Me.wbMain.Navigate "http://www.hisystems.com.au/databaseobjects/DatabaseObjects_QuickStart_vb6.htm"

End Sub

Private Sub lblWelcome_Click()

    Me.wbMain.Navigate App.Path & "\welcome.html"

End Sub

Private Sub wbMain_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)

    If Command = CSC_NAVIGATEBACK Then
        Me.lblBack.Enabled = Enable
    End If

End Sub

Private Sub wbMain_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    On Error Resume Next
    Me.wbMain.Document.body.Style.BorderStyle = "none"

End Sub
