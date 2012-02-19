VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NorthwindDatabase"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Connect_SQLServer( _
    ByVal strDataSource As String, _
    ByVal strDatabaseName As String)

    If Trim$(strDatabaseName) = vbNullString Then
        dbo.RaiseError dboErrorGeneral, "Null database name."
    End If
    
    gstrConnectionString = _
        "Provider=SQLOLEDB;Data Source=" & strDataSource & ";" & _
        "Database=" & strDatabaseName & ";UID=sa;pwd=;"

    dbo.Connect gstrConnectionString, dboConnectionTypeSQLServer
    
End Sub

Public Sub Connect_MicrosoftAccess( _
    ByVal strDatabaseFilePath As String)

    If Trim$(strDatabaseFilePath) = vbNullString Then
        dbo.RaiseError dboErrorGeneral, "Null database file path."
    End If
        
    gstrConnectionString = _
        "Data Source=" & strDatabaseFilePath & ";" & _
        "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=;"

    'Maintaining an open connection at all times keeps the connection pool alive
    Static objDummyConnection As ADODB.Connection
    Set objDummyConnection = New Connection
    objDummyConnection.Open gstrConnectionString

    dbo.Connect gstrConnectionString, dboConnectionTypeMicrosoftAccess

End Sub

Public Sub Connect_MySQL( _
    ByVal strDataSource As String, _
    ByVal strDatabaseName As String)

    If Trim$(strDatabaseName) = vbNullString Then
        dbo.RaiseError dboErrorGeneral, "Null database name."
    End If
    
    gstrConnectionString = _
        "Driver={MySQL}; Server=" & strDataSource & "; Database=" & strDatabaseName & "; UID=root; PWD=;"
        
    dbo.Connect gstrConnectionString, dboConnectionTypeMySQL

End Sub

Private Sub Class_Initialize()

    If gobjNorthwindInstance Is Nothing Then
        Set Data.gobjNorthwindInstance = Me
    Else
        dbo.RaiseError dboErrorGeneral, _
            "Cannot create multiple instances of this global multi-use class because this class automatically disconnects from the database on termination." & vbCrLf & _
            "Use the 'NorthwindDB.' reference which automatically creates a single global instance of this class."
    End If
    
End Sub

Public Property Get Products() As Products
    Set Products = Data.Products
End Property

Public Property Get Suppliers() As Suppliers
    Set Suppliers = Data.Suppliers
End Property

Public Property Get Categories() As Categories
    Set Categories = Data.Categories
End Property

Public Property Get Orders() As Orders
    Set Orders = Data.Orders
End Property
