Attribute VB_Name = "Data"
Option Explicit

Public gobjNorthwindInstance As NorthwindDatabase

Public gstrConnectionString As String

Public Function CreateConnection() As ADODB.Connection

    Dim objConnection As ADODB.Connection
    Set objConnection = New ADODB.Connection
    
    objConnection.ConnectionString = gstrConnectionString
    
    Set CreateConnection = objConnection

End Function

Public Property Get Products() As Products
'It is not necessary to keep a static variable of the Products instance here
'but in some particular circumstances it is useful to only have
'one instance created.
    
    Set Products = New Products
    
End Property

Public Property Get Product(ByVal lngID As Long) As Product

    Set Product = dbo.Object(Data.Products, lngID)
    
End Property

Public Property Get Orders() As Orders

    Set Orders = New Orders
    
End Property

Public Property Get Suppliers() As Suppliers

    Set Suppliers = New Suppliers
    
End Property

Public Property Get Supplier(ByVal lngID As Long) As Supplier

    Set Supplier = dbo.Object(Data.Suppliers, lngID)

End Property

Public Property Get Categories() As Categories

    Set Categories = New Categories
    
End Property

Public Property Get Category(ByVal lngID As Long) As Category

    Set Category = dbo.Object(Data.Categories, lngID)

End Property

