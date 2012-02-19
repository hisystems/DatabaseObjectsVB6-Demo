VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProductSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This class provides the ability to search on multiple fields on multiple criteria
'for a product, as opposed to the Product.Search function which only provides searching
'on 1 field.

'This class also demonstrates how to optimise the objects so that rather than loading
'referenced objects (in this case the product's Supplier object) via multiple SQL calls
'all of the associated tables are joined and returned in one recordset. Have a look in the
'Product's IDatabaseObject_Load to see how the objects are loaded from the same
'recordset, and in particular the ObjectFromFieldValues function.

Private pstrName As String
Private peInStock As VbTriState
Private peOnOrder As VbTriState
Private peDiscontinued As VbTriState

Public Property Let Name(ByVal strValue As String)

    pstrName = Trim$(strValue)
    
End Property

Public Property Let InStock(ByVal bValue As Boolean)

    peInStock = bValue
    
End Property

Public Property Let OnOrder(ByVal bValue As Boolean)

    peOnOrder = bValue
    
End Property

Public Property Let Discontinued(ByVal bValue As Boolean)

    peDiscontinued = bValue
    
End Property

Private Sub Class_Initialize()

    peInStock = vbUseDefault
    peOnOrder = vbUseDefault
    peDiscontinued = vbUseDefault
    
End Sub

Public Function Search() As Collection
    
    Dim objConnection As ADODB.Connection
    Dim objProducts As IDatabaseObjects
    Dim lngIndex As Long
    Dim objProduct As Product
    Dim rsResults As Recordset
    Dim colResults As Collection
    Set colResults = New Collection
    
    Dim objSelect As SQLSelect
    Set objSelect = New SQLSelect
    
    'The TableJoins function will join the Products and Suppliers tables so that the
    'results are returned in one recordset, rather than requiring multiple statements to read
    'each product's associated supplier record. See the Products' IDatabaseObjects_TableJoins
    'on how this is done.
    
    Set objProducts = Data.Products
    'Always add the primary table before calling the TableJoins function
    objSelect.Tables.Add objProducts.TableName
    Set objSelect.Tables.Joins = objProducts.TableJoins(objSelect.Tables(1), objSelect.Tables)
    
    'search for the product name if it has been set
    If pstrName <> vbNullString Then
        objSelect.Where.Add "ProductName", dboComparisonLike, "%" & pstrName & "%"
    End If
    
    If peInStock <> vbUseDefault Then
        If peInStock = vbTrue Then
            'select products that are in stock
            objSelect.Where.Add "UnitsInStock", dboComparisonGreaterThan, 0
        Else
            'select products that are not in stock
            objSelect.Where.Add "UnitsInStock", dboComparisonEqualTo, 0
        End If
    End If
    
    If peOnOrder <> vbUseDefault Then
        If peOnOrder = vbTrue Then
            'select products that are on order
            objSelect.Where.Add "UnitsOnOrder", dboComparisonGreaterThan, 0
        Else
            'select products that are not on order
            objSelect.Where.Add "UnitsOnOrder", dboComparisonEqualTo, 0
        End If
    End If
    
    If peDiscontinued <> vbUseDefault Then
        objSelect.Where.Add "Discontinued", dboComparisonEqualTo, CBool(peDiscontinued)
    End If
    
    Debug.Print objSelect.SQL
    Set objConnection = CreateConnection
    objConnection.Open
    
    Set rsResults = objConnection.Execute(objSelect.SQL)

    While Not rsResults.EOF
        Set objProduct = dbo.ObjectFromRecordset(Data.Products, rsResults)
        colResults.Add objProduct
        rsResults.MoveNext
    Wend

    Set Search = colResults
    
    rsResults.Close
    objConnection.Close

End Function
