VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Products"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDatabaseObjects
'Allow other DLL projects to load a product using it's database ID value
Implements IGlobalDatabaseObjects

Public Enum ProductSearchTypeEnum
    dbProductSearchDescriptionPrefix
    dbProductSearchDescription
End Enum

Private pobjCategory As Category

Friend Property Set Category(ByVal objValue As Category)
'Friend will allow public access within the NorthwindDB project but will be
'private outside of the NorthwindDB project

    Set pobjCategory = objValue

End Property

Public Function Add( _
    Optional ByVal strProductName As String) As Product

    Set Add = New Product
    Add.Name = strProductName

End Function

Public Property Get Item(ByVal vIndex As Variant) As Product
Attribute Item.VB_UserMemId = 0
'This property has been set to a default property (Tools -> Procedure Attributes,
'select the 'Item' property, click the 'Advanced >>' button and change
'the 'Procedure ID' to '(Default)') so that rather than accessing a Product by specifying
'NorthwindDB.Products.Item("AProductName") - you can specify NorthwindDB.Products("AProductName")
'to return a Product object.
    
    Set Item = DBO.ObjectItem(Me, vIndex)
    
End Property

Public Function Search( _
    ByVal strSearch As String, _
    ByVal eType As ProductSearchTypeEnum) As Collection

    Dim objConditions As SQLConditions
    Set objConditions = New SQLConditions

    Select Case eType
        Case dbProductSearchDescriptionPrefix
            objConditions.Add "ProductName", dboComparisonLike, strSearch & "%"
        Case dbProductSearchDescription
            objConditions.Add "ProductName", dboComparisonLike, "%" & strSearch & "%"
    End Select
    
    Set Search = DBO.ObjectsSearch(Me, objConditions)

End Function

Public Sub Delete(ByRef objProduct As Product)

    DBO.ObjectDelete Me, objProduct

End Sub

Public Function Exists(ByVal strProductName As String) As Boolean

    Exists = DBO.ObjectExists(Me, strProductName)

End Function

Public Property Get Count() As Long

    Count = DBO.ObjectsCount(Me)

End Property

Public Property Get Enumerator() As IUnknown
Attribute Enumerator.VB_UserMemId = -4

    Static colProducts As Collection
    
    Set colProducts = DBO.ObjectsCollection(Me)
    Set Enumerator = colProducts.[_NewEnum]

End Property

Private Function IDatabaseObjects_DistinctFieldAutoIncrements() As Boolean

    IDatabaseObjects_DistinctFieldAutoIncrements = True

End Function

Private Function IDatabaseObjects_DistinctFieldName() As String

    IDatabaseObjects_DistinctFieldName = "ProductID"

End Function

Private Function IDatabaseObjects_ItemInstance() As DBO.IDatabaseObject
    
    Set IDatabaseObjects_ItemInstance = New Product
    
End Function

Private Function IDatabaseObjects_KeyFieldName() As String

    IDatabaseObjects_KeyFieldName = "ProductName"

End Function

Private Function IDatabaseObjects_OrderBy() As DBO.SQLSelectOrderByFields

    Set IDatabaseObjects_OrderBy = New SQLSelectOrderByFields
    IDatabaseObjects_OrderBy.Add "ProductName", dboOrderAscending
    
End Function

Private Function IDatabaseObjects_Subset() As DBO.SQLConditions
    
    'if the category filter has been set then this collection should only
    'contain products within the specified category. The Item, Count, Delete and Exists
    'functions will all reflect this change
    If Not pobjCategory Is Nothing Then
        Set IDatabaseObjects_Subset = New SQLConditions
        IDatabaseObjects_Subset.Add "CategoryID", dboComparisonEqualTo, pobjCategory.ID
    End If

End Function

Private Function IDatabaseObjects_TableJoins(ByVal objPrimaryTable As SQLSelectTable, ByVal objTables As DBO.SQLSelectTables) As DBO.SQLSelectTableJoins
    
    'Implementing this function is optional, but is useful when attempting to optimise loading speeds.
    'This function is used by the ObjectsCollection, Object, ObjectItem, ObjectByKey, ObjectOrdinal and ObjectSearch functions.
    'If this function has been implemented the ObjectsSearch conditions can also search fields in the joined table(s).
    'In this example, Products table will always be joined with the Supplier table. We could also join the Products
    'table to the Category table, however the Product.Category property is not used often enough to warrant
    'always joining the category table whenever loading a product. Of course, you can always join different
    'tables in different situations, for example you might want join to other tables when searching and to
    'not join to other tables in normal circumstances.

    Dim objTableJoins As SQLSelectTableJoins
    Set objTableJoins = New SQLSelectTableJoins
    
    With objTableJoins.Add(objPrimaryTable, dboInnerJoin, objTables.Add("Suppliers"))
        .Where.Add "SupplierID", dboComparisonEqualTo, "SupplierID"
    End With
    
    'With objTableJoins.Add(objPrimaryTable, dboInnerJoin, objTables.Add("Categories"))
    '    .Where.Add "CategoryID", dboComparisonEqualTo, "CategoryID"
    'End With
    
    Set IDatabaseObjects_TableJoins = objTableJoins

End Function

Private Function IDatabaseObjects_TableName() As String

    IDatabaseObjects_TableName = "Products"
    
End Function

Private Function IGlobalDatabaseObjects_Object(ByVal vDistinctValue As Variant) As DBO.IDatabaseObject

    Set IGlobalDatabaseObjects_Object = Data.Product(vDistinctValue)

End Function
