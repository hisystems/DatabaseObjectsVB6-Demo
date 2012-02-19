VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Suppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDatabaseObjects

Public Property Get Item(ByVal vIndex As Variant) As Supplier
Attribute Item.VB_UserMemId = 0

    Set Item = dbo.ObjectItem(Me, vIndex)

End Property

Public Property Get Count() As Long

    Count = dbo.ObjectsCount(Me)

End Property

Public Property Get Enumerator() As IUnknown
Attribute Enumerator.VB_UserMemId = -4
'1. Select Tools > Procedure Attributes
'2. Select Enumerator in the Name list
'3. Click Advanced
'4. Enter -4 for the 'Procedure ID'

    Static colSuppliers As Collection
    
    Set colSuppliers = dbo.ObjectsCollection(Me)
    Set Enumerator = colSuppliers.[_NewEnum]
  
End Property

Public Function Exists(ByVal strName As String) As Boolean

    Exists = dbo.ObjectExists(Me, strName)

End Function

Public Sub Delete(ByRef objSupplier As Supplier)

    If objSupplier.IsDeletable Then
        dbo.ObjectDelete Me, objSupplier
    Else
        dbo.RaiseError dboErrorObjectNotDeletable, objSupplier.Name
    End If

End Sub

Private Function IDatabaseObjects_DistinctFieldAutoIncrements() As Boolean

    IDatabaseObjects_DistinctFieldAutoIncrements = True
    
End Function

Private Function IDatabaseObjects_DistinctFieldName() As String

    IDatabaseObjects_DistinctFieldName = "SupplierID"
    
End Function

Private Function IDatabaseObjects_ItemInstance() As dbo.IDatabaseObject

    Set IDatabaseObjects_ItemInstance = New Supplier
    
End Function

Private Function IDatabaseObjects_KeyFieldName() As String

    IDatabaseObjects_KeyFieldName = "CompanyName"
    
End Function

Private Function IDatabaseObjects_OrderBy() As dbo.SQLSelectOrderByFields
    
    Set IDatabaseObjects_OrderBy = New SQLSelectOrderByFields
    IDatabaseObjects_OrderBy.Add "CompanyName"
    
End Function

Private Function IDatabaseObjects_Subset() As dbo.SQLConditions

'    Dim objConditions As SQLConditions
'    Set objConditions = New SQLConditions
'
'    objConditions.Add "CompanyName", dboComparisonLike, "M%"
'    objConditions.AddLogicalOperator dboLogicalOr
'    objConditions.Add "CompanyName", dboComparisonLike, "N%"
'
'    Set IDatabaseObjects_Subset = objConditions
        
End Function

Private Function IDatabaseObjects_TableJoins(ByVal objPrimaryTable As SQLSelectTable, ByVal objTables As dbo.SQLSelectTables) As dbo.SQLSelectTableJoins

End Function

Private Function IDatabaseObjects_TableName() As String
    
    IDatabaseObjects_TableName = "Suppliers"
    
End Function
