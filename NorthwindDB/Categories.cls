VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Categories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDatabaseObjects

Public Property Get Item(ByVal vIndex As Variant) As Category
Attribute Item.VB_UserMemId = 0

    Set Item = dbo.ObjectItem(Me, vIndex)

End Property

Public Property Get Count() As Integer

    Count = dbo.ObjectsCount(Me)

End Property

Private Function IDatabaseObjects_DistinctFieldAutoIncrements() As Boolean
    
    IDatabaseObjects_DistinctFieldAutoIncrements = True
    
End Function

Private Function IDatabaseObjects_DistinctFieldName() As String

    IDatabaseObjects_DistinctFieldName = "CategoryID"

End Function

Private Function IDatabaseObjects_ItemInstance() As dbo.IDatabaseObject

    Set IDatabaseObjects_ItemInstance = New Category

End Function

Private Function IDatabaseObjects_KeyFieldName() As String

    IDatabaseObjects_KeyFieldName = "CategoryName"

End Function

Private Function IDatabaseObjects_OrderBy() As dbo.SQLSelectOrderByFields

    Set IDatabaseObjects_OrderBy = New SQLSelectOrderByFields
    IDatabaseObjects_OrderBy.Add "CategoryName", dboOrderAscending

End Function

Private Function IDatabaseObjects_Subset() As dbo.SQLConditions

End Function

Private Function IDatabaseObjects_TableJoins(ByVal objPrimaryTable As SQLSelectTable, ByVal objTables As dbo.SQLSelectTables) As dbo.SQLSelectTableJoins

End Function

Private Function IDatabaseObjects_TableName() As String

    IDatabaseObjects_TableName = "Categories"

End Function
