VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Orders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDatabaseObjects

Public Property Get Item(ByVal lngIndex As Long) As Order
Attribute Item.VB_UserMemId = 0

    Set Item = dbo.ObjectItem(Me, lngIndex)

End Property

Public Sub Delete(ByRef objOrder As Order)

    'The objOrder is passed ByRef because the DBO.ObjectDelete function
    'sets the objOrder = Nothing thereby ensuring it can no longer be accidentally used.
    
    dbo.ObjectsDeleteAll objOrder.Details
    dbo.ObjectDelete Me, objOrder

End Sub

Public Property Get Count() As Long

    Count = dbo.ObjectsCount(Me)

End Property

Private Function IDatabaseObjects_DistinctFieldAutoIncrements() As Boolean

    IDatabaseObjects_DistinctFieldAutoIncrements = True

End Function

Private Function IDatabaseObjects_DistinctFieldName() As String

    IDatabaseObjects_DistinctFieldName = "OrderID"

End Function

Private Function IDatabaseObjects_ItemInstance() As dbo.IDatabaseObject

    Set IDatabaseObjects_ItemInstance = New Order

End Function

Private Function IDatabaseObjects_KeyFieldName() As String

End Function

Private Function IDatabaseObjects_OrderBy() As dbo.SQLSelectOrderByFields

End Function

Private Function IDatabaseObjects_Subset() As dbo.SQLConditions

End Function

Private Function IDatabaseObjects_TableJoins(ByVal objPrimaryTable As SQLSelectTable, ByVal objTables As dbo.SQLSelectTables) As dbo.SQLSelectTableJoins

End Function

Private Function IDatabaseObjects_TableName() As String

    IDatabaseObjects_TableName = "Orders"
    
End Function
