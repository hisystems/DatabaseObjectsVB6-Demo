VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDatabaseObjects

Private pobjParent As Order

'store each detail in a temporary collection
Private pcolDetails As Collection

Friend Property Set Parent(ByVal objValue As Order)

    Set pobjParent = objValue
    LoadDetails

End Property

Friend Property Get Parent() As Order
    
    Set Parent = pobjParent

End Property

Public Property Get Item(ByVal intIndex As Integer) As OrderDetail
Attribute Item.VB_UserMemId = 0

    Set Item = pcolDetails(intIndex)
    
End Property

Public Property Get Count() As Integer

    Count = pcolDetails.Count

End Property

Friend Sub Save()

    Dim intIndex As Integer
        
    'Delete all of the Detail lines associated with this Order
    dbo.ObjectsDeleteAll Me

    'Save all of the detail lines. Note that in the IDatabaseObject_IsSaved function
    'in the OrderDetail object is set to false. This ensures that records are always
    'inserted into the database not updated because all of the existing detail records
    'have just been deleted.
    For intIndex = 1 To Me.Count
        dbo.ObjectSave Me, Me.Item(intIndex)
    Next

End Sub

Private Sub LoadDetails()
'Loading all of the OrderDetail objects into a collection is necessary so that
'any adding, editing or deleting of the detail lines does not affect the database
'until the Order.Save sub is called.
'
    Set pcolDetails = dbo.ObjectsCollection(Me)

End Sub

Private Function IDatabaseObjects_DistinctFieldAutoIncrements() As Boolean

End Function

Private Function IDatabaseObjects_DistinctFieldName() As String
'The ProductID uniquely identifies each OrderDetail object within the
'Order - see the IDatabaseObjects_Subset function below. The 'ProductID' field
'together with the 'OrderID' field will determine which record each OrderDetail
'object is associated with.

    IDatabaseObjects_DistinctFieldName = "ProductID"
    
End Function

Private Function IDatabaseObjects_ItemInstance() As dbo.IDatabaseObject
'Because all instances of an OrderDetail object are created via this function
'we can utilise this and set the parent relationship.

    Dim objOrderDetail As OrderDetail
    
    Set objOrderDetail = New OrderDetail
    Set objOrderDetail.Parent = Me

    Set IDatabaseObjects_ItemInstance = objOrderDetail
    
End Function

Private Function IDatabaseObjects_KeyFieldName() As String

End Function

Private Function IDatabaseObjects_OrderBy() As dbo.SQLSelectOrderByFields

End Function

Private Function IDatabaseObjects_Subset() As dbo.SQLConditions

    Set IDatabaseObjects_Subset = New SQLConditions
    IDatabaseObjects_Subset.Add "OrderID", dboComparisonEqualTo, pobjParent.ID

End Function

Private Function IDatabaseObjects_TableJoins(ByVal objPrimaryTable As SQLSelectTable, ByVal objTables As dbo.SQLSelectTables) As dbo.SQLSelectTableJoins

End Function

Private Function IDatabaseObjects_TableName() As String

    IDatabaseObjects_TableName = "Order Details"

End Function
