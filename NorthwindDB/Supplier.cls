VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDatabaseObject

Private plngID As Long
Private pstrName As String

Friend Property Get ID() As Long
'Friend will allow public access within the NorthwindDB project but will be
'private outside of the NorthwindDB project

    ID = plngID

End Property

Public Property Get Name() As String

    Name = pstrName

End Property

Public Property Let Name(ByVal strName As String)
    
    strName = Trim$(strName)
    
    'This check is not necessary as the DBO.ObjectSave checks that
    'the Key field is not null before it saves.
    If strName = vbNullString Then
        dbo.RaiseError dboErrorInvalidPropertyValue, "Null"
    Else
        pstrName = strName
    End If

End Property

Public Property Get IsDeletable() As Boolean

    'This checks the Products table to see if there are any references to this Supplier.
    'If there are then this supplier cannot be deleted.

    Dim objConnection As ADODB.Connection
    Dim objRecordset As ADODB.Recordset
    Dim objSelect As SQLSelect
    Set objSelect = New SQLSelect
    
    With objSelect
        .Fields.Add , , dboAggregateCount
        .Tables.Add "Products"
        .Where.Add "SupplierID", dboComparisonEqualTo, plngID
    End With
    
    Set objConnection = CreateConnection
    objConnection.Open
    
    Set objRecordset = objConnection.Execute(objSelect.SQL)
    IsDeletable = objRecordset.Fields(0) = 0

    objRecordset.Close
    objConnection.Close

End Property

Public Sub Save()

    dbo.ObjectSave Data.Suppliers, Me

End Sub

Private Property Let IDatabaseObject_DistinctValue(ByVal RHS As Variant)

    plngID = RHS
    
End Property

Private Property Get IDatabaseObject_DistinctValue() As Variant

    IDatabaseObject_DistinctValue = plngID

End Property

Private Property Get IDatabaseObject_IsSaved() As Boolean
'This function can be implemented using a boolean variable to hold
'whether this object has been saved or not. However, if the object is
'using an Identity field the following shortcut can be used as it is
'logically equivalent. The Property Let IDatabaseObject_IsSaved does
'not need to be implemented

    IDatabaseObject_IsSaved = plngID <> 0

End Property

Private Property Let IDatabaseObject_IsSaved(ByVal RHS As Boolean)

End Property

Private Sub IDatabaseObject_Load(ByVal objFields As dbo.SQLFieldValues)

    pstrName = objFields("CompanyName")

End Sub

Private Function IDatabaseObject_Save() As dbo.SQLFieldValues

    Dim objFields As SQLFieldValues
    Set objFields = New SQLFieldValues
    
    objFields.Add "CompanyName", pstrName
    
    Set IDatabaseObject_Save = objFields

End Function
