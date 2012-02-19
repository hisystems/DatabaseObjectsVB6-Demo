VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDatabaseObject

Private plngID As Long
Private pobjDetails As OrderDetails

Public ShipName As String
Public ShipAddress As String
Public ShipCity As String
Public ShipRegion As String
Public ShipPostalCode As String
Public ShipCountry As String

Public TheDate As Date
Public RequiredDate As Date

Friend Property Get ID() As Long
'Friend will allow public access within the NorthwindDB project but will be
'private outside of the NorthwindDB project

    ID = plngID

End Property

Public Property Get Number() As Long

    Number = plngID

End Property

Public Property Get Cost() As Currency

    Dim intIndex As Integer
    Dim curCost As Currency
    
    For intIndex = 1 To Me.Details.Count
        curCost = curCost + Me.Details(intIndex).Cost
    Next

    Cost = curCost

End Property

Public Property Get Details() As OrderDetails

    If pobjDetails Is Nothing Then
        Set pobjDetails = New OrderDetails
        Set pobjDetails.Parent = Me
    End If

    Set Details = pobjDetails

End Property

Public Sub Save()
        
    'Start the transaction so that if any errors occur the data is not committed to the database
    dbo.Transactions.Begin
        
    'Save the main header object before saving the details, because if this is a new
    'Order object then the OrderID won't have been set yet and the OrderDetail objects
    'won't save with the correctly associated OrderID.
    dbo.ObjectSave Data.Orders, Me

    If Not pobjDetails Is Nothing Then
        'save all of the detail lines
        pobjDetails.Save
    End If
    
    'Commit the transaction to the database
    dbo.Transactions.Commit

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

    Me.TheDate = objFields("OrderDate")
    Me.RequiredDate = objFields("RequiredDate")

    Me.ShipName = objFields("ShipName")
    Me.ShipAddress = objFields("ShipAddress")
    Me.ShipCity = objFields("ShipCity")
    Me.ShipRegion = IIf(IsNull(objFields("ShipRegion")), vbNullString, objFields("ShipRegion"))
    Me.ShipPostalCode = IIf(IsNull(objFields("ShipPostalCode")), vbNullString, objFields("ShipPostalCode"))
    Me.ShipCountry = objFields("ShipCountry")

End Sub

Private Function IDatabaseObject_Save() As dbo.SQLFieldValues

    Dim objFields As SQLFieldValues
    Set objFields = New SQLFieldValues
    
    objFields.Add "OrderDate", Me.TheDate
    objFields.Add "RequiredDate", Me.RequiredDate
    
    objFields.Add "ShipName", Me.ShipName
    objFields.Add "ShipAddress", Me.ShipAddress
    objFields.Add "ShipCity", Me.ShipCity
    objFields.Add "ShipRegion", IIf(Me.ShipRegion = vbNullString, vbNull, Me.ShipRegion)
    objFields.Add "ShipPostalCode", IIf(Me.ShipPostalCode = vbNullString, vbNull, Me.ShipPostalCode)
    objFields.Add "ShipCountry", Me.ShipCountry

    Set IDatabaseObject_Save = objFields

End Function
