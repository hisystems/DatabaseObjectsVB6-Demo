VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDatabaseObject

Private pobjParent As OrderDetails
Private pobjProduct As Product
Private pcurUnitPrice As Currency
Private pintQuantity As Integer
Private psngDiscount As Single

Friend Property Set Parent(ByVal objValue As OrderDetails)

    Set pobjParent = objValue

End Property

Public Property Let UnitPrice(ByVal curValue As Currency)
    
    'check that valid prices are set
    If curValue >= 0 Then
        pcurUnitPrice = curValue
    Else
        dbo.RaiseError dboErrorInvalidArgument, curValue
    End If

End Property

Public Property Get UnitPrice() As Currency
    
    UnitPrice = pcurUnitPrice
    
End Property

Public Property Let Quantity(ByVal intValue As Integer)

    'check that valid quantities are set
    If intValue >= 0 Then
        pintQuantity = intValue
    Else
        dbo.RaiseError dboErrorInvalidArgument, intValue
    End If

End Property

Public Property Get Quantity() As Integer

    Quantity = pintQuantity

End Property

Public Property Let Discount(ByVal sngValue As Single)

    If sngValue >= 0 Then
        psngDiscount = sngValue
    Else
        dbo.RaiseError dboErrorInvalidArgument, sngValue
    End If

End Property

Public Property Get Discount() As Single

    Discount = psngDiscount

End Property

Public Property Set Product(ByVal objValue As Product)

    If objValue Is Nothing Then
        dbo.RaiseError dboErrorObjectIsNothing
    End If

    Set pobjProduct = objValue

End Property

Public Property Get Product() As Product

    Set Product = pobjProduct

End Property

Public Property Get Cost() As Currency

    Cost = Me.UnitPrice * Me.Quantity

End Property

Private Property Let IDatabaseObject_DistinctValue(ByVal RHS As Variant)
    
    'This function is called when the object is loaded
    'but can be ignored and done in the IDatabaseObject_Load
    
End Property

Private Property Get IDatabaseObject_DistinctValue() As Variant

    IDatabaseObject_DistinctValue = pobjProduct.ID

End Property

Private Property Get IDatabaseObject_IsSaved() As Boolean
'This is always false because a new record is always written to the database for this
'object. This is because all of the OrderDetail objects are deleted before saving to the
'database. See the Order.Save function.
'

    IDatabaseObject_IsSaved = False

End Property

Private Property Let IDatabaseObject_IsSaved(ByVal RHS As Boolean)

End Property

Private Sub IDatabaseObject_Load(ByVal objFields As dbo.SQLFieldValues)

    Set pobjProduct = Data.Product(objFields("ProductID"))

    Me.UnitPrice = objFields("UnitPrice")
    Me.Quantity = objFields("Quantity")
    Me.Discount = objFields("Discount")

End Sub

Private Function IDatabaseObject_Save() As dbo.SQLFieldValues

    Dim objFields As SQLFieldValues
    Set objFields = New SQLFieldValues
    
    objFields.Add "OrderID", pobjParent.Parent.ID
    objFields.Add "UnitPrice", Me.UnitPrice
    objFields.Add "Quantity", Me.Quantity
    objFields.Add "Discount", Me.Discount
    
    Set IDatabaseObject_Save = objFields

End Function
