VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Below I have demonstrated 2 methods for implementing object relationships with the
'Category and Supplier properties. The Supplier property is the simplest method, however
'it can cause speed degradation because it must load the associated Supplier object everytime a Product is loaded.
'The speed degradation would be most noticable on a Product search screen or something similar.
'The second method demonstrated with the Category object will only load the object when it the property
'is explicity called - i.e. it will only store the CategoryID and load the associated Category object
'when the property is explicity called. This method results in faster loading speeds and still provides
'the necessary object relationships. Because this is a fairly common situation I have created the
'LateBoundObject class which handles this neatly, and can be used as an alternative to the method
'demonstrated below using the plngCategoryID and pobjCategory variables. This method is demonstrated
'this using the Category2 property.
'

Option Explicit

Implements IDatabaseObject

Private pbSaved As Boolean
Private plngID As Long
Private pstrName As String
Private pcurUnitPrice As Currency
Private pintUnitsInStock As Integer
Private pintUnitsOnOrder As Integer

Private plngCategoryID As Long
Private pobjCategory As Category

'Private pobjCategory2 As LateBoundObject

Private pobjSupplier As Supplier

Public QuantityPerUnit As String
Public Discontinued As Boolean

Friend Property Get ID() As Long
'Friend will allow public access within the NorthwindDB project but will be
'private outside of the NorthwindDB project

    ID = plngID

End Property

Public Sub AddUnits(ByVal intQuantity As Integer)

    If intQuantity >= 0 Then
        pintUnitsInStock = pintUnitsInStock + intQuantity
    Else
        dbo.RaiseError dboErrorInvalidArgument, "Quantity: " & intQuantity
    End If

End Sub

Public Sub RemoveUnits(ByVal intQuantity As Integer)

    If intQuantity >= 0 Then
        If intQuantity > pintUnitsInStock Then
            dbo.RaiseError dboErrorInvalidArgument, "Cannot remove more units than are available."
        End If
    
        pintUnitsInStock = pintUnitsInStock - intQuantity
    Else
        dbo.RaiseError dboErrorInvalidArgument, "Quantity: " & intQuantity
    End If

End Sub

Public Property Get UnitsInStock() As Integer

    UnitsInStock = pintUnitsInStock

End Property

Public Property Get IsInStock() As Boolean

    IsInStock = pintUnitsInStock > 0

End Property

Public Property Get IsOnOrder() As Boolean
    
    IsOnOrder = pintUnitsOnOrder > 0

End Property

Public Property Let Name(ByVal strValue As String)
    
    If Trim$(strValue) = vbNullString Then
        dbo.RaiseError dboErrorInvalidPropertyValue, "Null"
    Else
        pstrName = strValue
    End If

End Property

Public Property Get Name() As String

    Name = pstrName

End Property

Public Property Let UnitPrice(ByVal curValue As Currency)

    If curValue >= 0 Then
        pcurUnitPrice = curValue
    Else
        dbo.RaiseError dboErrorInvalidPropertyValue, curValue
    End If

End Property

Public Property Get UnitPrice() As Currency

    UnitPrice = pcurUnitPrice

End Property

Public Property Get Category() As Category
    
    If pobjCategory Is Nothing Then
        If plngCategoryID > 0 Then
            Set pobjCategory = Data.Category(plngCategoryID)
        Else
            Set pobjCategory = Data.Categories(1)
            plngCategoryID = pobjCategory.ID
        End If
    End If
    
    Set Category = pobjCategory

End Property

Public Property Set Category(ByVal objValue As Category)

    If objValue Is Nothing Then
        dbo.RaiseError dboErrorObjectIsNothing
    End If

    Set pobjCategory = objValue
    
    'Alternatively the DBO.ObjectDistinctValue function can be used to return the
    'same ID value. In this instance I have create a Friend property in the Category class
    'to make the code a little more readable.
    plngCategoryID = pobjCategory.ID

End Property

'Public Property Get Category2() As Category
'
'    'This property uses the LateBoundObject class to load the associated
'    'category object - on the first call to this property
'
'    Set Category2 = pobjCategory2.Object
'
'End Property
'
'Public Property Set Category2(ByVal objValue As Category)
'
'    Set pobjCategory2.Object = objValue
'
'End Property

Public Property Get Supplier() As Supplier
    
    Set Supplier = pobjSupplier

End Property

Public Sub Save()

    dbo.ObjectSave Data.gobjNorthwindInstance.Products, Me

End Sub

Private Sub Class_Initialize()
    
    'For a new product instance default to using the default category
    'Set pobjCategory2 = New LateBoundObject
    'pobjCategory2.Initialize Data.Categories, 0, bUseDefaultObject:=True

    'Alternatively, for a new product instance default the category to Nothing
    'Set pobjCategory2 = New LateBoundObject

End Sub

Private Property Get IDatabaseObject_DistinctValue() As Variant

    'This value MUST matchup to the value stored for the DistinctFieldName in the associated product's
    'IDatabaseObjects_DistinctFieldName function
    
    IDatabaseObject_DistinctValue = plngID
    
End Property

Private Property Let IDatabaseObject_DistinctValue(ByVal RHS As Variant)

    plngID = RHS

End Property

Private Property Get IDatabaseObject_IsSaved() As Boolean
'If the Product object uses an identity field then a shortcut for using this function is:
'IDatabaseObject_IsSaved = plngID <> 0  - as this signifies whether the object has been saved
'and is logically equivalent. The 'Property Let IDatabaseObject_IsSaved' would be left blank in that case.
'For the sake of simplicity I've used a variable pbSaved - which will function exactly the same. Of course,
'there are circumstances when the IDatabaseObject_IsSaved will NEED to use a variable as it doesn't use
'an Identity field, hence the need for the IsSaved function.
'

    IDatabaseObject_IsSaved = pbSaved
    
End Property

Private Property Let IDatabaseObject_IsSaved(ByVal RHS As Boolean)
    
    pbSaved = RHS
    
End Property

Private Sub IDatabaseObject_Load(ByVal objFields As dbo.SQLFieldValues)
'This function should copy all of the relevant field values from the database
'into the object.

    pstrName = objFields("ProductName")
    pcurUnitPrice = objFields("UnitPrice")
    pintUnitsOnOrder = objFields("UnitsOnOrder")
    pintUnitsInStock = objFields("UnitsInStock")
    Me.QuantityPerUnit = objFields("QuantityPerUnit")
    Me.Discontinued = objFields("Discontinued")
    plngCategoryID = objFields("CategoryID")
    
    'Set pobjCategory2 = New LateBoundObject
    'pobjCategory2.Initialize Data.Categories, objFields("CategoryID")
    
    If objFields.Exists("CompanyName") Then
        'load the object from the objFields because it contains all of the associated supplier fields
        Set pobjSupplier = dbo.ObjectFromFieldValues(Data.Suppliers, objFields)
    Else
        'load the supplier with an SQL call
        Set pobjSupplier = Data.Supplier(objFields("SupplierID"))
    End If

End Sub

Private Function IDatabaseObject_Save() As dbo.SQLFieldValues
'This function should copy all of the relevant field values from the object
'into the database.

    Dim objFields As SQLFieldValues
    Set objFields = New SQLFieldValues

    objFields.Add "ProductName", pstrName
    objFields.Add "UnitPrice", pcurUnitPrice
    objFields.Add "UnitsOnOrder", pintUnitsOnOrder
    objFields.Add "UnitsInStock", pintUnitsInStock
    objFields.Add "QuantityPerUnit", Me.QuantityPerUnit
    objFields.Add "Discontinued", Me.Discontinued
    objFields.Add "SupplierID", dbo.ObjectDistinctValue(pobjSupplier)
    objFields.Add "CategoryID", plngCategoryID
    'objFields.Add "CategoryID", pobjCategory2.DistinctValue

    Set IDatabaseObject_Save = objFields

End Function
