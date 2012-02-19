VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LateBoundObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'This class can be used to simplify the process of creating and returning late-bound IDatabaseObjects objects,
'which is particularly useful in speeding up load times of objects by delaying loading associated objects
'until they are actually required ie. when the property is explicitly called, not when the object is loaded.
'This is achieved by passing the associated collection and the object's DistinctValue to the Initialize function
'The object is loaded only on the first call to the Object property. Subsequent calls to the Object property
'return the already loaded object.
'If lngDistinctValue is 0 (this class doesn't work with non-integer values) then the Object property will always return Nothing.
'If UseDefaultObject is true (false by default) and lngDistinctValue is 0 then the first item in the collection will be returned.

Option Explicit

Private plngDistinctValue As Long
Private pobjDatabase As dbo.Database
Private pobjCollection As IDatabaseObjects
Private pobjObject As IDatabaseObject
Private pbUseDefaultObject As Boolean

Public Sub Initialize( _
    ByVal objCollection As IDatabaseObjects, _
    ByVal lngDistinctValue As Variant, _
    Optional ByVal bUseDefaultObject As Boolean = False, _
    Optional ByVal objDatabase As dbo.Database = Nothing)

    If objCollection Is Nothing Then
        RaiseError dboErrorObjectIsNothing
    End If

    Set pobjDatabase = objDatabase
    Set pobjCollection = objCollection
    plngDistinctValue = lngDistinctValue
    pbUseDefaultObject = bUseDefaultObject

End Sub

Public Property Get Object() As IDatabaseObject
Attribute Object.VB_UserMemId = 0

    If pobjObject Is Nothing Then
        If plngDistinctValue = 0 Then
            If pbUseDefaultObject Then
                If pobjDatabase Is Nothing Then
                    Set pobjObject = dbo.ObjectByOrdinalFirst(pobjCollection)
                Else
                    Set pobjObject = pobjDatabase.ObjectByOrdinalFirst(pobjCollection)
                End If
                plngDistinctValue = pobjObject.DistinctValue
            End If
        Else
            If pobjDatabase Is Nothing Then
                Set pobjObject = dbo.Object(pobjCollection, plngDistinctValue)
            Else
                Set pobjObject = pobjDatabase.Object(pobjCollection, plngDistinctValue)
            End If
        End If
    End If
    
    Set Object = pobjObject

End Property

Public Property Set Object(ByVal objValue As IDatabaseObject)

    Set pobjObject = objValue

    If objValue Is Nothing Then
        plngDistinctValue = 0
    Else
        plngDistinctValue = objValue.DistinctValue
    End If

End Property

Public Property Get DistinctValue() As Long

    'Get the distinct value for the default object if it hasn't been loaded yet
    If pobjObject Is Nothing And pbUseDefaultObject Then
        plngDistinctValue = Me.Object.DistinctValue
    End If

    'This property can be used when saving the distinctvalue associated with this object
    'The Me.Object.DistinctValue could be used but this could cause an unnecessary loading
    'of the associated object when saving if the 'Property Get Object' hadn't been called
    'while the object was loaded

    DistinctValue = plngDistinctValue

End Property

Public Property Get IsNothing() As Boolean

    'This property can be used when saving the object
    'Me.Object Is Nothing could be used but this could cause an unnessary loading
    'of the associated object when saving if the 'Property Get Object' hadn't been called
    'while the object was loaded

    IsNothing = Me.DistinctValue = 0
    
End Property
