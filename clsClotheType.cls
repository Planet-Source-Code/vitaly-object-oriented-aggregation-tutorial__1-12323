VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClothType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Color As String
Public Manufacturer As String

'\B/----------------------This is an example to using Property.----------------------
'These two Properties (And the private mvar) could be replaced with one:
'Public FirstTimeViewed as boolean
'Of course the initialize would be then: FirstTimeViewed=True
Private mvarFirstTimeViewed As Boolean
Public Property Let IsFirstTimeViewed(ByVal vData As Boolean)
    mvarFirstTimeViewed = vData
End Property
Public Property Get IsFirstTimeViewed() As Boolean
    IsFirstTimeViewed = mvarFirstTimeViewed
End Property
'/E\----------------------This is an example to using Property.----------------------

Private Sub Class_Initialize()
    'Booleans are loaded as False usually, this line runs when some variable becomes a
    'new instant of this class. It makes sure that all the instants of clsClothType will
    'be set to mvarFirstTimeViewed = True
    mvarFirstTimeViewed = True
End Sub
