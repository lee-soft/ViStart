VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileVersionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarProductBuildPart As Long 'local copy
Private mvarProductMajorPart As Long 'local copy
Private mvarProductMinorPart As Long 'local copy
Public Property Let ProductMinorPart(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductMinorPart = 5
    mvarProductMinorPart = vData
End Property


Public Property Get ProductMinorPart() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductMinorPart
    ProductMinorPart = mvarProductMinorPart
End Property



Public Property Let ProductMajorPart(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductMajorPart = 5
    mvarProductMajorPart = vData
End Property


Public Property Get ProductMajorPart() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductMajorPart
    ProductMajorPart = mvarProductMajorPart
End Property



Public Property Let ProductBuildPart(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductBuildPart = 5
    mvarProductBuildPart = vData
End Property


Public Property Get ProductBuildPart() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductBuildPart
    ProductBuildPart = mvarProductBuildPart
End Property




