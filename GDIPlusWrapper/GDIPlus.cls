VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPlusGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PixelFormat
   Alpha = 262144
   Canonical = 2097152
   DontCare = 0
   Extended = 1048576
   Format16bppArgb1555 = 397319
   Format16bppGrayScale = 1052676
   Format16bppRgb555 = 135173
   Format16bppRgb565 = 135174
   Format1bppIndexed = 196865
   Format24bppRgb = 137224
   Format32bppArgb = 2498570
   Format32bppPArgb = 925707
   Format32bppRgb = 139273
   Format48bppRgb = 1060876
   Format4bppIndexed = 197634
   Format64bppArgb = 3424269
   Format64bppPArgb = 1851406
   Format8bppIndexed = 198659
   Gdi = 131072
   Indexed = 65536
   Max = 15
   PAlpha = 524288
   Undefined = 0
End Enum

Public Function GDIPlusCreate(Optional suppressErrors As Boolean = False) As Boolean
   GDIPlusCreate = mGDIPlus.GDIPlusCreate(suppressErrors)
End Function

Public Sub GDIPlusDispose()
   mGDIPlus.GDIPlusDispose
End Sub
