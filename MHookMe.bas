Attribute VB_Name = "MHookMe"
' *************************************************************************
'  Copyright ©1997-2008 Karl E. Peterson and Zane Thomas,
'  All Rights Reserved, http://vb.mvps.org
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' May optionally be declared publicly so they may be called
' from the WindowProc's in client classes/forms/controls.
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Other Win32 APIs used only within this module.
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

' Used with GetWindowLong to retrieve the WindowProc address for hooked window.
Private Const GWL_WNDPROC As Long = -4&

' Property names used to stash info within window props.
' Optionally declare keyWndProc public?
Private Const keyObjPtr As String = "ObjectPointer"
Private Const keyWndProc As String = "OldWindowProc"

' **************************************************************
'  Public interface to setup and teardown subclassing.
' **************************************************************
Public Function HookWindow(ByVal hWnd As Long, thing As IHookSink) As Long
'Exit Function
   
   ' Stash pointer to object that will handle messages.
   Call SetProp(hWnd, keyObjPtr, ObjPtr(thing))
   ' Stash address of old window procedure.
   Call SetProp(hWnd, keyWndProc, GetWindowLong(hWnd, GWL_WNDPROC))
   ' Set new window procedure to point into this module.
   HookWindow = SetWindowLongW(hWnd, GWL_WNDPROC, AddressOf HookFunc)
End Function

Public Sub UnhookWindow(hWnd As Long)
   Dim lpWndProc As Long
   ' Retrieve stashed address of old window procedure.
   lpWndProc = GetProp(hWnd, keyWndProc)
   ' If valid, restore it to previous value.
   If (lpWndProc <> 0) Then
      Call SetWindowLongW(hWnd, GWL_WNDPROC, lpWndProc)
   End If
End Sub

' **************************************************************
'  A few public routines useful when handling messages.
' **************************************************************
Public Function InvokeWindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long
   ' This routine is provided for the handler to call whenever they want
   ' to pass message handling off to the default window procedure.
   InvokeWindowProc = CallWindowProc(GetProp(hWnd, "OldWindowProc"), hWnd, msg, wp, lp)
End Function

Public Function LOWORD(ByVal LongIn As Long) As Integer
   Call CopyMemory(LOWORD, LongIn, 2)
End Function

Public Function HiWord(ByVal LongIn As Long) As Integer
   Call CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Function

Public Function MakeLong(ByVal HiWord As Integer, ByVal LOWORD As Integer) As Long
   Call CopyMemory(MakeLong, LOWORD, 2)
   Call CopyMemory(ByVal (VarPtr(MakeLong) + 2), HiWord, 2)
End Function

' **************************************************************
'  Private functions called for each sunken message.
' **************************************************************
Private Function HookFunc(ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
   ' The object sinking messages for this window *must*
   ' implement the IHookSink interface!!!
   Dim Obj As IHookSink
   Dim lpObjPtr As Long
   
   ' Retreive pointer to handler object.
   lpObjPtr = GetProp(hWnd, keyObjPtr)
   If lpObjPtr Then
      ' Steal an object reference to handler.
      Set Obj = ObjPtrResolve(lpObjPtr)
      ' Call WindowProc, and return result to Windows.
      HookFunc = Obj.WindowProc(hWnd, msg, wp, lp)
   End If
End Function

Private Function ObjPtrResolve(ByVal lpThing As Long) As IHookSink
   Dim thing As IHookSink
   ' This function takes a dumb numeric pointer and turns it into
   ' a valid object reference.
   ' http://www.mvps.org/vbvision/collection_events.htm
   CopyMemory thing, lpThing, 4&
   Set ObjPtrResolve = thing
   CopyMemory thing, Nothing, 4&
End Function


