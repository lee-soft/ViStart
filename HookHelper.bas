Attribute VB_Name = "HookHelper"
Option Explicit

' Subclassing Without The Crashes from vbAccelerator
' http://www.vbaccelerator.com/home/vb/Code/Libraries/Subclassing/SSubTimer/article.asp

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
  lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Declare Function GetProp Lib "user32" Alias "GetPropA" _
  (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" _
  (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" _
  (ByVal hWnd As Long, ByVal lpString As String) As Long
  
Private Declare Function SetWindowLongW Lib "user32" _
  (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long


Private Const HOOK_OBJECT_REFERENCE As String = "HOOK_OBJ"
Private Const OLD_WINDOW_PROC As String = "OLD_PROC"

Public Function CallOldWindowProcessor(ByVal sourcehWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim oldProcessorPointer As Long: oldProcessorPointer = GetProp(sourcehWnd, OLD_WINDOW_PROC)
   CallOldWindowProcessor = CallWindowProc(oldProcessorPointer, sourcehWnd, msg, wParam, lParam)

End Function

Public Sub UnhookWindow(ByVal sourcehWnd As Long)

Dim oldWindowProcedure As Long: oldWindowProcedure = GetProp(sourcehWnd, OLD_WINDOW_PROC)

    If (oldWindowProcedure <> 0) Then
       Call SetWindowLongW(sourcehWnd, GWL_WNDPROC, oldWindowProcedure)
    End If
End Sub

Public Function HookWindow(ByVal sourcehWnd As Long, hookObj As IHookSink) As Long
'Exit Function

    'set the property, 'HOOK_OBJ' to the pointer of the hookObj on the source hWnd
    Call SetProp(sourcehWnd, HOOK_OBJECT_REFERENCE, PtrFromObject(hookObj))
    
Dim oldWindowProcedure As Long: oldWindowProcedure = GetWindowLong(sourcehWnd, GWL_WNDPROC)

    'set the property, 'OLD_PROC' to the pointer of the old/vb6 window procedure
    Call SetProp(sourcehWnd, OLD_WINDOW_PROC, oldWindowProcedure)
    
    'switch the default vb6 window processor with the global callback
    HookWindow = SetWindowLongW(sourcehWnd, GWL_WNDPROC, AddressOf CallbackFunctionForAllWindows)
End Function

Private Function CallbackFunctionForAllWindows(ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long

Dim hookObject As IHookSink
Dim hookObjectPointer As Long: hookObjectPointer = GetProp(hWnd, HOOK_OBJECT_REFERENCE)
   
    If hookObjectPointer Then
        'Get the hookObj from the pointer
        Set hookObject = ObjectFromPtr(hookObjectPointer)
        
        'Call the new window processor and pass the result to the caller (windows)
        CallbackFunctionForAllWindows = hookObject.WindowProc(hWnd, msg, wp, lp)
    End If
End Function

Private Property Get PtrFromObject(ByRef oThis As IHookSink) As Long

  ' Return the pointer to this object:
  PtrFromObject = ObjPtr(oThis)

End Property

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As IHookSink
Dim oThis As IHookSink

  ' Turn the pointer into an illegal, uncounted interface
  CopyMemory oThis, lPtr, 4
  ' Do NOT hit the End button here! You will crash!
  ' Assign to legal reference
  Set ObjectFromPtr = oThis
  ' Still do NOT hit the End button here! You will still crash!
  ' Destroy the illegal reference
  CopyMemory oThis, 0&, 4
  ' OK, hit the End button if you must--you'll probably still crash,
  ' but this will be your code rather than the uncounted reference!

End Property
