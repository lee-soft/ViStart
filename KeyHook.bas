Attribute VB_Name = "KeyBoard"
Option Explicit

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long

Private Type KEYBDINPUT
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Type HARDWAREINPUT
    uMsg As Long
    wParamL As Integer
    wParamH As Integer
End Type

Private Type GENERALINPUT
    dwType As Long
    xi(0 To 23) As Byte
End Type

Public Type KBDLLHOOKSTRUCT
    vkCode As Long        'value of the key you pressed
    scanCode As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type


Const VK_H = 72
Const VK_E = 69
Const VK_L = 76
Const VK_O = 79
Const KEYEVENTF_KEYUP = &H2
Const INPUT_KEYBOARD = 1

Public Const WH_KEYBOARD = 2
Public Const WH_KEYBOARD_LL = 13

Private Const LLKHF_EXTENDED = &H1
Private Const LLKHF_INJECTED = &H10
Private Const LLKHF_ALTDOWN = &H20
Private Const LLKHF_UP = &H80

Public Const VK_LWINKEY = &H5B      ' The Window's key on the left of the keyboard
Public Const VK_RWINKEY = &H5C      ' The Window's key on the right of the keyboard
                                    
Private m_hKeyboardHook As Long
Private m_keyCode As Long
Private m_allowWinKey As Boolean
Private m_getNext As Boolean
Private m_IgnoreOnce As Boolean
Private m_lastKey As Long

Private m_hWnd As Long
Private m_justWinKey As Boolean

Private m_otherKey As Boolean
Public g_ignoreHook As Boolean

Public Function UnhookKeyboard()
    'unhook the keyboard you will have some problems if this isnt called
    
    If m_hKeyboardHook <> 0 Then UnhookWindowsHookEx m_hKeyboardHook
End Function

Public Function HookKeyboard(hWnd As Long)
'Exit Function

    If m_hKeyboardHook <> 0 Then
        UnhookKeyboard
    End If

    m_hWnd = hWnd

    'hook the keyboard and recieve messages from the keyboard
    m_hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
End Function

Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, lParam As Long) As Long

Dim xpInfo As KBDLLHOOKSTRUCT
Dim eat As Boolean
    
    If g_ignoreHook Then
        LowLevelKeyboardProc = CallNextHookEx(m_hKeyboardHook, nCode, wParam, lParam) 'this will be called if there are multiple hooks made to the keyboard
        Exit Function
    End If
    
    CopyMemory xpInfo, lParam, Len(xpInfo) 'copy the structure from lParam to xpinfo
    
    'Debug.Print "xpInfo.vkCode:: " & xpInfo.vkCode
    
    If (Settings.CatchRightWindowsKey And xpInfo.vkCode = VK_RWINKEY) Or _
        (Settings.CatchLeftWindowsKey And xpInfo.vkCode = VK_LWINKEY) Then
    
        If xpInfo.Flags And LLKHF_UP Then
            
            If Not m_allowWinKey Then
                eat = True

                If g_bStartMenuVisible Then
                    'Make StartMenu be inactivate
                    PostMessage m_hWnd, WM_ACTIVATEAPP, ByVal MakeLong(0, WA_INACTIVE), 0
                Else
                    If Not m_otherKey Then
                        frmEvents.ActivateStartMenu
                    End If
                End If
                
                m_allowWinKey = True
                
                SetKeyDown vbKeyControl
                SetKeyUp xpInfo.vkCode
                SetKeyUp vbKeyControl
            End If
            
        Else
            m_otherKey = False
            m_allowWinKey = False
        End If
    Else
        If xpInfo.vkCode <> vbKeyLeft And xpInfo.vkCode <> vbKeyRight And xpInfo.vkCode <> vbKeyUp And xpInfo.vkCode <> vbKeyDown Then
            If g_bStartMenuVisible Then
            
                If xpInfo.vkCode = vbKeyEscape Then
                    PostMessage m_hWnd, WM_ACTIVATEAPP, ByVal MakeLong(0, WA_INACTIVE), 0
                End If
            
                
                If xpInfo.vkCode <> VK_RWINKEY And xpInfo.vkCode <> VK_LWINKEY And xpInfo.vkCode <> 162 Then
                    Debug.Print "ere:: " & xpInfo.vkCode
                    frmEvents.ActivateSearchText xpInfo.vkCode
                End If
            End If
        End If
    
        m_otherKey = True
    End If
    
    
    If eat Then
        LowLevelKeyboardProc = 1
    Else
        'this will be called if there are multiple hooks made to the keyboard
        LowLevelKeyboardProc = CallNextHookEx(m_hKeyboardHook, nCode, wParam, lParam) 'this will be called if there are multiple hooks made to the keyboard
    End If
    
    Exit Function
Handler:
    'this will be called if there are multiple hooks made to the keyboard
    LowLevelKeyboardProc = CallNextHookEx(m_hKeyboardHook, nCode, wParam, lParam) 'this will be called if there are multiple hooks made to the keyboard
End Function

Public Function SetKeyDown(KeyCode As Long)

Dim GInput(0 To 1) As GENERALINPUT
Dim KInput As KEYBDINPUT

    KInput.wVk = KeyCode 'the key we're going to press
    KInput.dwFlags = 0 'press the key
    'copy the structure into the input array's buffer.
    GInput(0).dwType = INPUT_KEYBOARD ' keyboard input
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    
    'send the input now
    Call SendInput(2, GInput(0), Len(GInput(0)))


End Function

Public Function SetKeyUp(KeyCode As Long)

Dim GInput(0 To 1) As GENERALINPUT
Dim KInput As KEYBDINPUT

    'do the same as above, but for releasing the key
    KInput.wVk = KeyCode ' the key we're going to realease
    KInput.dwFlags = KEYEVENTF_KEYUP ' release the key
    GInput(1).dwType = INPUT_KEYBOARD ' keyboard input
    CopyMemory GInput(1).xi(0), KInput, Len(KInput)
    'send the input now
    Call SendInput(2, GInput(0), Len(GInput(0)))

End Function



