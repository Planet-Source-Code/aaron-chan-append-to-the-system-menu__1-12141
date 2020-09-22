Attribute VB_Name = "modSysMenu"
'orginal example by VbWebExample
'enhanced by Aerodynamica Software

Public ProcOld As Long

'catch messages and call windows procedures
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'menu apis
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

'api constants
Public Const WM_SYSCOMMAND = &H112
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const GWL_WNDPROC = (-4)

'add new consts for new items
Public Const IDM_ITEM1 As Long = 0
Public Const IDM_ITEM2 As Long = 1
Public Const IDM_ABOUT As Long = 2

Public Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'do not debug this procedure, it will crash VB
    Select Case iMsg
        Case WM_SYSCOMMAND
            Select Case wParam
                Case IDM_ABOUT
                    MsgBox "Append to the System Menu. Based on example by VbWebExample, enhanced by Aerodynamica Software.", vbInformation, "About..."
                    Exit Function
                
                Case IDM_ITEM1
                    MsgBox "Item 1!", vbInformation, "Item 1"
                    Exit Function
                    
                Case IDM_ITEM2
                    MsgBox "Item 2!", vbInformation, "Item 2"
                    Exit Function
            End Select
    End Select
    
    'ass all messages on to VB and then return the value to windows
    WindowProc = CallWindowProc(ProcOld, hWnd, iMsg, wParam, lParam)
End Function
