<div align="center">

## System Wide \[ Global\] Keyboard Hook \- Pure VB


</div>

### Description

This Global Keyboard Hook / System wide keyboard hook is written entirely in VB. No Visual C++ code or any other tricks, just plain and simple API's. No more of the secretive VC ++ Dll's.

The sole purpose of this sample is only to demonstrate the ablility of Visual Basic to set a keyboard hook. No attention has been paid to the User interface or other enhancements. It has only the bare functional necessities.

It logs all keyboard events, doesnt filter out any of the system calls. The system calls this function every time a new keyboard input event is about to be posted into a thread input queue.
 
### More Info
 
Open the source code. go to the module and change the path of the text file to which the events are to be sent.



The sample logs the keyboard events to a 'c:\temp\keybd.txt' file. Please check if the file exists or change the file path in the code and compile.



Do not kill the program untill you have clicked the unhook button.

Tested on Win NT4 and Windows 2000. Yet to be tested on Win9x.


<span>             |<span>
---                |---
**Submitted On**   |2000-12-11 16:30:18
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Advanced
**User Rating**    |4.1 (33 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD1257712112000\.zip](https://github.com/Planet-Source-Code/system-wide-global-keyboard-hook-pure-vb__1-13506/archive/master.zip)

### API Declarations

```
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowsHook Lib "user32" Alias "SetWindowsHookA" (ByVal nFilterType As Long, ByVal pfnFilterProc As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Global Const WH_KEYBOARD_LL = 13
Public Const HC_ACTION = 0
Type HookStruct
  vkCode As Long
  scancode As Long
  flags As Long
  time As Long
  dwExtraInfo As Long
End Type
```





