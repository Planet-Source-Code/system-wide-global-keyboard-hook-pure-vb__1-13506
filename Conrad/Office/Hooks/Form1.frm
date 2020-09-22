VERSION 5.00
Begin VB.Form frmKeybd 
   Caption         =   "Keyboard Logger"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnhook 
      Caption         =   "Unhook"
      Height          =   495
      Left            =   2325
      TabIndex        =   1
      Top             =   855
      Width           =   1215
   End
   Begin VB.CommandButton cmdHook 
      Caption         =   "Hook"
      Height          =   495
      Left            =   855
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmKeybd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' form code

Private Sub cmdHook_Click()
 hook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf myfunc, App.hInstance, 0)
End Sub

Private Sub cmdUnhook_Click()
 UnhookWindowsHookEx hook
End Sub

Private Sub Form_Load()

End Sub
