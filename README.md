<div align="center">

## mAboutDialog


</div>

### Description

How would you like to have your own About menu option on the little system menu on the top left-hand corner of your form. I whould , so I wrote a module to do it with one line of code from the Load event on my form. For this code to work you have to create a About form first (FRMAbout).
 
### More Info
 
Call the code from your main form like so:

'Private Sub Form_Load()

' Call AddAboutForm(Me.hwnd, "About..")

'End Sub

It will check windows system messages for the click event on the system menu and then display your own FRMAbout.

DO NOT TRY AND STEP THIS CODE. Windows is doing calles to the function's in this module and could give you a GPF...


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Riaan Aspeling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/riaan-aspeling.md)
**Level**          |Unknown
**User Rating**    |4.7 (47 globes from 10 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/riaan-aspeling-maboutdialog__1-1629/archive/master.zip)

### API Declarations

```
'Paste this code into a module mAboutDialog
'
Option Explicit
'To variables and const we need
Public OldProcedure As Long
Public Const ABOUT_ID = 1010
Public Const WM_SYSCOMMAND = &H112
Public Const MF_SEPARATOR = &H800
Public Const MF_STRING = &H0
Public Const GWL_WNDPROC = &HFFFFFFFC
'The API's we need to do this
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
```


### Source Code

```
'Paste this code into a module mAboutDialog
'
'This is a subs function for windows system menu calls
Public Function SubsMenuProc(ByVal lFRMWinHandel As Long, ByVal lMessage As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 'Only capture system commands
 Select Case lMessage
  Case WM_SYSCOMMAND
   'Only capture our new about menu's clicks
   If wParam = ABOUT_ID Then
    'Show the about box
    FRMAbout.Show 1
    Exit Function
   End If
 End Select
 'Do the rest of windows stuff
 SubsMenuProc = CallWindowProc(OldProcedure, lFRMWinHandel, lMessage, wParam, lParam)
End Function
'This function should be called from the Onload event of the form you want
'the system menu to contain a About Menu
Public Sub AddAboutForm(ByVal lFormWindowHandel As Long, MenuDescription As String)
 Dim hSysMenu As Long
 'Get the handel to the system menu
 hSysMenu = GetSystemMenu(lFormWindowHandel, 0&)
 'Add a nice line
 Call AppendMenu(hSysMenu, MF_SEPARATOR, 0&, 0&)
 'Make sure you have a menu description
 If MenuDescription = "" Then MenuDescription = "About"
 'Add the About menu description
 Call AppendMenu(hSysMenu, MF_STRING, ABOUT_ID, MenuDescription)
 'Direct windows to the new function for the menu
 OldProcedure = SetWindowLong(lFormWindowHandel, GWL_WNDPROC, AddressOf SubsMenuProc)
End Sub
```

