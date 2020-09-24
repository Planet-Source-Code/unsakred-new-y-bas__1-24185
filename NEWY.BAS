Attribute VB_Name = "NewY"
'New Y Version 1.0
'This bas is For use with the Yahoo Instant Messenger Client
'Codes written By UnSaKreD (http://digitalrevolt.cjb.net)
'If my bas is used in a project please give me credit, thank you.
'Greets: Andy Maul,l0hr,plasma,Lithium,HaDeS,Rage,Top,Izeakel,doodster,invisible,jjmoneydawg,urdeadagain
'Special thanks to "Masta" he was the reason i started programming, big thanks to you man.
'I miss you come back damnit!
'Lots o love ;)


Option Explicit
      #If Win32 Then
        Private Declare Function sndPlaySound Lib "Winmm" Alias _
           "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) _
           As Long
      #Else
        Private Declare Function sndPlaySound Lib "MMSYSTEM" ( _
                           lpszSoundName As Any, ByVal uFlags%) As Integer
      #End If
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetCursorPos Lib "user32" () As Single
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function iswindowenabled Lib "user32" Alias "IsWindowEnabled" (ByVal hwnd As Long) As Long
Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long


Const EM_UNDO = &HC7
Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2
Global Const WM_MDICREATE = &H220
Global Const WM_MDIDESTROY = &H221
Global Const WM_MDIACTIVATE = &H222
Global Const WM_MDIRESTORE = &H223
Global Const WM_MDINEXT = &H224
Global Const WM_MDIMAXIMIZE = &H225
Global Const WM_MDITILE = &H226
Global Const WM_MDICASCADE = &H227
Global Const WM_MDIICONARRANGE = &H228
Global Const WM_MDIGETACTIVE = &H229
Global Const WM_MDISETMENU = &H230
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
Public Const SND_MEMORY = &H4

Public Const WM_SETFOCUS = &H7
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181
Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_UPDATEINIFILE = 1

Public Const SW_ERASE = &H4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_ShowMinimized = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_SEPARATOR = &H800&

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const ENTA = 13
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const EM_LINESCROLL = &HB6
Private Const SPI_SCREENSAVERRUNNING = 97
Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Type POINTAPI
x As Long
Y As Long
End Type
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub Pause(interval)
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
Public Sub WindowClose(lngHwnd As Long)
    Call SendMessageLong(lngHwnd&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub WindowHide(lngHwnd As Long)
    Call ShowWindow(lngHwnd&, SW_HIDE)
End Sub
Public Sub WindowShow(lngHwnd As Long)
    Call ShowWindow(lngHwnd&, SW_SHOW)
End Sub
Sub ClickMenu(lngwindow As Long, strmenutext As String)
'This is from Andymaul one of my closest friends
'Thanks man.
Dim intLoop As Integer, intSubLoop As Integer, intSub2Loop As Integer, intSub3Loop As Integer, intSub4Loop As Integer
Dim lngmenu(1 To 5) As Long
Dim lngcount(1 To 5) As Long
Dim lngSubMenuID(1 To 4) As Long
Dim strcaption(1 To 4) As String

    lngmenu(1) = GetMenu(lngwindow&)

    lngcount(1) = GetMenuItemCount(lngmenu(1))

        For intLoop% = 0 To lngcount(1) - 1

            DoEvents
    
            lngmenu(2) = GetSubMenu(lngmenu(1), intLoop%)
    
            lngcount(2) = GetMenuItemCount(lngmenu(2))
    
                For intSubLoop% = 0 To lngcount(2) - 1
    
                    DoEvents
        
                    lngSubMenuID(1) = GetMenuItemID(lngmenu(2), intSubLoop%)
        
                    strcaption(1) = String(75, " ")
        
                    Call GetMenuString(lngmenu(2), lngSubMenuID(1), strcaption(1), 75, 1)
        
                        If InStr(LCase(strcaption(1)), LCase(strmenutext$)) Then
        
                            Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(1), 0)
        
                            Exit Sub
        
                        End If
                    
                    lngmenu(3) = GetSubMenu(lngmenu(2), intSubLoop%)
                    
                    lngcount(3) = GetMenuItemCount(lngmenu(3))
                    
                        If lngcount(3) > 0 Then
                    
                            For intSub2Loop% = 0 To lngcount(3) - 1
                        
                                DoEvents
                            
                                lngSubMenuID(2) = GetMenuItemID(lngmenu(3), intSub2Loop%)
                            
                                strcaption(2) = String(75, " ")
                                
                                Call GetMenuString(lngmenu(3), lngSubMenuID(2), strcaption(2), 75, 1)
                                
                                    If InStr(LCase(strcaption(2)), LCase(strmenutext$)) Then
                                    
                                        Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(2), 0)
                                        
                                        Exit Sub
                                        
                                    End If
                                    
                                lngmenu(4) = GetSubMenu(lngmenu(3), intSub2Loop%)
                                
                                lngcount(4) = GetMenuItemCount(lngmenu(4))
                                
                                    If lngcount(4) > 0 Then
                                    
                                        For intSub3Loop% = 0 To lngcount(4) - 1
                                        
                                            DoEvents
                                            
                                            lngSubMenuID(3) = GetMenuItemID(lngmenu(4), intSub3Loop%)
                            
                                            strcaption(3) = String(75, " ")
                                
                                            Call GetMenuString(lngmenu(4), lngSubMenuID(3), strcaption(3), 75, 1)
                                
                                                If InStr(LCase(strcaption(3)), LCase(strmenutext$)) Then
                                    
                                                    Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(3), 0)
                                        
                                                    Exit Sub
                                        
                                                End If
                                            
                                            lngmenu(5) = GetSubMenu(lngmenu(4), intSub3Loop%)
                                            
                                            lngcount(5) = GetMenuItemCount(lngmenu(5))
                                            
                                                If lngcount(5) > 0 Then
                                                
                                                    For intSub4Loop% = 0 To lngcount(5) - 1
                                                    
                                                        DoEvents
                                                        
                                                        lngSubMenuID(4) = GetMenuItemID(lngmenu(5), intSub4Loop%)
                                                        
                                                        strcaption(4) = String(75, " ")
                                                        
                                                        Call GetMenuString(lngmenu(5), lngSubMenuID(4), strcaption(4), 75, 1)
                                                        
                                                            If InStr(LCase(strcaption(4)), LCase(strmenutext$)) Then
                                                            
                                                                Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(4), 0)
                                                                
                                                                Exit Sub
                                                                
                                                            End If
                                                            
                                                    Next intSub4Loop%
                                
                                                End If
                                
                                        Next intSub3Loop%
                                        
                                    End If
                        
                            Next intSub2Loop%
    
                        End If
    
                Next intSubLoop%

        Next intLoop%

End Sub

Private Function GetCaption(hwnd)
Dim hwndlength As Integer, hwndtitle As String, a As Integer
hwndlength% = GetWindowTextLength(hwnd)
hwndtitle$ = String$(hwndlength%, 0)
a% = GetWindowText(hwnd, hwndtitle$, (hwndlength% + 1))

GetCaption = hwndtitle$
End Function
Function FindPMWnd()
Dim imclass As Long
imclass& = FindWindow("IMClass", vbNullString)
If InStr(GetCaption(imclass&), LCase(" -- instant message")) Then: FindPMWnd = imclass&
End Function
Function FindMainWnd()
FindMainWnd = FindWindow("Yahoobuddymain", vbNullString)
End Function
Function FindChatWnd()
Dim imclass As Long
imclass& = FindWindow("IMClass", vbNullString)
If InStr(GetCaption(imclass&), LCase(" -- chat")) Then: FindChatWnd = imclass&
End Function
Public Function ClickButton(button As Long)
Dim click As Long
click& = sendmessagebynum(button, WM_LBUTTONDOWN, &HD, 0)
click& = sendmessagebynum(button, WM_LBUTTONUP, &HD, 0)
End Function
Sub WindowDisable(Window As Long)
Call EnableWindow(Window&, 0)
End Sub
Sub WindowEnable(Window As Long)
Call EnableWindow(Window&, 1)
End Sub
Public Sub SendText(what$)
Dim imc As Long, richedit As Long
imc& = FindWindow("IMClass", vbNullString)
richedit& = FindWindowEx(imc&, 0&, "RICHEDIT", vbNullString)
Call SendMessageByString(richedit, WM_SETTEXT, 0&, what$)
Call Pause(0.2)
Call ClickMenu(imc&, "Sen&d")
End Sub
Sub SendBoot(code$)
Dim imc As Long, richedit As Long
imc& = FindWindow("IMClass", vbNullString)
richedit& = FindWindowEx(imc&, 0&, "RICHEDIT", vbNullString)
richedit& = FindWindowEx(imc&, richedit&, "RICHEDIT", vbNullString)
Call PostMessage(richedit&, WM_CLOSE, 0&, 0&)
richedit& = FindWindowEx(imc&, 0&, "RICHEDIT", vbNullString)
Call SendMessageByString(richedit, WM_SETTEXT, 0&, code$)
Call Pause(0.2)
Call ClickMenu(imc&, "Sen&d")
End Sub
Sub SendTextScroll(what As String, times As Integer)
Do
SendText (what$)
times = times% - 1
Call Pause(0.3)
Loop Until times% = 0
End Sub

Sub SendPM(who$, what$, Follow As Boolean)
Dim yahoobuddymain As Long, Parent As Long, Child1 As Long, Child2 As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
Call ClickMenu(yahoobuddymain&, "Send a &Message")
Call Pause(0.2)
Parent& = FindWindow("IMClass", vbNullString)
If InStr(GetCaption(Parent&), "Chat") Then Exit Sub
Child1& = FindWindowEx(Parent&, 0&, "Edit", vbNullString)
Call SetFocusAPI(Child1&)
Call SendMessageByString(Child1&, WM_SETTEXT, 0&, who$)
Call sendmessagebynum(Child1&, WM_CHAR, 13, 0&)
Child2& = FindWindowEx(Parent&, 0&, "RICHEDIT", vbNullString)
Call SendMessageByString(Child2&, WM_SETTEXT, 0&, what$)
Call ClickMenu(Parent&, "Sen&d")
If Follow = False Then: Call WindowClose(Parent&)
End Sub
Sub SendChat(what$)
Dim Parent As Long, Child1 As Long, Child2 As Long
Parent& = FindWindow("IMClass", vbNullString)
If InStr(GetCaption(Parent&), LCase("-- instant message")) Then: Exit Sub
Call SetFocusAPI(Parent&)
Child2& = FindWindowEx(Parent&, 0&, "RICHEDIT", vbNullString)
Call SendMessageByString(Child2&, WM_SETTEXT, 0, what$)
Call ClickMenu(Parent&, "Sen&d")
End Sub
Sub SendChatBoot(code$, anti As Boolean, StayIn As Boolean)
Dim imc As Long, Rich As Long, button As Long
Dim richedit As Long
imc& = FindWindow("IMClass", vbNullString)
If InStr(GetCaption(imc&), "Chat") Then: SetFocusAPI (imc&)
imc& = FindWindow("IMClass", vbNullString)
richedit& = FindWindowEx(imc&, 0&, "RICHEDIT", vbNullString)
richedit& = FindWindowEx(imc&, richedit&, "RICHEDIT", vbNullString)
If anti = True Then: Call PostMessage(richedit&, WM_CLOSE, 0&, 0&)
Rich& = FindWindowEx(imc&, 0&, "RICHEDIT", vbNullString)
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, code$)
Call Pause(0.2)
Call ClickMenu(imc&, "Sen&d")
Call Pause(0.2)
If StayIn = False Then: WindowClose (imc&)
End Sub

Sub SendPMBoot(UserName$, code$, anti As Boolean, Follow As Boolean)
Dim Parent As Long, Child1 As Long, Child2 As Long, button As Long
Dim Yahoo As Long
Dim imclass As Long, richedit As Long
Yahoo& = FindWindow("YahooBuddyMain", vbNullString)
Call ClickMenu(Yahoo&, "Send a &Message")
imclass& = FindWindow("IMClass", vbNullString)
If InStr(GetCaption(imclass&), "Chat") Then Exit Sub
richedit& = FindWindowEx(imclass&, 0&, "RICHEDIT", vbNullString)
richedit& = FindWindowEx(imclass&, richedit&, "RICHEDIT", vbNullString)

If anti = True Then: Call PostMessage(richedit&, WM_CLOSE, 0&, 0&)

Parent& = FindWindow("IMClass", vbNullString)
richedit& = FindWindowEx(Parent&, 0&, "RICHEDIT", vbNullString)
richedit& = FindWindowEx(Parent&, richedit&, "RICHEDIT", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "Edit", vbNullString)

Call SetFocusAPI(Child1&)

Call SendMessageByString(Child1&, WM_SETTEXT, 0&, UserName$)

Call sendmessagebynum(Child1&, WM_CHAR, 13, 0&)

Child2& = FindWindowEx(Parent&, 0&, "RICHEDIT", vbNullString)


Call SetFocusAPI(Child2&)

Call SendMessageByString(Child2&, WM_SETTEXT, 0&, code$)

Call ClickMenu(Parent&, "Sen&d")

Call Pause(0.3)

If Follow = False Then: WindowClose (Parent&)
End Sub
Sub SendPMLagg(UserName$, code$, anti As Boolean, times As Integer)
Dim Parent As Long, Child1 As Long, Child2 As Long, button As Long
Dim Yahoo As Long
Dim imclass As Long, richedit As Long
Yahoo& = FindWindow("YahooBuddyMain", vbNullString)
Call ClickMenu(Yahoo&, "Send a &Message")
imclass& = FindWindow("IMClass", vbNullString)
If InStr(GetCaption(imclass&), "Chat") Then Exit Sub
richedit& = FindWindowEx(imclass&, 0&, "RICHEDIT", vbNullString)
richedit& = FindWindowEx(imclass&, richedit&, "RICHEDIT", vbNullString)

If anti = True Then: Call PostMessage(richedit&, WM_CLOSE, 0&, 0&)

Parent& = FindWindow("IMClass", vbNullString)
richedit& = FindWindowEx(Parent&, 0&, "RICHEDIT", vbNullString)
richedit& = FindWindowEx(Parent&, richedit&, "RICHEDIT", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "Edit", vbNullString)

Call SetFocusAPI(Child1&)

Call SendMessageByString(Child1&, WM_SETTEXT, 0&, UserName$)

Call sendmessagebynum(Child1&, WM_CHAR, 13, 0&)

Child2& = FindWindowEx(Parent&, 0&, "RICHEDIT", vbNullString)


Call SetFocusAPI(Child2&)

Do
Call SendMessageByString(Child2&, WM_SETTEXT, 0&, code$)

Call ClickMenu(Parent&, "Sen&d")

Call Pause(0.3)
times% = times% - 1
Loop Until times% = 0
Call WindowClose(Parent&)
End Sub
Public Function GetYahooText()
Dim imc As Long, Rich As Long
Dim texts As String, thetextlen As Long

imc& = FindWindow("imclass", vbNullString)
Rich& = FindWindowEx(imc&, 0&, "richedit", vbNullString)
Rich& = FindWindowEx(imc&, Rich, "richedit", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(Rich&, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(Rich&, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
If TheText = "" Then GoTo NoText
        thetextlen& = (Len(TheText) - 2)
        TheText$ = Left$(TheText, thetextlen&)
GetYahooText = TheText
NoText:
End Function
Sub SendPMScroll(who$, what$, times As Integer, Follow As Boolean)
Dim yahoobuddymain As Long, Parent As Long, Child1 As Long, Child2 As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
Call ClickMenu(yahoobuddymain&, "Send a &Message")
Call Pause(0.2)
Parent& = FindWindow("IMClass", vbNullString)
If InStr(GetCaption(Parent&), LCase("chat")) Then: Exit Sub
Call SetFocusAPI(Parent&)
Child1& = FindWindowEx(Parent&, 0&, "Edit", vbNullString)
Call SetFocusAPI(Child1&)
Call SendMessageByString(Child1&, WM_SETTEXT, 0, who$)
Child2& = FindWindowEx(Parent&, 0&, "RICHEDIT", vbNullString)
Do
Call SendMessageByString(Child2&, WM_SETTEXT, 0, what$)
times% = times% - 1
Call ClickMenu(Parent&, "Sen&d")
Call Pause(0.3)
Loop Until times% = 0
If Follow = False Then: Call WindowClose(Parent&)
End Sub
Sub SendFile(who$, File$, Message$)
Dim Yahoo As Long, imclass As Long, richedit As Long, editx As Long, button As Long
Yahoo = FindWindow("yahoobuddymain", vbNullString)
Call ClickMenu(Yahoo&, "Send a &File...")
Call Pause(0.2)
imclass = FindWindow("imclass", "Send a File...")
richedit = FindWindowEx(imclass, 0&, "richedit", vbNullString)
Call SetFocusAPI(richedit&)
Call SendMessageByString(richedit&, WM_SETTEXT, 0, who$)
Call Pause(0.2)
editx = FindWindowEx(imclass, 0&, "edit", vbNullString)
Call SetFocusAPI(editx&)
Call SendMessageByString(editx&, WM_SETTEXT, 0, File$)
Call Pause(0.2)
editx = FindWindowEx(imclass, editx, "edit", vbNullString)
Call SetFocusAPI(editx&)
Call SendMessageByString(editx, WM_SETTEXT, 0&, Message$)
button = FindWindowEx(imclass, 0&, "button", vbNullString)
button = FindWindowEx(imclass, button, "button", vbNullString)
Call SetFocusAPI(button&)
Call ClickButton(button&)
End Sub
Sub YahooClose()
Dim yahoobuddymain As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
Call ClickMenu(yahoobuddymain&, "C&lose")
End Sub
Sub ChatClear()
Dim imclass As Long, richedit As Long
imclass& = FindWindow("IMClass", vbNullString)
richedit& = FindWindowEx(imclass&, 0&, "RICHEDIT", vbNullString)
richedit& = FindWindowEx(imclass&, richedit&, "RICHEDIT", vbNullString)
If InStr(LCase(GetCaption(imclass&)), "chat") Then
Call SendMessageByString(richedit&, WM_SETTEXT, 0&, "")
Else: End If
End Sub
Sub ChatHide()
Call WindowHide(FindChatWnd)
End Sub
Sub ChatShow()
Call WindowShow(FindChatWnd)
End Sub
Sub ChatClearTheirs()
Dim Text As String
Text$ = vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
Call SendChat(Text$)
End Sub

Sub PMClear()
'Clears PM
Dim imclass As Long, richedit As Long
imclass& = FindWindow("IMClass", vbNullString)
richedit& = FindWindowEx(imclass&, 0&, "RICHEDIT", vbNullString)
richedit& = FindWindowEx(imclass&, richedit&, "RICHEDIT", vbNullString)
If InStr(LCase(GetCaption(imclass&)), "instant message") Then
Call SendMessageByString(richedit&, WM_SETTEXT, 0&, "")
Else: End If
End Sub
Sub PMHide()
Call WindowHide(FindPMWnd)
End Sub
Sub PMShow()
Call WindowShow(FindPMWnd)
End Sub

Function PMFrom()
'Get's the open pm user
Dim imclass As Long, Str As String
imclass& = FindWindow("IMClass", vbNullString)
If InStr(LCase(GetCaption(imclass&)), "instant message") Then
Str$ = GetCaption(imclass&)
Str$ = Replace(Str$, " -- Instant Message", "")
PMFrom = Str$
Else: PMFrom = "": End If
End Function
Function PMIgnore()
'Ignores current user.
Dim imclass As Long, Str As String
imclass& = FindWindow("IMClass", vbNullString)
If InStr(LCase(GetCaption(imclass&)), "instant message") Then
Call ClickMenu(imclass&, "&Ignore User...")
Else: End If
End Function
Function PMVoiceOnOff()
Call ClickMenu(FindPMWnd, "Enable &Voice")
End Function
Function Ignore(User As String)
'Ignores specific user
Dim yahoobuddymain As Long, Parent As Long, Child1 As Long, Child2 As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
Call ClickMenu(yahoobuddymain&, "Send a &Message")
Call Pause(0.2)
Parent& = FindWindow("IMClass", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "Edit", vbNullString)
Call SetFocusAPI(Child1&)
Call SendMessageByString(Child1&, WM_SETTEXT, 0&, User$)
Call sendmessagebynum(Child1&, WM_CHAR, 13, 0&)
Call ClickMenu(Parent&, "&Ignore User...")

End Function

Sub PMClose()
'Closes PM
Dim imclass As Long, richedit As Long
imclass& = FindWindow("IMClass", vbNullString)
richedit& = FindWindowEx(imclass&, 0&, "RICHEDIT", vbNullString)
richedit& = FindWindowEx(imclass&, richedit&, "RICHEDIT", vbNullString)
If InStr(LCase(GetCaption(imclass&)), "instant message") Then
Call WindowClose(imclass&)
Else: End If
End Sub
Sub ChatClose()
'Closes Chat
Dim imclass As Long, richedit As Long
imclass& = FindWindow("IMClass", vbNullString)
richedit& = FindWindowEx(imclass&, 0&, "RICHEDIT", vbNullString)
richedit& = FindWindowEx(imclass&, richedit&, "RICHEDIT", vbNullString)
If InStr(LCase(GetCaption(imclass&)), "chat") Then
Call WindowClose(imclass&)
Else: End If
End Sub
Function ChatVoiceOnOff()
Call ClickMenu(FindChatWnd, "Enable &Voice")
End Function
Sub SendChatScroll(what As String, times As Integer)
'Sends a Chat Scroll
Do
Call SendChat(what$)
Call Pause(0.3)
times% = times% - 1
Loop Until times% = 0
End Sub
Sub SendChatLagg(code$, times%)
'Send's a Chat Lagg
Do
Call SendChat(code$)
Call Pause(0.3)
Call ChatClear
times% = times% - 1
Loop Until times% = 0
End Sub
Sub SignIn(UserName$, Password$, SaveID As Boolean, AutoLogin As Boolean, Invisible As Boolean)
Dim x As Long, editx As Long, button As Long
Dim yahoobuddymain As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
Call ClickMenu(yahoobuddymain&, "C&lose")
x = FindWindow("#32770", "Login")
If x& = True Then GoTo SetText Else: Call ClickMenu(yahoobuddymain&, "&Login...")

SetText:
editx = FindWindowEx(x, 0&, "edit", vbNullString)
Call SetFocusAPI(editx)
Call SendMessageByString(editx, WM_SETTEXT, 0&, UserName$)

editx = FindWindowEx(x, editx, "edit", vbNullString)
Call SetFocusAPI(editx)
Call SendMessageByString(editx, WM_SETTEXT, 0&, Password$)

button& = FindWindowEx(x&, 0&, "Button", vbNullString)
If SaveID = True Then: Call SendMessage(button&, BM_SETCHECK, True, 0&)

button = FindWindowEx(x, button, "button", vbNullString)

If AutoLogin = True Then: Call SendMessageLong(button, BM_SETCHECK, True, 0&)

button = FindWindowEx(x, 0&, "button", vbNullString)
button = FindWindowEx(x, button, "button", vbNullString)
button = FindWindowEx(x, button, "button", vbNullString)

If Invisible = True Then: Call SendMessageLong(button, BM_SETCHECK, True, 0&)

button = FindWindowEx(x, 0&, "button", vbNullString)
button = FindWindowEx(x, button, "button", vbNullString)
button = FindWindowEx(x, button, "button", vbNullString)
button = FindWindowEx(x, button, "button", vbNullString)

Call SetFocusAPI(button&)

Call Pause(0.3)

Call ClickButton(button&)
End Sub

Public Function AntiError()
Dim child As Long
child& = FindWindow("#32770", "Chat Error")
Call SendMessage(child&, WM_CLOSE, 0&, 0&)
'Closes Chat Error
End Function
Public Function AntiLagg()
Dim t As Integer
Do: DoEvents
t% = t% + 1
If t% = 50 Then Exit Do
Loop
End Function
Public Sub MassPM(List As ListBox, Message As String, Follow As Boolean)

Dim Scrll As Integer, num As Integer, Str As String

num% = 0

For Scrll% = 0 To List.ListCount - 1
    Str$ = List.List(Scrll%)

        If num% >= 5 Then
            Pause (3)
            num% = 0
        End If
        If Follow = True Then: Call SendPM(Str$, Message$, True)
        If Follow = False Then: Call SendPM(Str$, Message$, False)
        Pause (0.2)
    num% = num% + 1
    DoEvents
Next

End Sub
Public Sub MassPMBoot(List As ListBox, Message As String, anti As Boolean)

Dim Scrll As Integer, num As Integer, Str As String

num% = 0

For Scrll% = 0 To List.ListCount - 1
    Str$ = List.List(Scrll%)

        If num% >= 5 Then
            Pause (3)
            num% = 0
        End If
        If anti = True Then: Call SendPMBoot(Str$, Message$, True, False)
        'Determins what the boolean's are set to.
        If anti = False Then: Call SendPMBoot(Str$, Message$, False, False)
        Pause (0.2)
    num% = num% + 1
    DoEvents
Next

End Sub

Function GetChatName()
Dim imclass As Long
Dim Str As String
imclass& = FindWindow(imclass&, vbNullString)
Str$ = GetCaption(imclass&)
'Get's Caption
GetChatName = Replace(Str$, "-- Chat", "")
'Get's Caption Filterd and returns caption w/ out Chat
End Function
Function GetPMName()
Dim imclass As Long
Dim Str As String
imclass& = FindWindow(imclass&, vbNullString)
Str$ = GetCaption(imclass&)
GetChatName = Replace(Str$, "-- Instant Message", "")
End Function
Function Lagg(TheText As String)
'ex: call sendtext(lagg(UnSaKreD))
Dim g As String, a As String
Dim w As Long
Dim r$
Dim u$
Dim t$
Dim p$
g$ = TheText
Dim s$
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<html></<html></html><html></html><html></html><html></html>" & r$ & "<html></<html></html><html></html><html></html><html></html>" & u$ & "<html></<html></html><html></html><html></html><html></html>" & s$ & "<html></<html></html><html></html><html></html><html></html>" & t$
Next w
Lagg = p$
End Function
