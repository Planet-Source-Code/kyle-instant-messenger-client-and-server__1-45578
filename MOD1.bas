Attribute VB_Name = "MOD1"
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As enm_NIM_Shell, pnid As NOTIFYICONDATA) As Boolean
Public nidProgramData As NOTIFYICONDATA
Public Const WM_MOUSEISMOVING = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_SETHOTKEY = &H32
Public Type NOTIFYICONDATA
         cbSize As Long
         hwnd As Long
         uId As Long
         uFlags As Long
         uCallbackMessage As Long
         hIcon As Long
         szTip As String * 64
End Type
Public Enum enm_NIM_Shell
         NIM_ADD = &H0
         NIM_MODIFY = &H1
         NIM_DELETE = &H2
         NIF_MESSAGE = &H1
         NIF_ICON = &H2
         NIF_TIP = &H4
         WM_MOUSEMOVE = &H200
End Enum
Sub RemoveItemFromListbox(lst As ListBox, item$)
Do
NoFreeze% = DoEvents()
If LCase$(lst.List(a)) = LCase$(item$) Then lst.RemoveItem (a)
a = 1 + a
Loop Until a >= lst.ListCount
End Sub
Sub Pause(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub
Sub Status(Dat As String)
    Server.stat.Panels.item(1).Text = "Status : " & Dat$
End Sub
Sub Log(Who As String, ip As String)
    Logg.txtLog.SelStart = Len(Logg.txtLog.Text)
    Logg.txtLog.Text = Logg.txtLog.Text & "Logged : " & Who & " has Logged onto Server." & vbCrLf & "IP : " & ip & vbCrLf & vbCrLf
End Sub
Sub AddData(Daz As String, Data2Add As String)
    Data.txtData.SelStart = Len(Data.txtData.Text)
    Data.txtData.Text = Data.txtData.Text & Daz & ":" & Data2Add & vbCrLf
End Sub
