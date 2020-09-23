Attribute VB_Name = "MOD"
Global NewBoo(30) As New PM
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
Sub Send(Data As String)
    With Login
        Select Case .WS.State
            Case Is = sckConnected
                .WS.SendData (Data$)
            Case Is <> sckConnected
                Status "Not Connected"
        End Select
    End With
End Sub
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
Sub Status(dat As String)
    Mess.stat.Panels.item(1).Text = "Status : " & dat$
End Sub
Sub NewError(whatError As String)
    Error.lblNewError.Caption = whatError$
    Error.Show
End Sub

