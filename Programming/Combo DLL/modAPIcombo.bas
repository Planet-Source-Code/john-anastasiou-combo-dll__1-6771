Attribute VB_Name = "modAPIcombo"
Public Declare Function SendMessageAsString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Long, lParam As Any) As Long
  
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_GETDROPPEDWIDTH = &H15F
Public Const CB_SETDROPPEDWIDTH = &H160
Public Const CB_LIMITTEXT = &H141
Public Const CB_SETITEMHEIGHT = &H153
Public Const CB_GETITEMHEIGHT = &H154

