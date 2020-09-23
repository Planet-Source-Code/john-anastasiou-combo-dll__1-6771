VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoComplete"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Function FindString(Combo As Object, ByVal mString As String, SelectIt As Boolean) As Integer
Dim lIndex As Integer
lIndex = SendMessageAsString(Combo.hWnd, CB_FINDSTRING, -1, mString & vbNullString)
FindString = lIndex
If SelectIt = True And Combo.ListCount > 0 Then
Combo.ListIndex = lIndex
End If

End Function

Public Sub SetItemHeight(Combo As Object, ByVal nItem As Integer)
SendMessage Combo.hWnd, CB_SETITEMHEIGHT, nItem, nHeight

End Sub

Public Function GetEditBoxHeight(Combo As Object) As Long
ret& = SendMessage(Combo.hWnd, CB_GETITEMHEIGHT, -1, 0)
GetEditBoxHeight = ret&
End Function

Public Function SetEditBoxHeight(Combo As Object, ByVal nHeight As Long) As Long
SetEditBoxHeight = SendMessage(Combo.hWnd, CB_SETITEMHEIGHT, -1, nHeight)
End Function

Public Sub SetMaxLen(Combo As Object, ByVal MaxChars As Long)
If MaxChars > 0 Then
    SendMessage Combo.hWnd, CB_LIMITTEXT, MaxChars, 0
Else
    SendMessage Combo.hWnd, CB_LIMITTEXT, 0, 0 'reset to default
End If
End Sub

Public Sub SetDroppedWidth(Combo As Object, ByVal nWidth As Integer)
SendMessage Combo.hWnd, CB_SETDROPPEDWIDTH, nWidth, 0

End Sub

Public Function GetDroppedWidth(Combo As Object) As Long
GetDroppedWidth = SendMessage(Combo.hWnd, CB_GETDROPPEDWIDTH, 0, 0)

End Function

Public Function GetDroppedState(Combo As Object) As Boolean
GetDroppedState = SendMessage(Combo.hWnd, CB_GETDROPPEDSTATE, 0, 0)

End Function








Public Sub Drop(ByRef Combo As Object, ByVal n As Boolean)
Attribute Drop.VB_Description = "Drop the comboBox"
    Call SendMessage(Combo.hWnd, CB_SHOWDROPDOWN, n, ByVal 0&)

End Sub

Public Sub AutoComplete(ByRef Combo As Object)
Dim lIndex As Integer
Dim iLen As Integer
Dim sTypedText As String
Const CB_LOCKED = &H255

    With Combo
        If ((InStr(1, .Text, .Tag, vbTextCompare) <> 0 And Len(.Tag) = Len(.Text) - 1) Or (Left(.Text, 1) <> Left(.Tag, 1) And .Tag <> "")) And .Tag <> CStr(CB_LOCKED) Then
            iLen = Len(.Text)
            lIndex = SendMessageAsString(.hWnd, CB_FINDSTRING, -1, .Text & vbNullString)
               If lIndex >= 0 Then
                sTypedText = .Text
                .Tag = CB_LOCKED
                .ListIndex = lIndex
                .Tag = sTypedText
                .SelStart = iLen
                .SelLength = Len(.List(.ListIndex)) - iLen
                Debug.Print "Found"
            End If
            Exit Sub
        ElseIf .Tag <> CStr(CB_LOCKED) Then
            .Tag = .Text
            Debug.Print "Not Found"
        End If
    End With
End Sub








