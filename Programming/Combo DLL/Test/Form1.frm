VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Selected if found"
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1725
   End
   Begin VB.TextBox Text4 
      Height          =   330
      Left            =   105
      TabIndex        =   12
      Text            =   "Item 2"
      Top             =   1440
      Width           =   1710
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Find String"
      Height          =   285
      Left            =   105
      TabIndex        =   11
      Top             =   855
      Width           =   1725
   End
   Begin VB.CommandButton Command7 
      Caption         =   "GetEditHeight"
      Height          =   285
      Left            =   2535
      TabIndex        =   10
      Top             =   1440
      Width           =   1725
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SetEditHeight"
      Height          =   285
      Left            =   2535
      TabIndex        =   9
      Top             =   1140
      Width           =   1725
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4245
      TabIndex        =   8
      Text            =   "5"
      Top             =   1155
      Width           =   435
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4230
      TabIndex        =   7
      Text            =   "5"
      Top             =   870
      Width           =   435
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SetMaxChars"
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Top             =   855
      Width           =   1725
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show Dropped Width"
      Height          =   270
      Left            =   2520
      TabIndex        =   5
      Top             =   0
      Width           =   2145
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SetDroppedWidth"
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   585
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4230
      TabIndex        =   3
      Text            =   "100"
      Top             =   585
      Width           =   435
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   915
      Top             =   2490
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Drop"
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   285
      Width           =   2145
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   555
      Left            =   3150
      TabIndex        =   1
      Top             =   2490
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   255
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   255
      Width           =   1560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cs As New clsAutoComplete
Private Sub Combo1_Change()
cs.AutoComplete Combo1

End Sub

Private Sub Command1_Click()
End

End Sub


Private Sub Command2_Click()
Call cs.Drop(Combo1, True)
End Sub

Private Sub Command3_Click()
cs.SetDroppedWidth Combo1, ByVal CInt(Text1.Text)
End Sub

Private Sub Command4_Click()
MsgBox cs.GetDroppedWidth(Combo1)
End Sub

Private Sub Command5_Click()
cs.SetMaxLen Combo1, CLng(Text2.Text)
End Sub


Private Sub Command6_Click()
MsgBox cs.SetEditBoxHeight(Combo1, CLng(Text3.Text))

End Sub

Private Sub Command7_Click()
MsgBox cs.GetEditBoxHeight(Combo1)
End Sub

Private Sub Command8_Click()
MsgBox cs.FindString(Combo1, Text4.Text, CBool(Check1.Value))
End Sub

Private Sub Form_Load()
For i% = 1 To 100
Combo1.AddItem "Item " & CStr(i%)
Next i%

End Sub


Private Sub Timer1_Timer()
Me.Caption = cs.GetDroppedState(Combo1)
End Sub


