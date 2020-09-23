VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "password thingie"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   6435
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "&New"
      Height          =   315
      Left            =   5280
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "User Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim log As String
On Error GoTo errmsg
Open "C:\windows\" & Text1.Text & ".pll" For Input As #1
Input #1, log
Close #1
If log = Text1.Text Then
Dim logg As String
On Error GoTo errmsg
Open "C:\windows\" & Text2.Text & ".pli" For Input As #1
Input #1, logg
Close #1
If logg = Text2.Text Then
MsgBox "Login succsessfull " & "[" & Text1.Text & "]" & " Please proceed", vbOKOnly Or vbInformation, "Logged in!": GoTo ed
End If
End If
errmsg: MsgBox "You do not have an account created...please create one by clicking on NEW", vbCritical, "Error"
ed:
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Dim user
Dim pat
Dim user1
retry:
user = InputBox("Enter a username that you would like to use")
If user = "" Then
GoTo retry
Else:
Open "C:\windows\" & user & ".pll" For Output As #1
Write #1, user
Close #1
retry1:
user1 = InputBox("Enter a password that you would like to use")
If user1 = "" Then
GoTo retry1
Else:
Open "C:\windows\" & user1 & ".pli" For Output As #1
Write #1, user1
Close #1
End If
End If
End Sub



Private Sub Command3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub
