VERSION 5.00
Begin VB.Form frmVote 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Dim sId As String, iVal As Integer
  sId = "11150"
  For iVal = 0 To 4
      If Option1(iVal).Value Then Exit For
  Next iVal
  Label1 = "Voting... Please wait"
'  VoteIt sId, iVal
  VoteIt sId, 0
  Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 0 Then Command2.Visible = False
End Sub

Private Sub Form_Load()
   Dim s(4) As String
   s(1) = "Perfectly"
   s(2) = "Superb"
   s(3) = "Great"
   s(4) = "Incredibly"
   With Label1
      .ForeColor = vbRed
      .Font.Name = "Times New Roman"
      .Font.Size = 22
      .Font.Italic = True
      .Font.Bold = True
      .Alignment = 2
      .Caption = "Don't forget to vote this code !"
   End With
   Option1(0).Caption = "Excelent"
   Dim offset As Long
   offset = Option1(0).Width
   For i = 1 To 4
      Load Option1(i)
      Option1(i).Left = Option1(i - 1).Left + offset
      Option1(i).Caption = s(i)
      Option1(i).Visible = True
   Next i
   Command1.Caption = "Vote It!"
   Command2.Caption = "Cancel"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  With Command2
       If X < .Left Or X > .Left + .Width Or Y < .Top Or Y > .Top + Height Then
          .Visible = True
       End If
  End With
End Sub

'Private Sub Option1_Click(Index As Integer)
'  If Index > 0 Then Option1(0).Value = True
'End Sub
