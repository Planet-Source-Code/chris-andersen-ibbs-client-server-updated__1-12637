VERSION 5.00
Begin VB.Form frmIM 
   Caption         =   "Instant Message"
   ClientHeight    =   3945
   ClientLeft      =   8250
   ClientTop       =   5505
   ClientWidth     =   5970
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   5970
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   3300
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   3360
      Width           =   4635
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5835
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

MDIForm1.sckClient.SendData "imcode1||" & Text2.Text & "||" & Me.Caption & "||" & strHandle
Text1.Text = Text1.Text & vbCrLf & "<" & strHandle & ">" & Text2.Text


End Sub

