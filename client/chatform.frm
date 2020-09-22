VERSION 5.00
Begin VB.Form chatform 
   Caption         =   "Chat Room"
   ClientHeight    =   7995
   ClientLeft      =   3765
   ClientTop       =   2280
   ClientWidth     =   6645
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   6645
   Begin VB.ListBox List1 
      Height          =   6885
      Left            =   4800
      TabIndex        =   3
      Top             =   180
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send "
      Height          =   495
      Left            =   5220
      TabIndex        =   2
      Top             =   7380
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   7035
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   180
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   7380
      Width           =   5055
   End
End
Attribute VB_Name = "chatform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

MDIForm1.sckClient.SendData "chatcode1||" & Text3.Text & "||" & strHandle

End Sub

Private Sub Form_Load()

vntArray = Split(strUserList, "||")

nItems = UBound(vntArray)

For n = 1 To nItems - 1
    List1.AddItem vntArray(n)
Next n

strChatFormState = "Open"

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Set ths variable so that I know not to print
'incoming chat text. I may re write it so that
'the server doesnt send it either.
strChatFormState = "Closed"

Set chatform = Nothing

End Sub

