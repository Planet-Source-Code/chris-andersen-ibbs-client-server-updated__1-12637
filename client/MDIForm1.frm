VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "iBBS Client"
   ClientHeight    =   7935
   ClientLeft      =   1305
   ClientTop       =   1800
   ClientWidth     =   12315
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   10380
      Top             =   300
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   10920
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnudisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuconnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "View"
      Begin VB.Menu mnuchat 
         Caption         =   "Chat"
      End
      Begin VB.Menu mnufiles 
         Caption         =   "Files"
      End
      Begin VB.Menu mnuim 
         Caption         =   "Instant Message"
      End
      Begin VB.Menu mnumb 
         Caption         =   "Message Forum"
      End
      Begin VB.Menu mnumail 
         Caption         =   "Check Mailbox"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'the client is fairly straight forward.
Dim strUserList As String
Dim IMWindowFound As Boolean
Dim itm1 As ListItem
Dim lngIcon As Long

Private Sub MDIForm_Load()

With sckClient
    .RemoteHost = frmLogin.txtIP.Text
    .RemotePort = "1001"
    .Connect
End With

Load Userlist
Userlist.Show
'Preset this variable so that incoming chat data doesnt print to the chat window
'until it has been opened
strChatFormState = "Closed"

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End

End Sub

Private Sub mnuchat_Click()

Load chatform
With chatform
    .Height = 8415
    .Width = 6705
    .Show
End With

End Sub

Private Sub mnuconnect_Click()

MDIForm1.Hide
Load frmLogin
frmLogin.Show

End Sub

Private Sub mnudisconnect_Click()

sckClient.Close

'For Each frm In Forms
'    If frm.Name <> "MDIForm1" Then
        'Unload frm.Name
'    End If
'Next

End Sub

Private Sub mnufiles_Click()

Load fileform

With fileform
    .Height = 8010
    .Width = 4350
    .Show
End With

End Sub

Private Sub mnuim_Click()

Load IMForm(IMNumber)
IMForm(IMNumber).Show
IMNumber = IMNumber + 1

End Sub

Private Sub mnuquit_Click()

End

End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)

Dim strSendCode As String
Dim vntArray As Variant
Dim strText As String
Dim nItems As Integer
Dim n As Integer

sckClient.GetData strSendCode, vbString

' split function will be used to parse items contained in a string,
' and delimitted by ||
' The Split function returns a variant array containing each parsed item
' as an element in the array

' use split function to parse it
vntArray = Split(strSendCode, "||")

' how many items were parsed?
nItems = UBound(vntArray)
  'Text1.Text = strsendcode
  

Select Case vntArray(0)
    Case "admin1"
        'Incoming Admin message
        strtest = MsgBox(vntArray(1), vbCritical, "Message from the Administrator")
    Case "connect1"
        'Determine if the server is allowing access
        If vntArray(1) = "logonyes" Then
            MDIForm1.Show
            Unload frmLogin
        Else:
            MsgBox ("Login Incorrect!")
        End If
        
    Case "imcode2"
        'MsgBox ("<" & vntArray(2) & ">" & vntArray(1))
        'First check if window for that IM is already open
        'If it is, send the text to text1
        'If not create a new IM window then send text to it
        IMWindowFound = False
        
        
        For Each frm In Forms
            'DoEvents
            If frm.Caption = vntArray(2) Then
                'IM already with this user.
                frm.Text1.Text = frm.Text1.Text & vbCrLf & "<" & vntArray(2) & ">" & vntArray(1)
                IMWindowFound = True
                Exit For
            End If
        Next
        
        If IMWindowFound = False Then
            'New User IM. Create a new IM Window
            Load IMForm(IMNumber)
            
            With IMForm(IMNumber)
                .Text1.Text = IMForm(IMNumber).Text1.Text & vbCrLf & "<" & vntArray(2) & ">" & vntArray(1)
                .Caption = vntArray(2)
                .Height = 4350
                .Width = 6090
                .Show
            End With
            
            IMNumber = IMNumber + 1
        End If
        
        IMWindowFound = False
        
    Case "chatcode2"
        Dim strChatHandle As String
        Dim strMessage As String
        
        'Handle incoming chat info
        strChatHandle = vntArray(1)
        strMessage = vntArray(2)
        
        If strChatFormState <> "Closed" Then
            chatform.Text1.Text = chatform.Text1.Text & "<" & strChatHandle & ">" & strMessage & vbCrLf
        End If
    Case "filelistcode2"
        nItems = UBound(vntArray)

        ' display each file available for download on the server
       
        fileform.ListView1.ListItems.Clear
        For n = 1 To nItems - 1
            vntarray2 = Split(vntArray(n), "\/")
            
            
            Set itm = fileform.ListView1.ListItems.Add(, , vntarray2(1))
            itm.SubItems(1) = vntarray2(0)
            'Change icon for file in list depending on its extension
            Select Case LCase(Right(vntarray2(1), 3))
                Case "txt"
                    itm.SmallIcon = 1
                Case "mp3"
                    itm.SmallIcon = 2
                Case "wav", "mid"
                    itm.SmallIcon = 3
                Case "zip"
                    itm.SmallIcon = 4
                Case "jpg", "gif"
                    itm.SmallIcon = 5
                Case Else
                    itm.SmallIcon = 6
            End Select
            
        Next n
        Set itm = Nothing
        
    Case "mbmessagescode2"
        'Implementation on hold for now pending better way to do it
        'nItems = UBound(vntArray)

        ' display each parsed item
       
        'ListView2.ListItems.Clear
        'For n = 1 To nItems - 1
'            vntarray2 = Split(vntArray(n), "\/")
'
'            Text1.Text = Text1.Text & vbCrLf & vntArray(n)
'            Set itm = ListView2.ListItems.Add(, , vntarray2(0))
'            itm.SubItems(1) = vntarray2(1)
'            itm.SubItems(2) = vntarray2(2)
'            itm.SubItems(3) = vntarray2(3)
'            itm.SubItems(4) = vntarray2(4)
'        Next n
'
'        Set itm = Nothing
        
    Case "userlistcode2"
        nItems = UBound(vntArray)
        
        If strChatFormState <> "Closed" Then
            chatform.List1.Clear
        End If
  
        Userlist.userlist1.ListItems.Clear
        
        For n = 1 To nItems - 1
            strUserList = strUserList & vntArray(n) & "||"
            If strChatFormState <> "Closed" Then
               chatform.List1.AddItem vntArray(n)
            End If
            
            Set itm1 = Userlist.userlist1.ListItems.Add(, , vntArray(n))
            
        Next n
        
        
End Select

End Sub

Private Sub Timer1_Timer()

'Send Userlist request if User's window is open
If sckClient.State = 7 Then
    sckClient.SendData "userlist1"
End If

End Sub
