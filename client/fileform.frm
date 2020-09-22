VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fileform 
   Caption         =   "File Section"
   ClientHeight    =   8820
   ClientLeft      =   8475
   ClientTop       =   1425
   ClientWidth     =   4200
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   4200
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   8160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":0CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":1148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":159A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7740
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSWinsockLib.Winsock ftpclient 
      Left            =   3780
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7335
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   12938
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File Size"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Double Click File To Download"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "fileform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ftpClient_ConnectionRequest(ByVal requestID As Long)

ftpclient.Close
ftpclient.Accept requestID

End Sub

Private Sub ftpClient_DataArrival(ByVal bytesTotal As Long)

Dim data As String

ftpclient.GetData data

'MsgBox bytesTotal

Put #fFile, , data

If strFileLen = Loc(fFile) Then
    Close #fFile
    ftpclient.Close
    ftpclient.Listen
End If

End Sub

Private Sub Form_Load()

ftpclient.LocalPort = "21"
ftpclient.Listen

MDIForm1.sckClient.SendData "filelistcode1||"

End Sub

Private Sub ListView1_DblClick()

MDIForm1.sckClient.SendData "getfilecode1||" & fileform.ListView1.SelectedItem

fFile = FreeFile

strFileLen = fileform.ListView1.SelectedItem.SubItems(1)

Open App.Path & "\dl\" & fileform.ListView1.SelectedItem For Binary Access Write As #fFile


End Sub
