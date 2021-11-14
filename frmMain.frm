VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ICQ Control Center"
   ClientHeight    =   6420
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10830
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtURL 
      Height          =   345
      Left            =   7095
      TabIndex        =   41
      Text            =   "www.ic-crypt.org.uk"
      Top             =   5940
      Width           =   3585
   End
   Begin VB.CommandButton cmdURL 
      Caption         =   "SEND &URL"
      Height          =   315
      Left            =   7095
      TabIndex        =   40
      Top             =   4950
      Width           =   1305
   End
   Begin VB.TextBox txtURLUIN 
      Height          =   285
      Left            =   9570
      MaxLength       =   9
      TabIndex        =   38
      Text            =   "14996057"
      Top             =   4980
      Width           =   1095
   End
   Begin VB.TextBox textPagerFrom 
      Height          =   285
      Left            =   9540
      TabIndex        =   36
      Text            =   "me@me.com"
      Top             =   675
      Width           =   1095
   End
   Begin VB.TextBox FromName 
      Height          =   285
      Left            =   9540
      MaxLength       =   9
      TabIndex        =   35
      Text            =   "anonymous"
      Top             =   225
      Width           =   1095
   End
   Begin VB.TextBox TextUIN 
      Height          =   285
      Left            =   9540
      MaxLength       =   9
      TabIndex        =   33
      Text            =   "14996057"
      Top             =   1110
      Width           =   1095
   End
   Begin VB.TextBox TextSubject 
      Height          =   315
      Left            =   7110
      MaxLength       =   30
      TabIndex        =   28
      Top             =   1830
      Width           =   3540
   End
   Begin VB.CommandButton BtnSend 
      Caption         =   "&Send www pager"
      Height          =   315
      Left            =   9015
      TabIndex        =   27
      Top             =   3750
      Width           =   1680
   End
   Begin VB.TextBox TextMessage 
      Height          =   930
      Left            =   7125
      MaxLength       =   450
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   2520
      Width           =   3525
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   10965
      TabIndex        =   25
      Top             =   4035
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Advanced Mode - >>"
      Height          =   315
      Left            =   4995
      TabIndex        =   23
      Top             =   2940
      Width           =   1830
   End
   Begin VB.Frame Frame3 
      Caption         =   "Owner Details"
      Height          =   2580
      Left            =   165
      TabIndex        =   16
      Top             =   105
      Width           =   1830
      Begin VB.TextBox txtOwnerIP 
         Height          =   315
         Left            =   165
         TabIndex        =   22
         Top             =   2040
         Width           =   1305
      End
      Begin VB.TextBox txtOwnerUIN 
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   1320
         Width           =   1305
      End
      Begin VB.TextBox txtOwnerNickname 
         Height          =   315
         Left            =   195
         TabIndex        =   18
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Owner IP ADDY"
         Height          =   195
         Left            =   195
         TabIndex        =   21
         Top             =   1785
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Owner UIN NO"
         Height          =   195
         Left            =   210
         TabIndex        =   19
         Top             =   1065
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Owner Nickname"
         Height          =   195
         Left            =   225
         TabIndex        =   17
         Top             =   345
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Send Message"
      Height          =   1605
      Left            =   2325
      TabIndex        =   11
      Top             =   1110
      Width           =   4500
      Begin VB.TextBox txtUIN 
         Height          =   315
         Left            =   3315
         TabIndex        =   14
         Top             =   1065
         Width           =   1065
      End
      Begin VB.TextBox txtMessage 
         Height          =   1050
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   330
         Width           =   2925
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send"
         Height          =   315
         Left            =   3315
         TabIndex        =   12
         Top             =   345
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "TO UIN :"
         Height          =   195
         Left            =   3330
         TabIndex        =   15
         Top             =   780
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Last Recieved File"
      Height          =   825
      Left            =   2355
      TabIndex        =   9
      Top             =   120
      Width           =   4470
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   195
         TabIndex        =   10
         Top             =   315
         Width           =   4065
      End
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "ABOUT"
      Height          =   315
      Left            =   4995
      TabIndex        =   8
      Top             =   5970
      Width           =   1830
   End
   Begin VB.CommandButton cmdStatus 
      Caption         =   "Status - ONLINE"
      Height          =   315
      Left            =   285
      TabIndex        =   7
      Top             =   5955
      Width           =   1830
   End
   Begin VB.ListBox List3 
      Height          =   2010
      Left            =   4965
      TabIndex        =   5
      Top             =   3765
      Width           =   1830
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
      Height          =   315
      Left            =   2550
      TabIndex        =   2
      Top             =   5955
      Width           =   1845
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   2535
      TabIndex        =   1
      Top             =   3735
      Width           =   1830
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   240
      TabIndex        =   0
      Top             =   3735
      Width           =   1830
   End
   Begin MSWinsockLib.Winsock SockPager 
      Left            =   8280
      Top             =   1380
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Send a web address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7095
      TabIndex        =   43
      Top             =   4290
      Width           =   3120
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "URL TO SEND"
      Height          =   195
      Left            =   7095
      TabIndex        =   42
      Top             =   5535
      Width           =   1095
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "TO ICQ UIN:"
      Height          =   195
      Left            =   8520
      TabIndex        =   39
      Top             =   5040
      Width           =   915
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Pager From e-mail address :"
      Height          =   195
      Left            =   7110
      TabIndex        =   37
      Top             =   735
      Width           =   1950
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Send pager to ICQ UIN:"
      Height          =   195
      Left            =   7110
      TabIndex        =   34
      Top             =   1155
      Width           =   1695
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   7095
      TabIndex        =   32
      Top             =   1530
      Width           =   600
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Message:"
      Height          =   195
      Left            =   7110
      TabIndex        =   31
      Top             =   2280
      Width           =   690
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   " From Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6990
      TabIndex        =   30
      Top             =   165
      Width           =   1905
   End
   Begin VB.Label LabelStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   7125
      TabIndex        =   29
      Top             =   3735
      Width           =   1650
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Advanced mode enables use of wwwpager and much more :)"
      Height          =   195
      Left            =   225
      TabIndex        =   24
      Top             =   2955
      Width           =   4335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Contact UIN Number"
      Height          =   195
      Left            =   2535
      TabIndex        =   6
      Top             =   3420
      Width           =   1590
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contact IP address"
      Height          =   195
      Left            =   4995
      TabIndex        =   4
      Top             =   3450
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contact NickName"
      Height          =   195
      Left            =   255
      TabIndex        =   3
      Top             =   3420
      Width           =   1530
   End
   Begin VB.Menu mnuOnline 
      Caption         =   "online"
      Visible         =   0   'False
      Begin VB.Menu mnuOffline 
         Caption         =   "Offline/Disconnect"
      End
      Begin VB.Menu mnuPriviacy 
         Caption         =   "Privacy (Invisible)"
      End
      Begin VB.Menu mnuDND 
         Caption         =   "DND (Do Not Disturb)"
      End
      Begin VB.Menu mnuOccupied 
         Caption         =   "Occupied (Urgent msgs)"
      End
      Begin VB.Menu mnuNA 
         Caption         =   "N/A (Extended away)"
      End
      Begin VB.Menu mnuAWAY 
         Caption         =   "Away"
      End
      Begin VB.Menu mnuChat 
         Caption         =   "Free For Chat"
      End
      Begin VB.Menu mnuAvailable 
         Caption         =   "Available / Connect"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemNumber As Long
Dim SelectedUIN As Long

Private Sub BtnSend_Click()
 On Error Resume Next
   
   Dim cSend As String
   Dim cFrom As String
   
   Dim cData As String
   
   ' Verify datas
   If Not IsNumeric(TextUIN.Text) Then
      MsgBox "The ICQ UIN not Numeric !"
         
      TextUIN.SetFocus
      Exit Sub
   End If
   
   'If CStr(Val(TextUIN.Text)) = "14996057" Then
   '   MsgBox "Please... Don't Test With my UIN"
         
   '   TextUIN.SetFocus
   '   Exit Sub
   'End If
         
   If Trim(TextMessage.Text) = "" Then
      MsgBox "Don't Allow Blank Messages"
         
      TextMessage.SetFocus
      Exit Sub
   End If

   ' Status
   LabelStatus.Caption = "Starting..."
   
   ' Close Socket
   SockPager.Close
      
   ' Change the " " for "+"
   cMail = ChangeSpaces(textPagerFrom.Text)
   cFrom = ChangeSpaces(FromName.Text)
   cSubject = ChangeSpaces(TextSubject.Text)
   cMessage = ChangeSpaces(TextMessage.Text)

   ' Fill the String
   cData = "from=" + cFrom + "&fromemail=" & cMail & "&subject=" & cSubject & "&body=" & cMessage & "&to=" & Trim(TextUIN.Text) & "&Send=" & """"
      

   cSend = "POST /scripts/WWPMsg.dll HTTP/1.0" & vbCrLf
   cSend = cSend & "Referer: http://wwp.mirabilis.com" & vbCrLf
   cSend = cSend & "User-Agent: Mozilla/4.06 (Win95; I)" & vbCrLf
   cSend = cSend & "Connection: Keep-Alive" & vbCrLf
   cSend = cSend & "Host: wwp.mirabilis.com:80" & vbCrLf
   cSend = cSend & "Content-type: application/x-www-form-urlencoded" & vbCrLf
   cSend = cSend & "Content-length: " & Len(cData) & vbCrLf
   cSend = cSend & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf & vbCrLf
   cSend = cSend & cData & vbCrLf & vbCrLf & vbCrLf & vbCrLf

   ' Send Message
   SockPager.Tag = cSend
   SockPager.Connect "wwp.mirabilis.com", 80
End Sub

Private Sub cmdAbout_Click()
frmAbout.Show

End Sub

Private Sub cmdRecent_Click()
PopupMenu mnuRecentFileZ

End Sub

Private Sub cmdURL_Click()
Rtn = SendURL(Val(txtURLUIN.Text), txtURL.Text)

End Sub

Private Sub Command1_Click()
Dim iUIN As Long
iUIN = Val(txtUIN.Text)

Rtn = SendMessage(iUIN, txtMessage.Text)

'Rtn = SendExternal(iUIN, "ICQ Control Center", txtMessage.Text, 1)

'NOTE : Send External is NOT as most people think
'people including myself at one point thought this
'was a method for sending external messages without
'ICQ Interaction, all send external actually does
'is to send an external chat request, if the person
'on the other side has not got this app configured
'as a legitamate external application then this is
'even more pointless.



End Sub

Private Sub Command2_Click()

If Me.Width = 7035 Then
Command2.Caption = "Normal mode <<-"
Me.Width = 10920
Me.Left = Screen.Width / 2 - Me.Width / 2

Else
Command2.Caption = "Advanced mode ->>"

Me.Width = 7035
Me.Left = Screen.Width / 2 - Me.Width / 2

End If


End Sub



Private Sub List2_DblClick()
'Call DisplayUserInfo(SelectedUIN)


End Sub

Private Sub SockPager_Connect()
   On Error Resume Next
   
   ' Status
   LabelStatus.Caption = "Sending..."
  
   SockPager.SendData SockPager.Tag
End Sub

Private Sub SockPager_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   ' Status
   LabelStatus.Caption = "Error..."
   
   SockPager.Tag = ""
End Sub

Private Sub SockPager_SendComplete()
   ' Status
   LabelStatus.Caption = "www pager sent..."
      
   SockPager.Tag = ""
End Sub

Private Function ChangeSpaces(cString As String) As String
   On Error Resume Next
  
   ' Variaveis
   Dim cChar As String
   Dim cReturn As String
  
   Dim nLoop As Long
  
   ' Faz a Troca
   cReturn = ""
  
   For nLoop = 1 To Len(cString)
       cChar = Mid(cString, nLoop, 1)
      
       If cChar = " " Then
          cChar = "+"
       End If
      
       cReturn = cReturn + cChar
   Next
  
   ChangeSpaces = cReturn
End Function


Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub cmdStatus_Click()
PopupMenu mnuOnline

End Sub

Private Sub Form_Load()
'(C)


  Me.Width = 7035
  Me.Left = Screen.Width / 2 - Me.Width / 2
  Me.Top = Screen.Height / 2 - Me.Height / 2
  

  sName = "Visual Basic"
  sPassword = "aaaaaaaa"
  sLicense = "E94AD7C14D1DBAE8"
   
  Rtn = SetLicenseKey(sName, sPassword, sLicense)
  Rtn = GetVersion(iCQversion)
  

  Call GetUserList
  Call SendNotify
  
  iCQowner = GetOwnerUin
  iCQownerNick = GetOwnerNickName
  
  txtOwnerNickname.Text = GetOwnerNickName
  txtOwnerUIN.Text = iCQowner
  txtOwnerIP.Text = GetOwnerIP
    

  textPagerFrom.Text = GetOwnerEmail
  
  
  
  Me.Caption = iCQowner & " - [ ICQ Control Center ]"
  
  
  
 X = GetOwnerStatus
   
 If X = BICQAPI_USER_STATE_ONLINE Then cmdStatus.Caption = "Status - ONLINE"
 If X = BICQAPI_USER_STATE_CHAT Then cmdStatus.Caption = "Status - Free For Chat"
 If X = BICQAPI_USER_STATE_AWAY Then cmdStatus.Caption = "Status - Away"
 If X = BICQAPI_USER_STATE_NA Then cmdStatus.Caption = "Status - N/A [Extended Away]"
 If X = BICQAPI_USER_STATE_OCCUPIED Then cmdStatus.Caption = "Status - Occupied [Urgent MSGS]"
 If X = BICQAPI_USER_STATE_DND Then cmdStatus.Caption = "Status - DND [Do Not Disturb]"
 If X = BICQAPI_USER_STATE_INVISIBLE Then cmdStatus.Caption = "Status - Privacy [Invisible]"
 If X = BICQAPI_USER_STATE_OFFLINE Then cmdStatus.Caption = "Status - OFFLINE"

 
 
 
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
Rtn = UnRegisterNotify

End Sub

Private Sub List1_Click()
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
List2.Selected(i) = True
List3.Selected(i) = True

End If
Next i

End Sub

Private Sub List2_Click()
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
txtUIN.Text = List2.List(i)
TextUIN.Text = List2.List(i)
txtURLUIN.Text = List2.List(i)
SelectedUIN = List2.List(i)

List1.Selected(i) = True
List3.Selected(i) = True
End If
Next i


End Sub

Private Sub List3_Click()
For i = 0 To List3.ListCount - 1
If List3.Selected(i) = True Then
List2.Selected(i) = True
List3.Selected(i) = True
End If
Next i

End Sub

Private Sub mnuAvailable_Click()
X = SetOwnerState(BICQAPI_USER_STATE_ONLINE)


End Sub

Private Sub mnuAWAY_Click()
X = SetOwnerState(BICQAPI_USER_STATE_AWAY)

End Sub

Private Sub mnuChat_Click()
X = SetOwnerState(BICQAPI_USER_STATE_CHAT)

End Sub

Private Sub mnuDND_Click()
X = SetOwnerState(BICQAPI_USER_STATE_DND)

End Sub

Private Sub mnuNA_Click()
X = SetOwnerState(BICQAPI_USER_STATE_NA)

End Sub

Private Sub mnuOccupied_Click()
X = SetOwnerState(BICQAPI_USER_STATE_OCCUPIED)

End Sub

Private Sub mnuOffline_Click()
X = SetOwnerState(BICQAPI_USER_STATE_OFFLINE)

End Sub

Private Sub mnuOnline_Click()
X = SetOwnerState(BICQAPI_USER_STATE_ONLINE)

End Sub

Private Sub mnuPriviacy_Click()
X = SetOwnerState(BICQAPI_USER_STATE_INVISIBLE)

End Sub

