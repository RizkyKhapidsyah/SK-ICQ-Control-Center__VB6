Attribute VB_Name = "ICQBAS"
'Copyright (C) 2000 DIGITAL VAMPIRE
'DV@KNAC.COM - ICQ 14996057

'www.ic-crypt.org.uk



Declare Function SetLicenseKey Lib "icqmapi.dll" Alias "ICQAPICall_SetLicenseKey" (ByVal pszName As String, ByVal pszPassword As String, ByVal pszLicense As String) As Boolean
Declare Function SetOwnerState Lib "icqmapi.dll" Alias "ICQAPICall_SetOwnerState" (ByVal iState As Long) As Boolean

Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (Dest As Any, Src As Any, ByVal cb&)

Global iCQversion As Long
Global iCQowner As Long
Global iCQownerNick As String

Global recentcount As Long

Declare Function SendMessage Lib "icqmapi.dll" Alias _
     "ICQAPICall_SendMessage" (ByVal iUIN As Long, _
                               ByVal pszMessage As String) As Boolean
                              
Declare Function SendExternal Lib "icqmapi.dll" Alias _
     "ICQAPICall_SendExternal" (ByVal iUIN As Long, _
                                ByVal pszExternal As String, _
                                ByVal pszMessage As String, _
                                ByVal bAutoSend As Long) As Boolean

Declare Function SendURL Lib "icqmapi.dll" Alias _
     "ICQAPICall_SendURL" (ByVal iUIN As Long, _
                           ByVal pszURL As String) As Boolean



Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Global Const BICQAPI_USER_STATE_ONLINE = 0
Global Const BICQAPI_USER_STATE_CHAT = 1
Global Const BICQAPI_USER_STATE_AWAY = 2
Global Const BICQAPI_USER_STATE_NA = 3
Global Const BICQAPI_USER_STATE_OCCUPIED = 4
Global Const BICQAPI_USER_STATE_DND = 5
Global Const BICQAPI_USER_STATE_INVISIBLE = 6
Global Const BICQAPI_USER_STATE_OFFLINE = 7

Global Const SHLength = 512
Global Const SL1 = 20
Global Const SL2 = 100
Global Const SL3 = 50

Public Type BSICQAPI_User
    m_iUIN As Long                   'the user’s ICQ #.
    m_hFloatWindow As Long           'the handle of the “Float” window if the
                                     'user is floating.
    m_iIP As Long                    'the user’s IP address.
    m_szNickname As String * SL1     'the user’s nickname.
    m_szFirstName As String * SL1    'the user’s first name.
    m_szLastName As String * SL1     'the user’s last name.
    m_szEmail As String * SL2        'the user’s email address.
    m_szCity As String * SL2         'the user’s city.
    m_szState As String * SL2        'the user’s state.
    m_iCountry As Long               'the user’s country code.
    m_szCountryName As String * SL2  'the user’s country name.
    m_szHomePage As String * SL2     'the user’s homepage.
    m_iAge As Long                   'the user’s age.
    m_szPhone As String * SL1        'the user’s phone.
    m_bGender As Long                'user’s gender. The codes are:
                                     '  0 - Not Specified,
                                     '  1 - Female,
                                     '  2 - Male.
    m_iHomeZip As Long               'the user’s zip code.
    m_iStateFlags As Long            'get one of the following values :

End Type


Public Type BSICQAPI_Group
    m_szName As String * SL3
    m_iUserCount As Long
    m_ppUsers As BSICQAPI_User
End Type

Declare Function GetOnlineListDetails Lib "icqmapi.dll" Alias "ICQAPICall_GetOnlineListDetails" (ByRef iCount As Long, ByRef ppUsers As Any) As Boolean
Declare Function GetFullOwnerData Lib "icqmapi.dll" Alias "ICQAPICall_GetFullOwnerData" (ByRef ppUser As BSICQAPI_User, ByVal iVersion As Long) As Boolean

Declare Function GetWindowHandle Lib "icqmapi.dll" Alias "ICQAPICall_GetWindowHandle" (ByVal hWindow As Long) As Boolean
Declare Function GetVersion Lib "icqmapi.dll" Alias "ICQAPICall_GetVersion" (ByRef iVersion As Long) As Boolean
Declare Function GetOnlineListType Lib "icqmapi.dll" Alias "ICQAPICall_GetOnlineListType" (ByRef iListType As Long) As Boolean
Declare Function GetFullUserData Lib "icqmapi.dll" Alias "ICQAPICall_GetFullUserData" (ByRef pUser As BSICQAPI_User, ByVal iVersion As Long) As Boolean

Declare Sub FreeUser Lib "icqmapi.dll" Alias "ICQAPIUtil_FreeUser" (ByVal pUser As Long)
Declare Sub FreeUsers Lib "icqmapi.dll" Alias "ICQAPIUtil_FreeUsers" (ByVal iCount As Long, ByRef ppUsers As BSICQAPI_User)


Public Usr As BSICQAPI_User, ArrayUsrs() As BSICQAPI_User
Public Grp As BSICQAPI_Group, ArrayGrps() As BSICQAPI_Group
Public Rtn As Boolean, Vrsn As Long, MyStr As String


Global Const ICQAPINOTIFY_ONLINELIST_CHANGE = 0
Global Const ICQAPINOTIFY_FILE_RECEIVED = 8
Global Const ICQAPINOTIFY_ONLINE_FULLUSERDATA_CHANGE = 1
Global Const ICQAPINOTIFY_APPBAR_STATE_CHANGE = 2
Global Const ICQAPINOTIFY_ONLINE_PLACEMENT_CHANGE = 3
Global Const ICQAPINOTIFY_OWNER_CHANGE = 4
Global Const ICQAPINOTIFY_OWNER_FULLUSERDATA_CHANGE = 5
Global Const ICQAPINOTIFY_ONLINELIST_HANDLE_CHANGE = 6
Global Const ICQAPINOTIFY_LAST = 80
Global Const ICQAPINOTIFY_ONLINELISTCHANGE_ONOFF = 1
Global Const ICQAPINOTIFY_ONLINELISTCHANGE_FLOAT = 2
Global Const ICQAPINOTIFY_ONLINELISTCHANGE_POS = 3

Declare Function SetUserNotify Lib "icqmapi.dll" Alias _
"ICQAPIUtil_SetUserNotificationFunc" (ByVal uNotificationCode As Long, _
                                      ByVal pUserFunc As Any) As Long
                                      
                                      
Declare Function RegisterNotify Lib "icqmapi.dll" Alias _
     "ICQAPICall_RegisterNotify" (ByVal iVersion As Long, _
                                  ByVal iCount As Long, _
                                  ByVal piEvents As String) As Boolean
                                  
Declare Function UnRegisterNotify Lib "icqmapi.dll" Alias _
     "ICQAPICall_UnRegisterNotify" () As Boolean

Public Function GetDottedIP(LongIP As Long) As String

'Copyright (C) 2000 DIGITAL VAMPIRE
'DV@KNAC.COM - ICQ 14996057

'www.ic-crypt.org.uk

Dim Octet As Variant
Octet = Array(Right(Hex(LongIP), 2), Mid(Hex(LongIP), 5, 2), Mid(Hex(LongIP), 3, 2), Left(Hex(LongIP), 2))
GetDottedIP = Trim(Str(Val("&H" + Octet(0)))) + "." + _
                Trim(Str(Val("&H" + Octet(1)))) + "." + _
                Trim(Str(Val("&H" + Octet(2)))) + "." + _
                Trim(Str(Val("&H" + Octet(3))))
End Function

Sub GetUserList()

'Copyright (C) 2000 DIGITAL VAMPIRE
'DV@KNAC.COM - ICQ 14996057

'www.ic-crypt.org.uk

Dim N As Long, i As Long
Dim uinArray() As Long, usrArray() As BSICQAPI_User
Dim pUsers() As Long, ppUsers As Long


frmMain.List1.Clear
frmMain.List2.Clear
frmMain.List3.Clear

On Error Resume Next

Rtn = False
Rtn = GetOnlineListDetails(N, ppUsers)               'Get pointer to the array
                                                     'of user structures
ReDim pUsers(1 To N), uinArray(1 To N)
ReDim usrArray(1 To N)


Call CopyMemory(pUsers(1), ByVal ppUsers, 4 * N)     'Get all pointers
                                                     'to the users structures
For i = 1 To N
    Call CopyMemory(uinArray(i), ByVal pUsers(i), 4) 'Get structures one By one
     
    
    usrArray(i).m_iUIN = uinArray(i)                    'Initialize user structure
    
    FreeUser (pUsers(i))
    
    Call GetFullUserData(usrArray(i), iCQversion)          'Get user's details
    
    frmMain.List1.AddItem usrArray(i).m_szNickname
    frmMain.List2.AddItem usrArray(i).m_iUIN
    frmMain.List3.AddItem GetDottedIP(usrArray(i).m_iIP)
    
    
Next i


End Sub




Public Function PointerToString(p As Long) As String
'Copyright (C) 2000 DIGITAL VAMPIRE
'DV@KNAC.COM - ICQ 14996057
  
'www.ic-crypt.org.uk


  Dim c&
  c = lstrlen(p)
  Debug.Print c & "<- lstrlen"
  PointerToString = String$(c, 0)
  RtlMoveMemory ByVal PointerToString, ByVal p, c
  PointerToString = TrimNull(PointerToString)
End Function

Public Function TrimNull(s As String) As String

'Copyright (C) 2000 DIGITAL VAMPIRE
'DV@KNAC.COM - ICQ 14996057


'www.ic-crypt.org.uk
  
  Dim iWhere%
  iWhere = InStr(1, s, Chr(0))
  
  If iWhere > 0 Then
    TrimNull = Left$(s, iWhere - 1)
  Else
    TrimNull = s
    Debug.Print s & "<- no null present"
  End If
End Function



                                  


Public Sub ICQAPINotify_AppBarStateChange(ByVal iDockingState As Long)

' Code: ICQAPINOTIFY_APPBAR_STATE_CHANGE
' In (Arguments): iDockingState
' iDockingState is the new docking state
' 0 - floating
' 1 - docked right
' 2 - docked left
' 3 - docked top
' 4 - docked bottom
' Description: Sent if the contact list docking status has changed.
'Rtn = GetDockingState(iDockingState)

Call SendNotify
Select Case iDockingState
Case 0
frmMain.Caption = iCQowner & " - [ ICQ Control Center ] - ICQ Docked State = Floating"
Case 1
frmMain.Caption = iCQowner & " - [ ICQ Control Center ] - ICQ Docked State = Docked Right"
Case 2
frmMain.Caption = iCQowner & " - [ ICQ Control Center ] - ICQ Docked State = Docked Left"
Case 3
frmMain.Caption = iCQowner & " - [ ICQ Control Center ] - ICQ Docked State = Docked Top"
Case 4
frmMain.Caption = iCQowner & " - [ ICQ Control Center ] - ICQ Docked State = Docked Bottom"



End Select

If iDockingState = 0 Then

End If


End Sub

Public Sub ICQAPINotify_FileReceived(ByVal pszFileNames As Long)
'Code: ICQAPINOTIFY_FILE_RECEIVED
'In (Arguments): pszFileNames
'pszFileName - the name of the received file.
'Description: Sent when a file transfer event ended successfully.
'             When multiple files were sent, the notification will
'             be sent after each received file.

  
'Dim filename As String * 128

    

SendNotify
X = PointerToString(pszFileNames)
frmMain.Text1.Text = X




End Sub
Public Sub ICQAPINotify_FullUserDataChange(ByVal iUIN As Long)

'Code: ICQAPINOTIFY_ONLINE_FULLUSERDATA_CHANGE
'In (Arguments): iUIN
'iUIN is the ICQ# of the user
'Description: Sent when user’s details where updated
'             (e.g. when the owner pressed update,
'              in the info dialog).

End Sub
Public Sub ICQAPINotify_OnlineListChange(ByVal iType As Long)
'Code: ICQAPINOTIFY_ONLINELIST_CHANGE
'In (Arguments): iType
'iType is the type of change:
'1 - user gone on/off
'2 - float window on/off
'3 - user changed position in the list

'Description: Sent when the online users list changes.  A change
'             can occur due to the following reasons:
'             1. A user went online / offline.
'             2. The owner dragged a user to the desktop or back (floating).
'             3. A user moved up / down in the list (e.g. a user moved to
'                the top of the list because of an incoming event).

'This notification will be sent if you have asked for ICQAPINOTIFY_ONLINELIST_CHANGE

'NOTE on : Case 2 float window on/off exclusion
'
'to restrict the over-use and visible nasty-ness
'aka the listboxes "refreshing" we want to update
'the list on every part EXCEPT when user is
'floated on or off.. hehehe that *could* be done
'I just done see the point since Floating On or Off
'does not effect the physical online list in any way
'except that it produces an extra floating window :)
'on some configurations / ICQ versions
'this *MAY* well produce different "physical"
'effects however, on mine it dose as aforementioned
'thus for my box making implementing this feature
'for these functions pointless.

'bynote : you may well want to implement this
'for other reasons, put simply - depends on YOUR
'purpose for implemtation :)

Select Case iType
Case 1
    Call GetUserList
Case 2
    Call GetUserList
    
    ' ohhh bugger it's fun when you are bored :)
    ' and om .. testing purposes of course hehehe :)
        
Case 3
    Call GetUserList
    
End Select

'NOTE on : case 3 - user changed position in the
'list while to save system resources, this *could*
'be happily excluded, however depends on your reasons
'for implenting this procedure and also depends
'on your thirst for "niceties" hehehe :)

Call SendNotify

'Over-Use but incase of lag or weirdness !
'theoretically not needed, but under tests
'it was a wise move to include it, due to
'sometimes system weirdness notifcation lag
'and poor icq specification / code.




End Sub
Public Sub ICQAPINotify_OnlineListHandleChange(ByVal hWindow As Long)

'Code: ICQAPINOTIFY_ONLINE_LISTHANDLE_CHANGE
'In (Arguments): hWindow
'hWindow - the handle of the current contact list listbox
'Description: Sent when the user switches between available
'             contact lists (i.e. “Online” tab or “All” tab).

End Sub
Public Sub ICQAPINotify_OnlinePlacementChange()

'Code: ICQAPINOTIFY_ONLINE_PLACEMENT_CHANGE
'In (Arguments): None
'Description: Sent when the online tab is added/removed above
'             the contact list (due to the user changing the preferences)

End Sub
Public Sub ICQAPINotify_OwnerChange(ByVal iUIN As Long)

'Code: ICQAPINOTIFY_OWNER_CHANGE
'In (Arguments): iUIN
'iUIN - the UIN of the new owner
'Description: Sent when the current owner has changed.

End Sub
Public Sub ICQAPINotify_OwnerFullDataChange()
'Code: ICQAPINOTIFY_OWNERFULLDATA_CHANGE

'Out (Arguments): None
'Description: Sent when the owner updates his own details.

End Sub

Sub SendNotify()

Dim sEvents As String

Rtn = False                                           'Intialize booean variable

Call SetUserNotify(ICQAPINOTIFY_ONLINELIST_HANDLE_CHANGE, AddressOf ICQAPINotify_OnlineListHandleChange)
Call SetUserNotify(ICQAPINOTIFY_ONLINELIST_CHANGE, AddressOf ICQAPINotify_OnlineListChange)
Call SetUserNotify(ICQAPINOTIFY_APPBAR_STATE_CHANGE, AddressOf ICQAPINotify_AppBarStateChange)
Call SetUserNotify(ICQAPINOTIFY_ONLINE_PLACEMENT_CHANGE, AddressOf ICQAPINotify_OnlinePlacementChange)
Call SetUserNotify(ICQAPINOTIFY_FILE_RECEIVED, AddressOf ICQAPINotify_FileReceived)
Call SetUserNotify(ICQAPINOTIFY_OWNER_CHANGE, AddressOf ICQAPINotify_OwnerChange)

sEvents = Chr(ICQAPINOTIFY_ONLINELIST_CHANGE) & Chr(ICQAPINOTIFY_ONLINELIST_CHANGE) & _
          Chr(ICQAPINOTIFY_FILE_RECEIVED) & Chr(ICQAPINOTIFY_APPBAR_STATE_CHANGE) & _
          Chr(ICQAPINOTIFY_ONLINE_PLACEMENT_CHANGE) & Chr(ICQAPINOTIFY_OWNER_CHANGE)
        

Rtn = RegisterNotify(iCQversion, 6, sEvents)


If Rtn = False Then
    MsgBox ("Error: Could not register callbacks.")
End If

End Sub
Function GetOwnerUin()


Dim pOwner As BSICQAPI_User, ppOwner As BSICQAPI_User
Dim uin As Long, ownrSize As Long
'On Error Resume Next
Rtn = False                                           'Intialize booean variable
                                                      'if ICQ client is running online
ownrSize = Len(ppOwner)

Rtn = GetFullOwnerData(ppOwner, iCQversion)                 'Get the pointer to the pointer
                                                      'of the owner's structure.
Call CopyMemory(pOwner, ByVal ppOwner.m_iUIN, ownrSize)  'Get pointer to the
                                                      'owner's structure.
'Call DisplayUser(1, 1, pOwner)                        'Show owner's details

GetOwnerUin = pOwner.m_iUIN


Call FreeUsers(1, ppOwner)                            'Free owner structure pointer
End Function
Function GetOwnerEmail()

Dim pOwner As BSICQAPI_User, ppOwner As BSICQAPI_User
Dim uin As Long, ownrSize As Long
'On Error Resume Next
Rtn = False                                           'Intialize booean variable
                                                      'if ICQ client is running online
ownrSize = Len(ppOwner)

Rtn = GetFullOwnerData(ppOwner, iCQversion)                 'Get the pointer to the pointer
                                                      'of the owner's structure.
Call CopyMemory(pOwner, ByVal ppOwner.m_iUIN, ownrSize)  'Get pointer to the
                                                      'owner's structure.
'Call DisplayUser(1, 1, pOwner)                        'Show owner's details
If InStr(1, pOwner.m_szEmail, 1) Then
GetOwnerEmail = Trim(pOwner.m_szEmail)
Else
GetOwnerEmail = "no@one.com"
End If



Call FreeUsers(1, ppOwner)                            'Free owner structure pointer
End Function

Function GetOwnerIP()

Dim pOwner As BSICQAPI_User, ppOwner As BSICQAPI_User
Dim uin As Long, ownrSize As Long
'On Error Resume Next
Rtn = False                                           'Intialize booean variable
                                                      'if ICQ client is running online
ownrSize = Len(ppOwner)

Rtn = GetFullOwnerData(ppOwner, iCQversion)                 'Get the pointer to the pointer
                                                      'of the owner's structure.
Call CopyMemory(pOwner, ByVal ppOwner.m_iUIN, ownrSize)  'Get pointer to the
                                                      'owner's structure.
'Call DisplayUser(1, 1, pOwner)                        'Show owner's details

GetOwnerIP = GetDottedIP(pOwner.m_iIP)

Call FreeUsers(1, ppOwner)                            'Free owner structure pointer
End Function


Function GetOwnerNickName()

Dim pOwner As BSICQAPI_User, ppOwner As BSICQAPI_User
Dim uin As Long, ownrSize As Long
'On Error Resume Next
Rtn = False                                           'Intialize booean variable
                                                      'if ICQ client is running online
ownrSize = Len(ppOwner)

Rtn = GetFullOwnerData(ppOwner, iCQversion)                 'Get the pointer to the pointer
                                                      'of the owner's structure.
Call CopyMemory(pOwner, ByVal ppOwner.m_iUIN, ownrSize)  'Get pointer to the
                                                      'owner's structure.

GetOwnerNickName = pOwner.m_szNickname


Call FreeUsers(1, ppOwner)                            'Free owner structure pointer
End Function


Function GetOwnerStatus()

Dim pOwner As BSICQAPI_User, ppOwner As BSICQAPI_User
Dim uin As Long, ownrSize As Long
'On Error Resume Next
Rtn = False                                           'Intialize booean variable
                                                      'if ICQ client is running online
ownrSize = Len(ppOwner)

Rtn = GetFullOwnerData(ppOwner, iCQversion)                 'Get the pointer to the pointer
                                                      'of the owner's structure.
Call CopyMemory(pOwner, ByVal ppOwner.m_iUIN, ownrSize)  'Get pointer to the
                                                      'owner's structure.
'Call DisplayUser(1, 1, pOwner)                        'Show owner's details


Call FreeUsers(1, ppOwner)                            'Free owner structure pointer

GetOwnerStatus = pOwner.m_iStateFlags
End Function






