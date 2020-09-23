Attribute VB_Name = "SubMain"
' =================================================================
' Password Guard source code
' Version 1.5
' Copyright (C) 2000-2001 Khaery Rida
' =================================================================

' Thanx for using Password Guard!
' Please log on http://www.geocities.com/lio889 for more great VB programs!
' Comments or Questions? Please do NOT hesitate at emailing me:
' lio_889@ziplip.com

' Declare Windows' API functions
Public Declare Sub SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetCursorPos Lib "User32" (ByRef lpPoint As Point) As Long
Public Declare Function WindowFromPoint Lib "User32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Global Constants
Global Const MainTitle = "Password Guard"
Global Const MasterKey = "PGmECk"
Global Const Key1 = "PswrdGrd"
Global Const Key2 = "CherAlog"
Global Const FileTitle = "Password Gaurd Data File"
Global Const conKey = ""
Global Const sDivide = "ƒ"

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40

' Color Constants
Global Const LightOn = &HC0&
Global Const LightOff = &H800000
Global Const Active = &H80000005
Global Const NoActive = &H8000000F
Global Const Navy = &H800000
Global Const SRCCOPY = &HCC0020

' Message Box Constants
Global Const MB_YESNO = 4
Global Const MB_ICONQUESTION = 32
Global Const MB_DEFBUTTON1 = &H0&
Global Const MB_DEFBUTTON2 = 256
Global Const IDYES = 6

' Maximum size for Log File (in bytes)
Global Const MaxLogFileLen = 6000

' User-defined types
Public Type DataRecord
    Description As String
    Server As String
    UserName As String
    Password As String
    Notes As String
End Type

Public Type PasswordEncodingFlags
    encodedPassword As String
    arFlag(1 To 16) As Long
End Type
    
Public Type Point
    X As Long
    Y As Long
End Type

' Public Variables
' ============

' Stores selected record's information.
Public tDescription As String
Public tServer As String
Public tUserName As String
Public tPassword As String
Public tNotes As String

' Stores the serial number and registry's key
' for the selected record.
Public currentRecord As Long
Public currentKey As String

' Determines whether the selected record's
' information has been saved (TRUE),
' or not (FALSE) after being modified.
Public Saved As Boolean

' Used at loading time. Stores the Name of the Registry
' Section currently being scanned.
' (HKEY_CURRENT_USER\VB and VBA Program Settings\Password Guard\{currentIndex})
Public currentIndex As Long

' Stores the number of records.
Public ItemCount As Long

Public MainList As ListItem
Public Keyword As String

' Stores temporary information
Public currentCounter As Long
Public currentElement As Long
Public tmpAsc As Long
Public regString As String
Public tmpString As String
Public tmpString2 As String
Public tmpString3 As String

' Stores information for the currently LOGGED IN user.
Public Index As String
Public UserKeyword As String
Public UserIDKeyword As String
Public MasterPasswordKeyword As String

Public UserID As String
Public MasterPassword As String
Public UserRegSection As String
Public UserIndex As String
Public ItemIndex As String
Public HintAnswer As String
Public FMasterPassword As PasswordEncodingFlags

' Stores information of the Access Logging feature.
Public TrackUser As Boolean
Public TrackAll As Boolean
Public EncodeFile As Boolean
Public LogFile As String

' for Message Boxes.
Public DgDef, Msg, Response, Title

' Stores funny words, these funny words will be
' used to generate a new Master Password if
' the user forgot his/her ol Master Password
Public LogWord(1 To 10) As String
Public SearchResult() As String

Public Sub Main()

' Load forms:
    Load frmLogIn
    Load frmMain
    Load frmUserID
    Load frmAbout
    
    Dim cIndex As String
    Dim cUserIDIndex As String
    Dim cOptionsIndex As String
    Dim cOptions As String
    
    frmLogIn.cmdSubmit.Enabled = False
    regString = GetSetting(MainTitle, "Settings", "Index", "/NEWRUN/")
    
    If regString = "/NEWRUN/" Then
        Index = "/NEWRUN/"
        UserCount = 0
        frmLogIn.Show
        Exit Sub
    End If
    
    Index = decrypt(regString, Key1 & Key2)
    
    For currentIndex = 1 To Len(Index) Step 6
        UserCount = UserCount + 1
        UserRegSection = Mid$(Index, currentIndex, 6)
        regString = GetSetting(MainTitle, UserRegSection, "Index")
        cIndex = decrypt(regString, Key2 & UserRegSection)
        cOptionsIndex = Mid$(cIndex, 17, 8)
        regString = GetSetting(MainTitle, UserRegSection, cOptionsIndex)
        cOptions = decrypt(regString, Key1 & Mid$(cOptionsIndex, 3, 1) & Mid$(cOptionsIndex, 5, 1))
        
        If Left$(cOptions, 1) = "1" Then
            cUserIndex = Left$(cIndex, 8)
            regString = GetSetting(MainTitle, UserRegSection, cUserIndex)
            frmLogIn.lstUserID.AddItem decrypt(regString, Left$(UserRegSection, 2) & Right$(cUserIndex, 2) & Key1)
        End If
    Next
    
    ' Log file statements
    LogWord(1) = "User logged in, &D at &T."
    LogWord(2) = "User logged out at &T."
    LogWord(3) = "User modified Master Password at &T."
    LogWord(4) = "User modified Hint Question at &T."
    LogWord(5) = "User modified Hint Answer at &T."
    LogWord(6) = "User added a new record at &T."
    LogWord(7) = "User modified the record &A at &T."
    LogWord(8) = "User deleted the record &A at &T."
    LogWord(9) = "User imported data from the file &A at &T."
    LogWord(10) = "User exporeted data to the file &A at &T."
    
    frmLogIn.lstUserID.Text = ""
    frmLogIn.Show
    Exit Sub
    
End Sub
Public Function FileExists(Path$) As Integer

' This function is used to ensure that a file is openable.
  
    X = FreeFile

    On Error Resume Next
    Open Path$ For Input As X
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close X

End Function
Public Function IsValidUserID(UserID As String) As String
    
    Dim vUserIndex As String
    Dim vUserRegSection As String
    
    For currentIndex = 1 To Len(Index) Step 6
        vUserRegSection = Mid$(Index, currentIndex, 6)
        regString = GetSetting(MainTitle, vUserRegSection, "Index")
        vUserIndex = decrypt(regString, Key2 & vUserRegSection)
        regString = GetSetting(MainTitle, vUserRegSection, Left$(vUserIndex, 8))
        UserIDKeyword = Left$(vUserRegSection, 2) & Right$(Left$(vUserIndex, 8), 2) & Key1
        
        If decrypt(regString, UserIDKeyword) = UserID Then
            IsValidUserID = vUserRegSection & vUserIndex
            Exit Function
        End If
    Next
    
    IsValidUserID = ""
    
End Function

Public Function IsValidMasterPassword(regSection As String, regKey As String, UserID As String, MasterPassword As String) As String

    regString = GetSetting(MainTitle, regSection, regKey)
    MasterPasswordKeyword = Mid$(regSection, 3, 1) & Left$(UserID, 1) & Left$(MasterPassword, Len(MasterPassword) - 5) & Right$(MasterPassword, 1) & Right$(regKey, 1) & Right$(UserID, 1)
    tmpString = decrypt(regString, MasterPasswordKeyword)
    
    If Trim(MasterPassword) = tmpString Then
        IsValidMasterPassword = Left$(UserID, 1) & Mid$(regSection, 5, 2) & MasterPassword
    Else
        IsValidMasterPassword = ""
    End If
    
End Function

Public Sub LogIn()

Screen.MousePointer = 11
        
    ' Fill in frmUserID
    frmUserID.txtUserID.Text = UserID
    frmUserID.txtUserID.Tag = Left$(UserIndex, 8)
    frmUserID.txtUserID.Locked = True
    
    frmUserID.txtMasterPassword1.Text = MasterPassword
    frmUserID.txtMasterPassword2.Text = MasterPassword
    frmUserID.txtMasterPassword1.Tag = Mid$(UserIndex, 9, 8)
    
    frmUserID.frameLog.Tag = Mid$(UserIndex, 17, 8)
    frmUserID.txtQuestion.Tag = Mid$(UserIndex, 25, 8)
    frmUserID.txtAnswer.Tag = Mid$(UserIndex, 33, 8)
    frmUserID.txtMasterPassword2.Tag = Mid$(UserIndex, 41, 8)
    HintAnswer = decrypt(GetSetting(MainTitle, UserRegSection, frmUserID.txtAnswer.Tag), UserKeyword & Mid$(frmUserID.txtAnswer.Tag, 3, 1) & Mid$(frmUserID.txtAnswer.Tag, 5, 1))
    
    frmUserID.txtUserID.Text = decrypt(GetSetting(MainTitle, UserRegSection, frmUserID.txtUserID.Tag), UserIDKeyword)
    frmUserID.txtMasterPassword1.Text = decrypt(GetSetting(MainTitle, UserRegSection, frmUserID.txtMasterPassword1.Tag), MasterPasswordKeyword)
    frmUserID.txtMasterPassword2.Text = frmUserID.txtMasterPassword1.Text
    frmUserID.txtQuestion.Text = decrypt(GetSetting(MainTitle, UserRegSection, frmUserID.txtQuestion.Tag), Key1 & Mid$(frmUserID.txtQuestion.Tag, 3, 1) & Mid$(frmUserID.txtQuestion.Tag, 5, 1))
    frmUserID.txtAnswer.Text = HintAnswer
    regString = decrypt(GetSetting(MainTitle, UserRegSection, frmUserID.frameLog.Tag), Key1 & Mid$(frmUserID.frameLog.Tag, 3, 1) & Mid$(frmUserID.frameLog.Tag, 5, 1))
    frmUserID.chkDisplay.Value = Left$(regString, 1)
    frmUserID.chkPassword.Value = Mid$(regString, 2, 1)
    frmUserID.chkRemove.Value = Mid$(regString, 3, 1)
    frmUserID.chkLog.Value = Mid$(regString, 4, 1)
    frmUserID.chkLogAll.Value = Mid$(regString, 5, 1)
    frmUserID.chkEncrypt.Value = Mid$(regString, 6, 1)
    
    frmUserID.chkLog_Click
    
    If Len(regString) > 6 Then frmUserID.txtLogFile.Text = Mid$(regString, 7)
    
    ' Process Log File
    If frmUserID.chkLog.Value = 1 Then TrackUser = True Else TrackUser = False
    If frmUserID.chkLogAll.Value = 1 Then TrackAll = True Else TrackAll = False
    If frmUserID.chkEncrypt.Value = 1 Then EncodeFile = True Else EncodeFile = False
    LogFile = frmUserID.txtLogFile.Text
    
    If LogFile = "" Then TrackUser = False: TrackAll = False: EncodeFile = False
    If (Not FileExists(LogFile + "")) And (Len(Dir(LogFile)) > 0) Then TrackUser = False: TrackAll = False: EncodeFile = False
    
    If FileExists(LogFile + "") And TrackUser = True Then
        Dim oldStoredData As String
                
        oldStoredData = ""
        Open LogFile For Input As #2
            If LOF(2) = 0 Or LOF(2) > MaxLogFileLen Then
                Close #2
                Kill LogFile
                GoTo ErrHandlerZero
            End If
        Do Until EOF(2)
            Line Input #2, tmpString
            oldStoredData = oldStoredData & tmpString & vbCrLf
        Loop
        Close #2
    
ErrHandlerZero:
        Open LogFile For Output As #2
        If Not oldStoredData = "" Then Print #2, oldStoredData
        If Not oldStoredData = "" Then Print #2, ""
   End If
NoUserIDFormSet:
   Call LogAction(LogWord(1))
   ItemCount = 0
   regString = GetSetting(MainTitle, UserRegSection, "Item", "")
   If Not regString = "" Then ItemIndex = decrypt(regString, UserKeyword) Else ItemIndex = ""
   
   frmMain.lstItem.ListItems.Clear
   frmMain.txtDescription.Text = ""
   frmMain.txtServer.Text = ""
   frmMain.txtUserName.Text = ""
   frmMain.txtPassword.Text = ""
   frmMain.txtNotes.Text = ""
   
   If Len(ItemIndex) = 0 Then
        frmMain.txtDescription.Enabled = False
        frmMain.txtServer.Enabled = False
        frmMain.txtUserName.Enabled = False
        frmMain.txtPassword.Enabled = False
        frmMain.txtNotes.Enabled = False
        frmMain.mnuRemoveRecord.Enabled = False
        frmMain.mnuSearch.Enabled = False
        frmMain.imgRemove.Visible = False
        frmMain.lblRemove.Visible = False
        frmMain.imgSearch.Visible = False
        frmMain.lblSearch.Visible = False
        
        GoTo NoItems
    End If
    
    Dim curKey As Long
    Dim KeyString As String
    Dim tmp1stRecord As DataRecord

    ' Add all items to the ListView control
    For curKey = 1 To Len(ItemIndex) Step 8
        KeyString = Mid$(ItemIndex, curKey, 8)
        regString = GetSetting(MainTitle, UserRegSection, KeyString)
        tmpString = decrypt(regString, UserKeyword & Mid$(KeyString, 3, 1) & Mid$(KeyString, 5, 1))
        currentChr = 0
        tmpString2 = ""
        Do Until tmpString2 = sDivide
            currentChr = currentChr + 1
            tmpString2 = Mid$(tmpString, currentChr, 1)
        Loop
        Set MainList = frmMain.lstItem.ListItems.Add(, KeyString, Left$(tmpString, currentChr - 1), 1, 1)
        ItemCount = ItemCount + 1
    Next
    
    tmp1stRecord = ReadRecord(Left$(ItemIndex, 8))
    tDescription = tmp1stRecord.Description
    tServer = tmp1stRecord.Server
    tUserName = tmp1stRecord.UserName
    tPassword = tmp1stRecord.Password
    tNotes = tmp1stRecord.Notes
    
    frmMain.txtDescription.Text = tDescription
    frmMain.txtServer.Text = tServer
    frmMain.txtUserName.Text = tUserName
    frmMain.txtPassword.Text = tPassword
    frmMain.txtNotes.Text = tNotes
            
    frmMain.lstItem.ListItems(1).Selected = True
    currentKey = Left$(ItemIndex, 8)
    currentRecord = 1
    Saved = True
    
NoItems:
    frmMain.Caption = MainTitle & " - [Welcome " & UserID & "]"
    Screen.MousePointer = 0
    frmMain.Show
    
End Sub


Public Function RandomPinString(PinNum As Integer) As String
    
    Dim tOffset As Integer
        
GenerateRndPinString:
        Randomize
        For currentPin = 1 To PinNum
            tOffset = (Rnd * 10000 Mod 255) + 1
            RandomPinString = RandomPinString & Format$(Hex$(tOffset), "@@")
        Next
        
        ' The Format$ function is used to make sure that always 2 bytes are returned.
        ' For example, instead of returning "B", the Format$ function, returns "B "
        ' in this way, the resulting RandomPinString will always consist of (PinNum * 2) characters.
        
        If IsNumeric(RandomPinString) Then
            RandomPinString = ""
            GoTo GenerateRndPinString
        End If
        ' Since this RandomPinString will be used as a ListItem key, we need to ensure that the
        ' result RadnomPinString is not a whole numeric value. For example, a RandomPinString
        ' may take te value 24982751 which is an invalid ListItem's key.
       
End Function
Public Function ObtainFilePassword(uAction As String, File As String) As String
    
    Load frmKeyword
    frmKeyword.txtKeyword.Text = ""
    frmKeyword.cmdOK.Enabled = False
    
    If uAction = "Export" Then
        frmKeyword.lblAction.Caption = "lock"
        frmKeyword.Caption = "Export Records to File"
        frmKeyword.img.Picture = frmMain.imgExport.Picture
    ElseIf uAction = "Import" Then
        frmKeyword.lblAction.Caption = "unlock"
        frmKeyword.Caption = "Import Records from File"
        frmKeyword.img.Picture = frmMain.imgImport.Picture
    End If
    
    frmKeyword.lblFile.Caption = LCase$(File)
    frmKeyword.img.Height = frmKeyword.img.Height - 20
    frmKeyword.img.Stretch = True
    frmKeyword.Show 1
    ObtainFilePassword = Trim(frmKeyword.txtKeyword.Text)
    Unload frmKeyword
    
End Function
Public Sub UpdateProgress(PositionSetBack As Long, StringLength As Long)
    
' This function is used to update the progress bar while an operation is in progress.
    
    Static position
    Dim txt As String, r As Long, estTotal As Long
    estTotal = frmMain.picProgress.Tag
    
    If PositionSetBack = 1 Then position = 0

    position = position + CSng((StringLength / estTotal) * 100)
    If position > 100 Then
        position = 100
    End If
    
    txt$ = Format$(CLng(position)) + "%"
    
    frmMain.picProgress.Line (0, 0)-((position * (frmMain.picProgress.ScaleWidth / 100)), frmMain.picProgress.ScaleHeight), Navy, BF
    frmMain.picProgress.CurrentX = (frmMain.picProgress.ScaleWidth - frmMain.picProgress.TextWidth(txt$)) \ 2
    frmMain.picProgress.CurrentY = (frmMain.picProgress.ScaleHeight - frmMain.picProgress.TextHeight(txt$)) \ 2
    r = BitBlt(frmMain.picProgress.hDC, 0, 0, frmMain.picProgress.ScaleWidth, frmMain.picProgress.ScaleHeight, frmMain.picProgress.hDC, 0, 0, SRCCOPY)

End Sub
Public Function ReadRecord(regKey As String) As DataRecord
    
    Dim memberCount As Long
    Dim memberLen As Long
    Dim LastPos As Long

    regString = GetSetting(MainTitle, UserRegSection, regKey)
    tmpString = decrypt(regString, UserKeyword & Mid$(regKey, 3, 1) & Mid$(regKey, 5, 1))
    memberCount = 0
    memberLen = 0
    LastPos = 0
    
    For currentChr = 1 To Len(tmpString)
        tmpString2 = Mid$(tmpString, currentChr, 1)
        memberLen = memberLen + 1
        
        If tmpString2 = sDivide Then
            memberCount = memberCount + 1
            
            Select Case memberCount
                Case 1
                    ReadRecord.Description = Left$(tmpString, currentChr - 1)
                Case 2
                    ReadRecord.Server = Mid$(tmpString, LastPos, memberLen - 1)
                Case 3
                    ReadRecord.UserName = Mid$(tmpString, LastPos, memberLen - 1)
                Case 4
                    ReadRecord.Password = Mid$(tmpString, LastPos, memberLen - 1)
                Case 5
                    ReadRecord.Notes = Mid$(tmpString, LastPos, memberLen - 1)
            End Select
            
            memberLen = 0
            LastPos = currentChr + 1
        End If
    Next
    
End Function
Public Sub Search(Mode As Long, Description As String, Server As String, UserName As String, Password As String, Notes As String)
        
    ReDim SearchResult(0)
    frmSearchResults.lstSearch.Clear
    
    Dim searchKey As String
    Dim searchOutput As DataRecord
    Dim currentSearch As Long
    Dim RecNum As Long
    Dim MemberMatch As Long
    
    RecNum = -1
    For currentSearch = 1 To Len(ItemIndex) Step 8
        searchKey = Mid$(ItemIndex, currentSearch, 8)
        searchOutput = ReadRecord(searchKey)
        
        If Mode = 1 Then    ' Match any
            If Len(Description) > 0 Then
                If Left$(searchOutput.Description, Len(Description)) = Description Or Right$(searchOutput.Description, Len(Description)) = Description Then GoTo RecordMatch1
            End If
            If Len(Server) > 0 Then
                If Left$(searchOutput.Server, Len(Server)) = Server Or Right$(searchOutput.Server, Len(Server)) = Server Then GoTo RecordMatch1
            End If
            If Len(UserName) > 0 Then
                If Left$(searchOutput.UserName, Len(UserName)) = UserName Or Right$(searchOutput.UserName, Len(UserName)) = UserName Then GoTo RecordMatch1
            End If
            If Len(Password) > 0 Then
                If Left$(searchOutput.Password, Len(Password)) = Password Or Right$(searchOutput.Password, Len(Password)) = Password Then GoTo RecordMatch1
            End If
            If Len(Notes) > 0 Then
                If Left$(searchOutput.Notes, Len(Notes)) = Password Or Right$(searchOutput.Notes, Len(Notes)) = Notes Then GoTo RecordMatch1
            End If
            GoTo NextSearch
       
RecordMatch1:
            RecNum = RecNum + 1
            SearchResult(RecNum) = searchKey
            ' Resize array
            ReDim Preserve SearchResult(UBound(SearchResult) + 1)
            frmSearchResults.lstSearch.AddItem searchOutput.Description
        End If
        
        If Mode = 2 Then    ' Match all
            
            MemberMatch = 0
            If Len(Description) > 0 Then
                If Left$(searchOutput.Description, Len(Description)) = Description Then MemberMatch = MemberMatch + 1
                If Right$(searchOutput.Description, Len(Description)) = Description Then MemberMatch = MemberMatch + 1
                If MemberMatch = 0 Then GoTo NextSearch
            End If
            
             MemberMatch = 0
             If Len(Server) > 0 Then
                If Left$(searchOutput.Server, Len(Server)) = Server Then MemberMatch = MemberMatch + 1
                If Right$(searchOutput.Server, Len(Server)) = Server Then MemberMatch = MemberMatch + 1
                If MemberMatch = 0 Then GoTo NextSearch
            End If
            
            MemberMatch = 0
            If Len(UserName) > 0 Then
                If Left$(searchOutput.UserName, Len(UserName)) = UserName Then MemberMatch = MemberMatch + 1
                If Right$(searchOutput.UserName, Len(UserName)) = UserName Then MemberMatch = MemberMatch + 1
                If MemberMatch = 0 Then GoTo NextSearch
            End If
            
            MemberMatch = 0
            If Len(Password) > 0 Then
                If Left$(searchOutput.Password, Len(Password)) = Password Then MemberMatch = MemberMatch + 1
                If Right$(searchOutput.Password, Len(Password)) = Password Then MemberMatch = MemberMatch + 1
                If MemberMatch = 0 Then GoTo NextSearch
            End If

            MemberMatch = 0
            If Len(Notes) > 0 Then
                If Left$(searchOutput.Notes, Len(Notes)) = Notes Then MemberMatch = MemberMatch + 1
                If Right$(searchOutput.Notes, Len(Notes)) = Notes Then MemberMatch = MemberMatch + 1
                If MemberMatch = 0 Then GoTo NextSearch
            End If
            
            ' Record Match
            RecNum = RecNum + 1
            SearchResult(RecNum) = searchKey
            ReDim Preserve SearchResult(UBound(SearchResult) + 1)
            frmSearchResults.lstSearch.AddItem searchOutput.Description
            GoTo NextSearch
            
        End If
NextSearch:
    Next
        
    frmSearchResults.chkOnTop.Value = 1
    frmSearchResults.chkOnTop_Click
    frmSearchResults.Show
    
End Sub

Public Sub LogAction(Data As String, Optional addInfo As String)
    
    Dim curDataChar As Long
    Dim destData As String
    Dim formatString As String
    
    For curDataChar = 1 To Len(Data)
        tmpString = Mid$(Data, curDataChar, 1)
        If tmpString = "&" Then
            curDataChar = curDataChar + 1
            formatString = Mid$(Data, curDataChar, 1)
            Select Case formatString
                Case "D"
                    destData = destData & GetDate
                Case "T"
                    destData = destData & GetTime
                Case "A"
                    destData = destData & addInfo
                End Select
        Else
            destData = destData & tmpString
        End If
    Next curDataChar
    If TrackUser = False Then Exit Sub
    If EncodeFile = True Then destData = crypt(destData, UserKeyword)
    On Error Resume Next
    Print #2, destData
    Exit Sub
    
End Sub

Public Function GetTime() As String
    GetTime = Format(Now, "hh:mm:ss")
End Function

Public Function GetDate() As String
    GetDate = Format(Now, "dd/mm/yyyy")
End Function

Public Function AscEncode(sourceInput As String) As String
    
    Dim ascResult As Long
        
    If Len(sourceInput) > 30 Then AscEncode = "": Exit Function
    For currentCounter = 1 To Len(sourceInput)
        If (currentCounter < Len(sourceInput)) And (currentCounter > 1) Then currentCounter = currentCounter + 1
        ascResult = 0
        ascResult = ascResult + Asc(Mid$(sourceInput, currentCounter, 1))
        If currentCounter < Len(sourceInput) Then ascResult = ascResult + Asc(Mid$(sourceInput, currentCounter + 1, 1))
        ascResult = ((ascResult * ascResult) Mod 7) + 100
        AscEncode = AscEncode & Chr$(ascResult)
    Next currentCounter
End Function
