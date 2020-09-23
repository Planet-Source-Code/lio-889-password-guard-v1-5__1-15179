VERSION 5.00
Begin VB.Form frmUserID 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User ID settings"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameDelete 
      Caption         =   "Delete User ID"
      Height          =   2055
      Left            =   5760
      TabIndex        =   19
      Top             =   2880
      Width           =   3615
      Begin VB.CommandButton cmdDeleteUserID 
         Caption         =   "&Delete User ID"
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Image imgDeleteUserID 
         Height          =   465
         Left            =   120
         Picture         =   "frmUserID.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   450
      End
      Begin VB.Label imgDeleteUserIDCap 
         Caption         =   "You may delete this User ID by clicking ""Delete User ID""."
         Height          =   495
         Left            =   720
         TabIndex        =   20
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame frameLog 
      Caption         =   "Access Loging"
      Height          =   2775
      Left            =   5760
      TabIndex        =   15
      Top             =   0
      Width           =   3615
      Begin VB.CheckBox chkEncrypt 
         Caption         =   "Encrypt Log File"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdViewLog 
         Caption         =   "&View log file"
         Height          =   375
         Left            =   1440
         TabIndex        =   23
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CheckBox chkLogAll 
         Caption         =   "Log all actions (e.g. adding records...)"
         Height          =   375
         Left            =   840
         TabIndex        =   22
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtLogFile 
         Height          =   285
         Left            =   600
         TabIndex        =   18
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CheckBox chkLog 
         Caption         =   "Log access to this User ID"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   360
         Width           =   2535
      End
      Begin VB.Image imgLoging 
         Height          =   480
         Left            =   120
         Picture         =   "frmUserID.frx":0C4A
         Stretch         =   -1  'True
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lblLogFileCap 
         Caption         =   "Log File"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Ca&ncel"
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Okay"
      Height          =   375
      Left            =   7800
      TabIndex        =   13
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame frameLogin 
      Caption         =   "Login Information"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.CheckBox chkRemove 
         Caption         =   "Always confirm record removal"
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   5040
         Width           =   3015
      End
      Begin VB.CheckBox chkPassword 
         Caption         =   "Do NOT reveal Password field"
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   4680
         Width           =   2895
      End
      Begin VB.TextBox txtAnswer 
         Height          =   285
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   12
         Top             =   3720
         Width           =   3495
      End
      Begin VB.TextBox txtQuestion 
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Top             =   3240
         Width           =   3495
      End
      Begin VB.CheckBox chkDisplay 
         Caption         =   "Display User ID in login ComboList"
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   4320
         Width           =   3375
      End
      Begin VB.TextBox txtMasterPassword2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtMasterPassword1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtUserID 
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label lblAnswerMax 
         AutoSize        =   -1  'True
         Caption         =   "(6-30 characters)"
         Height          =   195
         Left            =   3960
         TabIndex        =   29
         Top             =   4080
         Width           =   1500
      End
      Begin VB.Label lblOptionsCap 
         AutoSize        =   -1  'True
         Caption         =   "Options:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   4320
         Width           =   795
      End
      Begin VB.Label lblLogInCap 
         Caption         =   "Except the User ID, all the values below are Case Sensitive (6-12 characters for Master Password)."
         Height          =   375
         Left            =   840
         TabIndex        =   25
         Top             =   480
         Width           =   4575
      End
      Begin VB.Image imgLoginInformation 
         Height          =   480
         Left            =   240
         Picture         =   "frmUserID.frx":1C0C
         Top             =   480
         Width           =   435
      End
      Begin VB.Image imgQnA 
         Height          =   450
         Left            =   240
         Picture         =   "frmUserID.frx":27D6
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   570
      End
      Begin VB.Label lblAnswerCap 
         AutoSize        =   -1  'True
         Caption         =   "Your answer :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label lblQuestionCap 
         AutoSize        =   -1  'True
         Caption         =   "Hint Question:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   3240
         Width           =   1365
      End
      Begin VB.Label lblQnACap 
         Caption         =   "In case you forgot your Master Password, type the question you'd like to display and your answer."
         Height          =   495
         Left            =   960
         TabIndex        =   7
         Top             =   2640
         Width           =   4455
      End
      Begin VB.Label lblMasterPassword2Cap 
         Caption         =   "Confirm Master Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2000
         Width           =   1575
      End
      Begin VB.Label lblMasterPassword1Cap 
         Caption         =   "Master Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1480
         Width           =   975
      End
      Begin VB.Label lblUserIDCap 
         AutoSize        =   -1  'True
         Caption         =   "User ID:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmUserID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub chkLog_Click()
If chkLog.Value = 0 Then
    chkLogAll.Enabled = False
    chkEncrypt.Enabled = False
    txtLogFile.Enabled = False
    txtLogFile.BackColor = NoActive
    cmdViewLog.Enabled = False

ElseIf chkLog.Value = 1 Then
    chkLogAll.Enabled = True
    chkEncrypt.Enabled = True
    txtLogFile.Enabled = True
    txtLogFile.BackColor = Active
    cmdViewLog.Enabled = True

End If

End Sub

Private Sub cmdCancel_Click()
    
    Screen.MousePointer = 11
    If cmdOK.Caption = "&Create" Then
        txtUserID.Text = ""
    
    ElseIf cmdOK.Caption = "&Okay" Then
        Dim OptionStringC As String
        txtMasterPassword1.Text = MasterPassword
        txtMasterPassword2.Text = MasterPassword
        txtQuestion.Text = decrypt(GetSetting(MainTitle, UserRegSection, txtQuestion.Tag), Key1 & Mid$(txtQuestion.Tag, 3, 1) & Mid$(txtQuestion.Tag, 5, 1))
        txtAnswer.Text = decrypt(GetSetting(MainTitle, UserRegSection, txtAnswer.Tag), UserKeyword & Mid$(txtAnswer.Tag, 3, 1) & Mid$(txtAnswer.Tag, 5, 1))
        OptionStringC = decrypt(GetSetting(MainTitle, UserRegSection, frameLog.Tag), Key1 & Mid$(frameLog.Tag, 3, 1) & Mid$(frameLog.Tag, 5, 1))
        chkDisplay.Value = Left$(OptionStringC, 1)
        chkPassword.Value = Mid$(OptionStringC, 2, 1)
        chkRemove.Value = Mid$(OptionStringC, 3, 1)
        chkLog.Value = Mid$(OptionStringC, 4, 1)
        chkLogAll.Value = Mid$(OptionStringC, 5, 1)
        chkEncrypt.Value = Mid$(OptionStringC, 6, 1)
        If Len(OptionString) > 6 Then txtLogFile.Text = Mid$(OptionString, 7) Else txtLogFile.Text = ""
        Call chkLog_Click
    End If
    Screen.MousePointer = 0
    frmUserID.Hide
    
End Sub

Private Sub cmdDeleteUserID_Click()
    
    Title = "Delete User ID"
    Msg = "Warning: You are about to Delete the User ID " & UserID & ". " & vbCrLf
    Msg = Msg & "Deleting the User ID will remove all the user's data." & vbCrLf
    Msg = Msg & "Do you wish to delete this User ID?"
    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

    Response = MsgBox(Msg, DgDef, Title)
    If Response = IDYES Then
        GoTo DelUserID
    Else
        Exit Sub
    End If
    
DelUserID:
    tmpString2 = ""
    For currentChr = 1 To Len(Index) Step 6
        tmpString = Mid$(Index, currentChr, 6)
        If tmpString = UserRegSection Then GoTo DontAdd2Index
        tmpString2 = tmpString2 & tmpString
DontAdd2Index:
    Next

    If Len(tmpString2) = 0 Then Index = "/NEWRUN/" Else Index = tmpString2
    SaveSetting MainTitle, "Settings", "Index", crypt(Index, Key1 & Key2)
    DeleteSetting MainTitle, UserRegSection
    Unload frmMain
    txtUserID.Text = ""
    frmUserID.Hide
        
End Sub

Private Sub cmdOK_Click()

    If cmdOK.Caption = "&Create" Then
        If Len(Trim(txtUserID.Text)) = 0 Then
            MsgBox "Please choose a User ID.", 48, "Invalid User ID"
            txtUserID.SetFocus
            Exit Sub
        End If
        
        If Index = "/NEWRUN/" Then GoTo NoUserIDCheck
        tmpString = IsValidUserID(Trim(LCase$(txtUserID.Text)))
        If Not tmpString = "" Then
            MsgBox "The User ID you've chossen already exists in my database. Please choose another.", 48, "Invalid User ID"
            txtUserID.SelStart = 0
            txtUserID.SelLength = Len(txtUserID.Text)
            txtUserID.SetFocus
            Exit Sub
        End If
NoUserIDCheck:
    End If
    
        If Len(Trim(txtMasterPassword1.Text)) = 0 Then
            MsgBox "Please choose a Master Password to protect your data.", 48, MainTitle
            txtMasterPassword1.SelStart = 0
            txtMasterPassword1.SelLength = Len(txtMasterPassword1.Text)
            txtMasterPassword1.SetFocus
            Exit Sub
        End If
        
        If Not Trim(txtMasterPassword1.Text) = Trim(txtMasterPassword2.Text) Then
            MsgBox "Please make sure that you've confirmed your Master Password correctly!", 48, "Master Password NOT confirmed"
            txtMasterPassword2.SelStart = 0
            txtMasterPassword2.SelLength = Len(txtMasterPassword2.Text)
            txtMasterPassword2.SetFocus
            Exit Sub
        End If
        
        If Len(Trim(txtMasterPassword1.Text)) < 6 Then
            MsgBox "Master Password should consist of 6 characters at least.", 48, "Invalid Master Password"
            txtMasterPassword1.SelStart = 0
            txtMasterPassword1.SelLength = Len(txtMasterPassword1.Text)
            txtMasterPassword1.SetFocus
            Exit Sub
        End If
        
        
        If Len(Trim(txtQuestion.Text)) = 0 Then
            MsgBox "Incomplete information.", 48, MainTitle
            txtQuestion.SelStart = 0
            txtQuestion.SelLength = Len(txtQuestion.Text)
            txtQuestion.SetFocus
            Exit Sub
        End If
        
        If Len(Trim(txtAnswer.Text)) = 0 Then
            MsgBox "Incomplete information.", 48, MainTitle
            txtAnswer.SelStart = 0
            txtAnswer.SelLength = Len(txtAnswer.Text)
            txtAnswer.SetFocus
            Exit Sub '
        End If
        
         If Len(Trim(txtAnswer.Text)) < 6 Then
            MsgBox "Hint Answer should consist of 6 characters at least.", 48, "Invalid Master Password"
            txtAnswer.SelStart = 0
            txtAnswer.SelLength = Len(txtAnswer.Text)
            txtAnswer.SetFocus
            Exit Sub
        End If

        For currentCounter = 1 To Len(Trim(txtAnswer.Text))
            If Asc(Mid$(Trim(txtAnswer.Text), currentCounter, 1)) > 127 Then
                MsgBox "Invalid character in Hint Answer '" & Mid$(Trim(txtAnswer.Text), currentCounter, 1) & "'", 48, MainTitle
                Exit Sub
            End If
        Next currentCounter
        
        
        If Len(Trim(txtLogFile.Text)) = 0 And chkLog.Value = 1 Then
            MsgBox "Please enter the path of your Log File." & vbCrLf & "For example c:\logfile.txt", 48, "Access Loging"
            txtLogFile.SetFocus
            Exit Sub
        End If
                   
        If chkLog.Value = 1 Then
            
            ' =======================
            ' Analyze Logfile's path
            ' =======================
            txtLogFile.Text = Trim(txtLogFile.Text)
            
            Dim tmpLen As Long
            Dim curCharFile As Long
            Dim backSlashCount As Long
            Dim Extension As Boolean
            
            backSlashCount = 0
            For curCharFile = 1 To Len(txtLogFile.Text)
                If Mid$(txtLogFile.Text, curCharFile, 1) = "\" Then backSlashCount = backSlashCount + 1
            Next curCharFile
            
            Extension = False
            tmpLen = 0
            For curCharFile = 1 To Len(txtLogFile.Text)
                If Left$(Right$(txtLogFile.Text, curCharFile), 1) = "." Then
                    tmpLen = tmpLen + 1
                    If (curCharFile <= 5) And (curCharFile > 1) Then Extension = True
                End If
            Next curCharFile
            
            If (tmpLen >= 2) Or (tmpLen > 0 And Extension = False) Then
                MsgBox "The path you have specified for the Log File is invalid.", 48, MainTitle
                txtLogFile.SelStart = 0
                txtLogFile.SelLength = Len(txtLogFile.Text)
                txtLogFile.SetFocus
                Exit Sub
            End If
            
            If Not Extension Then txtLogFile.Text = txtLogFile.Text & ".TXT"
            
            If backSlashCount > 0 Then
                For curCharFile = 1 To Len(txtLogFile.Text)
                    If Left$(Right$(txtLogFile.Text, curCharFile), 1) = "\" Then
                        tmpString = Left$(txtLogFile.Text, Len(txtLogFile.Text) - curCharFile + 1)
                        GoTo CheckPath
                    End If
                Next
                
CheckPath:
                If Not (Len(Dir(tmpString)) > 0) Then
                    MsgBox "The path you have specified for the Log File is invalid." & vbCrLf & "The folder " & UCase(Left$(tmpString, Len(tmpString) - 1)) & " does NOT exist.", 48, MainTitle
                    txtLogFile.SelStart = 0
                    txtLogFile.SelLength = Len(txtLogFile.Text)
                    txtLogFile.SetFocus
                    Exit Sub
                End If
                
            Else
                txtLogFile.Text = App.Path & "\" & txtLogFile.Text
            
            End If
        End If
        
        If TrackUser = False And chkLog.Value = 1 Then
            MsgBox "The Access Loging feature will be activated the next time you run " & MainTitle & ".", 48, MailTitle
        End If
        
        ' ===================
        ' Start creating new User ID
        ' ===================
        Screen.MousePointer = 11
        Dim OptionsString As String
        Dim FlagString As String
        Dim encodedAnswer As String
        
        If cmdOK.Caption = "&Create" Then
            
            UserID = Trim(LCase$(txtUserID.Text))       ' User ID is Not case sensitive
            MasterPassword = Trim(txtMasterPassword1.Text)
            HintAnswer = Trim(txtAnswer.Text)
            txtUserID.Text = UserID
            txtMasterPassword1.Text = MasterPassword
            txtMasterPassword2.Text = MasterPassword
            txtQuestion.Text = Trim(txtQuestion.Text)
            txtAnswer.Text = Trim(txtAnswer.Text)
            txtLogFile.Text = Trim(txtLogFile.Text)
                 
            OptionsString = chkDisplay.Value & chkPassword.Value & chkRemove.Value & chkLog.Value & chkLogAll.Value & chkEncrypt.Value & txtLogFile.Text
            UserIndex = txtUserID.Tag & txtMasterPassword1.Tag & frameLog.Tag & txtQuestion.Tag & txtAnswer.Tag & txtMasterPassword2.Tag
            UserKeyword = Left$(UserID, 1) & Mid$(UserRegSection, 5, 2) & MasterPassword
            UserIDKeyword = Left$(UserRegSection, 2) & Right$(txtUserID.Tag, 2) & Key1
            MasterPasswordKeyword = Mid$(UserRegSection, 3, 1) & Left$(UserID, 1) & Left$(MasterPassword, Len(MasterPassword) - 5) & Right$(MasterPassword, 1) & Right$(txtMasterPassword1.Tag, 1) & Right$(UserID, 1)
            
            ' Encrypt Master Password
            FMasterPassword = cryptPassword(MasterPassword, MasterPasswordKeyword)
            
            ' The FlagString$ contains necessary encoded data which are obtained from
            ' the function cryptPassword. It will be only used to recover the Master Password
            ' if the user forgot it.
            
            FlagString = Len(txtAnswer.Text) & sDivide & txtAnswer.Text & sDivide
            For currentCounter = 16 To 1 Step -1
                FlagString = FlagString & FMasterPassword.arFlag(currentCounter)
                If Not currentCounter = 1 Then FlagString = FlagString & sDivide
            Next currentCounter
            
            ' Encode the Hint Answer
            encodedAnswer = AscEncode(txtAnswer.Text)
            ItemIndex = ""
            ItemCount = 0
            
            ' Save in Registry
            SaveSetting MainTitle, UserRegSection, txtUserID.Tag, crypt(UserID, UserIDKeyword)
            SaveSetting MainTitle, UserRegSection, txtMasterPassword1.Tag, FMasterPassword.encodedPassword
            SaveSetting MainTitle, UserRegSection, txtQuestion.Tag, crypt(txtQuestion.Text, Key1 & Mid$(txtQuestion.Tag, 3, 1) & Mid$(txtQuestion.Tag, 5, 1))
            SaveSetting MainTitle, UserRegSection, txtAnswer.Tag, crypt(txtAnswer.Text, UserKeyword & Mid$(txtAnswer.Tag, 3, 1) & Mid$(txtAnswer.Tag, 5, 1))
            SaveSetting MainTitle, UserRegSection, frameLog.Tag, crypt(OptionsString, Key1 & Mid$(frameLog.Tag, 3, 1) & Mid$(frameLog.Tag, 5, 1))
            SaveSetting MainTitle, UserRegSection, txtMasterPassword2.Tag, crypt(FlagString, encodedAnswer & Mid$(txtMasterPassword2.Tag, 5, 1))
            SaveSetting MainTitle, UserRegSection, "Index", crypt(UserIndex, Key2 & UserRegSection)
            SaveSetting MainTitle, UserRegSection, "Item", ""
            If Index = "/NEWRUN/" Then Index = UserRegSection Else Index = Index & UserRegSection
            SaveSetting MainTitle, "Settings", "Index", crypt(Index, Key1 & Key2)
            Screen.MousePointer = 0
            MsgBox "The User ID " & UserID & " was successfuly registered.", 64, MainTitle
            
            
        ElseIf cmdOK.Caption = "&Okay" Then
            
            txtMasterPassword1.Text = Trim(txtMasterPassword1.Text)
            txtQuestion.Text = Trim(txtQuestion.Text)
            txtAnswer.Text = Trim(txtAnswer.Text)
            txtLogFile.Text = Trim(txtLogFile.Text)
            OptionsString = chkDisplay.Value & chkPassword.Value & chkRemove.Value & chkLog.Value & chkLogAll.Value & chkEncrypt.Value & txtLogFile.Text
            SaveSetting MainTitle, UserRegSection, txtQuestion.Tag, crypt(txtQuestion.Text, Key1 & Mid$(txtQuestion.Tag, 3, 1) & Mid$(txtQuestion.Tag, 5, 1))
            SaveSetting MainTitle, UserRegSection, txtAnswer.Tag, crypt(txtAnswer.Text, UserKeyword & Mid$(txtAnswer.Tag, 3, 1) & Mid$(txtAnswer.Tag, 5, 1))
            SaveSetting MainTitle, UserRegSection, frameLog.Tag, crypt(OptionsString, Key1 & Mid$(frameLog.Tag, 3, 1) & Mid$(frameLog.Tag, 5, 1))
            MasterPasswordKeyword = Mid$(UserRegSection, 3, 1) & Left$(UserID, 1) & Left$(txtMasterPassword1.Text, Len(txtMasterPassword1.Text) - 5) & Right$(txtMasterPassword1.Text, 1) & Right$(txtMasterPassword1.Tag, 1) & Right$(UserID, 1)
            
            If Not txtAnswer.Text = HintAnswer Then
                
                ' User changed Hint Answer
                
                hintAnswerKeyword = AscEncode(txtAnswer.Text)
                FMasterPassword = cryptPassword(txtMasterPassword1.Text, MasterPasswordKeyword)
                FlagString = Len(txtAnswer.Text) & sDivide & txtAnswer.Text & sDivide
                For currentCounter = 16 To 1 Step -1
                    FlagString = FlagString & FMasterPassword.arFlag(currentCounter)
                    If Not currentCounter = 1 Then FlagString = FlagString & sDivide
                Next currentCounter
                SaveSetting MainTitle, UserRegSection, txtMasterPassword2.Tag, crypt(FlagString, hintAnswerKeyword & Mid$(txtMasterPassword2.Tag, 5, 1))
                HintAnswer = txtAnswer.Text
            End If
            
            If Not txtMasterPassword1.Text = MasterPassword Then
                ' User changed the Master Password; all the stored data should be decrypted and
                ' encrypted again with the new UserKeyword!
                Call LogAction(LogWord(3))
                Dim oldUserKeyword As String
                Dim sKey As String
                Dim curConverKey As Long
                
                oldUserKeyword = UserKeyword
                MasterPassword = txtMasterPassword1.Text
                UserKeyword = Left$(UserID, 1) & Mid$(UserRegSection, 5, 2) & MasterPassword
                
                For curConverKey = 1 To Len(ItemIndex) Step 8
                    sKey = Mid$(ItemIndex, curConverKey, 8)
                    regString = decrypt(GetSetting(MainTitle, UserRegSection, sKey), oldUserKeyword & Mid$(sKey, 3, 1) & Mid$(sKey, 5, 1))
                    SaveSetting MainTitle, UserRegSection, sKey, crypt(regString, UserKeyword & Mid$(sKey, 3, 1) & Mid$(sKey, 5, 1))
                Next
                regString = decrypt(GetSetting(MainTitle, UserRegSection, txtAnswer.Tag), oldUserKeyword & Mid$(txtAnswer.Tag, 3, 1) & Mid$(txtAnswer.Tag, 5, 1))
                SaveSetting MainTitle, UserRegSection, txtAnswer.Tag, crypt(regString, UserKeyword & Mid$(txtAnswer.Tag, 3, 1) & Mid$(txtAnswer.Tag, 5, 1))
                     
                ' Update Master Password flags
                hintAnswerKeyword = AscEncode(txtAnswer.Text)
                FMasterPassword = cryptPassword(txtMasterPassword1.Text, MasterPasswordKeyword)
                FlagString = Len(txtAnswer.Text) & sDivide & txtAnswer.Text & sDivide
                For currentCounter = 16 To 1 Step -1
                    FlagString = FlagString & FMasterPassword.arFlag(currentCounter)
                    If Not currentCounter = 1 Then FlagString = FlagString & sDivide
                Next currentCounter
                SaveSetting MainTitle, UserRegSection, txtMasterPassword2.Tag, crypt(FlagString, hintAnswerKeyword & Mid$(txtMasterPassword2.Tag, 5, 1))
                HintAnswer = txtAnswer.Text
                
                SaveSetting MainTitle, UserRegSection, "Item", crypt(ItemIndex, UserKeyword)
                SaveSetting MainTitle, UserRegSection, Mid$(UserIndex, 9, 8), FMasterPassword.encodedPassword
                
            End If
            
        End If
        
        Screen.MousePointer = 0
        frmUserID.Hide
         
End Sub

