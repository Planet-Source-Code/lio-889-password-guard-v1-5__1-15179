VERSION 5.00
Begin VB.Form frmLogIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Guard"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
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
   ScaleHeight     =   2340
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create new User ID"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "&Login"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtMasterPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1440
      Width           =   3615
   End
   Begin VB.ComboBox lstUserID 
      Height          =   315
      ItemData        =   "frmLogIn.frx":0000
      Left            =   1440
      List            =   "frmLogIn.frx":0002
      TabIndex        =   4
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label lblForgotMasterPassword 
      Caption         =   "Forgot your Master Password?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   445
      Left            =   5160
      MouseIcon       =   "frmLogIn.frx":0004
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblMasterPasswordCap 
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
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   990
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
      TabIndex        =   2
      Top             =   960
      Width           =   795
   End
   Begin VB.Image imgLock 
      Height          =   480
      Left            =   240
      Picture         =   "frmLogIn.frx":030E
      Top             =   360
      Width           =   435
   End
   Begin VB.Label lblEnter 
      AutoSize        =   -1  'True
      Caption         =   "Please enter your User ID and your Master Password."
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   4620
   End
   Begin VB.Label lblWelcome 
      AutoSize        =   -1  'True
      Caption         =   "Welcome to Password Guard!"
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
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2835
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCreate_Click()
       
On Error Resume Next
Unload frmUserID
Load frmUserID

GenerateNewRegSection:
        UserRegSection = RandomPinString(3)
    
        ' Check whether any of the previous stored Users matches the same
        ' registry section.
        If Not Index = "/NEWRUN/" Then
            For currentChr = 1 To Len(Index) Step 6
                tmpString = Mid$(Index, currentChr, 6)
                If tmpString = UserRegSection Then GoTo GenerateNewRegSection
            Next
        End If
        
        ' Generate UserID's Reg Key name
        frmUserID.txtUserID.Tag = RandomPinString(4)

        ' Generate MasterPassword's Reg Key name
GenerateMasterPasswordKey:
        frmUserID.txtMasterPassword1.Tag = RandomPinString(4)
        If frmUserID.txtMasterPassword1.Tag = frmUserID.txtUserID.Tag Then GoTo GenerateMasterPasswordKey
        
        ' Generate Alternative Question's Reg Key name
GenerateQuestionKey:
        frmUserID.txtQuestion.Tag = RandomPinString(4)
        If frmUserID.txtQuestion.Tag = frmUserID.txtUserID.Tag Then GoTo GenerateQuestionKey
        If frmUserID.txtQuestion.Tag = frmUserID.txtMasterPassword1.Tag Then GoTo GenerateQuestionKey
        
        
        ' Generate Alternative Answer's Reg Key name
GenerateAnswerKey:
        frmUserID.txtAnswer.Tag = RandomPinString(4)
        If frmUserID.txtAnswer.Tag = frmUserID.txtUserID.Tag Then GoTo GenerateAnswerKey
        If frmUserID.txtAnswer.Tag = frmUserID.txtMasterPassword1.Tag Then GoTo GenerateAnswerKey
        If frmUserID.txtAnswer.Tag = frmUserID.txtQuestion.Tag Then GoTo GenerateAnswerKey

        ' Generate Additional options' Reg Key name
GenerateOptionsKey:
        
        frmUserID.frameLog.Tag = RandomPinString(4)
        If frmUserID.frameLog.Tag = frmUserID.txtUserID.Tag Then GoTo GenerateOptionsKey
        If frmUserID.frameLog.Tag = frmUserID.txtMasterPassword1.Tag Then GoTo GenerateOptionsKey
        If frmUserID.frameLog.Tag = frmUserID.txtQuestion.Tag Then GoTo GenerateOptionsKey
        If frmUserID.frameLog.Tag = frmUserID.txtAnswer.Tag Then GoTo GenerateOptionsKey
        
GenerateFlagsKey:
        frmUserID.txtMasterPassword2.Tag = RandomPinString(4)
        If frmUserID.txtMasterPassword2.Tag = frmUserID.txtUserID.Tag Then GoTo GenerateFlagsKey
        If frmUserID.txtMasterPassword2.Tag = frmUserID.txtMasterPassword1.Tag Then GoTo GenerateFlagsKey
        If frmUserID.txtMasterPassword2.Tag = frmUserID.txtQuestion.Tag Then GoTo GenerateFlagsKey
        If frmUserID.txtMasterPassword2.Tag = frmUserID.txtAnswer.Tag Then GoTo GenerateFlagsKey
        If frmUserID.txtMasterPassword2.Tag = frmUserID.frameLog.Tag Then GoTo GenerateFlagsKey
        
        frmUserID.Caption = "Create a new User ID"
        frmUserID.cmdOK.Caption = "&Create"
        frmUserID.chkLog.Value = 0
        frmUserID.chkLogAll.Enabled = False
        frmUserID.chkEncrypt.Enabled = False
        frmUserID.txtLogFile.Enabled = False
        frmUserID.txtLogFile.BackColor = NoActive
        frmUserID.frameDelete.Visible = False
        frmUserID.cmdViewLog.Visible = False
        frmUserID.Show 1
        
        If frmUserID.txtUserID.Text = "" Then
            Unload frmUserID
            Exit Sub
        End If
        
        frmUserID.Caption = "User ID settings"
        frmUserID.txtUserID.Locked = True
        frmUserID.frameDelete.Visible = True
        frmUserID.cmdViewLog.Visible = True
        frmUserID.cmdOK.Caption = "&OK"
        
        frmLogIn.Hide
        LogIn
        
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdSubmit_Click()
           
    If Len(Trim(lstUserID.Text)) = 0 Then GoTo LoginAccessDenied
    If Len(Trim(txtMasterPassword.Text)) < 5 Then GoTo LoginAccessDenied
    If Len(Trim(txtMasterPassword.Text)) > 12 Then GoTo LoginAccessDenied
    
    UserIndex = IsValidUserID(Trim(LCase$(lstUserID.Text)))
    If UserIndex = "" Then GoTo LoginAccessDenied
    UserRegSection = Left$(UserIndex, 6)
    UserIndex = Right$(UserIndex, Len(UserIndex) - 6)
    UserKeyword = IsValidMasterPassword(UserRegSection, Mid$(UserIndex, 9, 8), Trim(LCase$(lstUserID.Text)), txtMasterPassword.Text)
    If UserKeyword = "" Then GoTo LoginAccessDenied
    UserID = Trim(LCase$(lstUserID.Text))                 ' User ID is NOT case sensitive.
    MasterPassword = Trim(txtMasterPassword.Text) ' Master Password IS case sensitive.
    lstUserID.Text = ""
    txtMasterPassword.Text = ""
    frmLogIn.Hide
    
    On Error Resume Next
    Unload frmMain
    On Error Resume Next
    Unload frmUserID
    Load frmUserID
    Load frmMain
    Call LogIn 'Log User in
    
    Exit Sub
    
LoginAccessDenied:
    MsgBox "One or more error occured: invalid User ID or Master Password.", 16, "Access Denied"
    lstUserID.SelStart = 0
    lstUserID.SelLength = Len(lstUserID.Text)
    lstUserID.SetFocus
    Exit Sub
    
End Sub


Private Sub lblForgotMasterPassword_Click()
    
    On Error Resume Next
    Unload frmForgot
    Load frmForgot
    
    If Len(Trim(lstUserID.Text)) = 0 Then
        MsgBox "Please enter your User ID first.", 48, MainTitle
        lstUserID.SelStart = 0
        lstUserID.SelLength = Len(lstUserID.Text)
        lstUserID.SetFocus
        Exit Sub
    End If
    
    UserIndex = IsValidUserID(Trim(LCase$(lstUserID.Text)))
    If UserIndex = "" Then
        MsgBox "Invalid User ID.", 16, MainTitle
        lstUserID.SelStart = 0
        lstUserID.SelLength = Len(lstUserID.Text)
        lstUserID.SetFocus
        Exit Sub
    End If
    
    txtMasterPassword.Text = ""
    UserRegSection = Left$(UserIndex, 6)
    UserIndex = Right$(UserIndex, Len(UserIndex) - 6)
    frmForgot.cmdOK.Enabled = False
    frmForgot.txtMasterPassword.Tag = Mid$(UserIndex, 9, 8)
    hintQuestionKey = Mid$(UserIndex, 25, 8)
    frmForgot.txtAnswer.Tag = Mid$(UserIndex, 41, 8)
    frmForgot.lblUserID.Caption = Trim(LCase$(lstUserID.Text))
    frmForgot.lblQuestion.Caption = decrypt(GetSetting(MainTitle, UserRegSection, hintQuestionKey), Key1 & Mid$(hintQuestionKey, 3, 1) & Mid$(hintQuestionKey, 5, 1))
    frmForgot.txtMasterPassword.Locked = True
    
    frmLogIn.Hide
    frmForgot.Show
    
End Sub

Private Sub lstUserID_Change()
If Len(lstUserID.Text) > 0 And Len(txtMasterPassword.Text) > 0 Then
    cmdSubmit.Enabled = True
Else
    cmdSubmit.Enabled = False
End If

End Sub

Private Sub txtMasterPassword_Change()
If Len(lstUserID.Text) > 0 And Len(txtMasterPassword.Text) > 0 Then
    cmdSubmit.Enabled = True
Else
    cmdSubmit.Enabled = False
End If

End Sub

Private Sub txtMasterPassword_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Len(lstUserID.Text) > 0 And Len(txtMasterPassword.Text) > 0 Then
            cmdSubmit_Click
        Else
            Beep
            Exit Sub
        End If
    End If

End Sub
