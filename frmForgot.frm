VERSION 5.00
Begin VB.Form frmForgot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm your identity"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
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
   ScaleHeight     =   3660
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMasterPassword 
      Height          =   330
      Left            =   1800
      TabIndex        =   11
      Top             =   2775
      Width           =   3855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Proceed"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   -120
      TabIndex        =   7
      Top             =   1440
      Width           =   6255
   End
   Begin VB.TextBox txtAnswer 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Type in your Hint Answer and click Proceed, then, Password Guard will recover your Master Password."
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label Label6 
      Caption         =   "Your Master Password:"
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
      Left            =   75
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Hint Answer :"
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
      Left            =   60
      TabIndex        =   5
      Top             =   2400
      Width           =   1305
   End
   Begin VB.Label lblQuestion 
      AutoSize        =   -1  'True
      Caption         =   "HintQuestion"
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Label lblUserID 
      AutoSize        =   -1  'True
      Caption         =   "UserID"
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
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
      Left            =   90
      TabIndex        =   2
      Top             =   2040
      Width           =   1395
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   240
      Picture         =   "frmForgot.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   570
   End
   Begin VB.Label Label2 
      Caption         =   "Please confirm your identity by typing the Answer of your Hint Question (Case Sensitive)."
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   795
   End
End
Attribute VB_Name = "frmForgot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    
    frmForgot.Hide
    frmLogIn.Show
    
End Sub

Private Sub cmdOK_Click()
    
    Dim hintAnswerKeyword As String
    Dim arFlag() As String, arKey(1 To 16) As Long
    Dim AnswerLen As Integer, curKeyRef As Integer
    Dim resultMPasswordKey As String
    
    txtAnswer.Text = Trim(txtAnswer.Text)
    hintAnswerKeyword = AscEncode(txtAnswer.Text)
    If hintAnswerKeyword = "" Then GoTo InvalidAnswer
    
    On Error GoTo InvalidAnswer
    tmpString = decrypt(GetSetting(MainTitle, UserRegSection, txtAnswer.Tag), hintAnswerKeyword & Mid$(txtAnswer.Tag, 5, 1))
    On Error GoTo InvalidAnswer
    arFlag = Split(tmpString, sDivide)
                
    On Error GoTo InvalidAnswer
    AnswerLen = Int(arFlag(LBound(arFlag)))
    
    ' Compare user-provided Hint Answer with the stored one
    If Not AnswerLen = Len(txtAnswer.Text) Then GoTo InvalidAnswer
    If Not txtAnswer.Text = arFlag(LBound(arFlag) + 1) Then GoTo InvalidAnswer
    
    curKeyRef = 0
    resultMPassword = ""
    tmpAsc = 0
    For currentCounter = 1 To Len(hintAnswerKeyword)
        tmpAsc = tmpAsc + Asc(Mid$(hintAnswerKeyword, currentCounter, 1))
    Next
        
    ' Restore flags from arFlag() string array
    For currentCounter = LBound(arFlag) + 2 To UBound(arFlag)
        curKeyRef = curKeyRef + 1
        On Error GoTo ErrAnswer
        arKey(curKeyRef) = arFlag(currentCounter)
'        arKey(curKeyRef) = arKey(curKeyRef) Xor tmpAsc
    Next
    
    resultMPasswordKey = ""
    For currentCounter = 16 To 1 Step -1
        On Error GoTo ErrAnswer
        If Not arKey(currentCounter) = 0 Then resultMPasswordKey = resultMPasswordKey & Chr$(arKey(currentCounter))
    Next
    txtMasterPassword.Text = decrypt(GetSetting(MainTitle, UserRegSection, txtMasterPassword.Tag), resultMPasswordKey)
    txtAnswer.Locked = True
    cmdOK.Enabled = False
    frmLogIn.txtMasterPassword.Text = txtMasterPassword.Text
    cmdCancel.Caption = "&Close"
    Screen.MousePointer = 0
    Exit Sub


InvalidAnswer:
        
        Screen.MousePointer = 0
        MsgBox "Sorry. Invalid Answer.", 16, "Access Denied"
        txtAnswer.SelStart = 0
        txtAnswer.SelLength = Len(txtAnswer.Text)
        txtAnswer.SetFocus
        Exit Sub

ErrAnswer:
    Screen.MousePointer = 0
    MsgBox "Unexpected error occured during decryption process!", 16, MainTitle
    Exit Sub
End Sub

Private Sub txtAnswer_Change()
    If Len(Trim(txtAnswer.Text)) > 0 Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
    
End Sub

