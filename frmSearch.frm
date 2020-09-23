VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
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
   ScaleHeight     =   2970
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNotes 
      Caption         =   "Notes"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtNotes 
      Height          =   615
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   2280
      Width           =   2775
   End
   Begin VB.OptionButton optAny 
      Caption         =   "Match any"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4920
      TabIndex        =   12
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton optAll 
      Caption         =   "Match all"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   2775
   End
   Begin VB.CheckBox chkPassword 
      Caption         =   "Password"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox chkUserName 
      Caption         =   "User Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox chkServer 
      Caption         =   "Server"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CheckBox chkDescription 
      Caption         =   "Description"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblSearchCap 
      Caption         =   "Please supply the information you'd like search for."
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   240
      Width           =   4455
   End
   Begin VB.Image imgSearch 
      Height          =   435
      Left            =   240
      Picture         =   "frmSearch.frx":0000
      Top             =   240
      Width           =   405
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub chkDescription_Click()
    
    If chkDescription.Value = 0 Then
        txtDescription.Enabled = False
        txtDescription.BackColor = NoActive
    ElseIf chkDescription.Value = 1 Then
        txtDescription.Enabled = True
        txtDescription.BackColor = Active
    End If
    
    
End Sub

Public Sub chkNotes_Click()
    
    If chkNotes.Value = 0 Then
        txtNotes.Enabled = False
        txtNotes.BackColor = NoActive
    ElseIf chkNotes.Value = 1 Then
        txtNotes.Enabled = True
        txtNotes.BackColor = Active
    End If

End Sub

Public Sub chkPassword_Click()
    
    If chkPassword.Value = 0 Then
        txtPassword.Enabled = False
        txtPassword.BackColor = NoActive
    ElseIf chkPassword.Value = 1 Then
        txtPassword.Enabled = True
        txtPassword.BackColor = Active
    End If

End Sub

Public Sub chkServer_Click()
    
    If chkServer.Value = 0 Then
        txtServer.Enabled = False
        txtServer.BackColor = NoActive
    ElseIf chkServer.Value = 1 Then
        txtServer.Enabled = True
        txtServer.BackColor = Active
    End If

End Sub

Public Sub chkUserName_Click()
    If chkUserName.Value = 0 Then
        txtUserName.Enabled = False
        txtUserName.BackColor = NoActive
    ElseIf chkUserName.Value = 1 Then
        txtUserName.Enabled = True
        txtUserName.BackColor = Active
    End If

End Sub

Private Sub cmdCancel_Click()
    cmdSearch.Tag = "0"
    frmSearch.Hide
    
End Sub

Private Sub cmdSearch_Click()
    
    Dim CheckCount As Long
    CheckCount = 0
    CheckCount = chkDescription.Value + chkServer.Value + chkUserName.Value
    CheckCount = CheckCount + chkPassword.Value + chkNotes.Value
    
    If CheckCount = 0 Then
        MsgBox "Can't search with no keywords!", 48, MainTitle
        Exit Sub
    End If
    
    If chkDescription.Value = 1 And Len(Trim(txtDescription.Text)) = 0 Then
        MsgBox "Please fill in the Description field.", 48, MainTitle
        txtDescription.SelStart = 0
        txtDescription.SelLength = Len(txtDescription.Text)
        txtDescription.SetFocus
        Exit Sub
    End If
    
    If chkServer.Value = 1 And Len(Trim(txtServer.Text)) = 0 Then
        MsgBox "Please fill in the Server field.", 48, MainTitle
        txtServer.SelStart = 0
        txtServer.SelLength = Len(txtServer.Text)
        txtServer.SetFocus
        Exit Sub
    End If

    If chkUserName.Value = 1 And Len(Trim(txtUserName.Text)) = 0 Then
        MsgBox "Please fill in the User Name field.", 48, MainTitle
        txtUserName.SelStart = 0
        txtUserName.SelLength = Len(txtUserName.Text)
        txtUserName.SetFocus
        Exit Sub
    End If

    If chkPassword.Value = 1 And Len(Trim(txtPassword.Text)) = 0 Then
        MsgBox "Please fill in the Password field.", 48, MainTitle
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
        txtPassword.SetFocus
        Exit Sub
    End If

    If chkNotes.Value = 1 And Len(Trim(txtNotes.Text)) = 0 Then
        MsgBox "Please fill in the Notes field.", 48, MainTitle
        txtNotes.SelStart = 0
        txtNotes.SelLength = Len(txtNotes.Text)
        txtNotes.SetFocus
        Exit Sub
    End If
    
    If chkDescription.Value = 0 Then txtDescription.Text = ""
    If chkServer.Value = 0 Then txtServer.Text = ""
    If chkUserName.Value = 0 Then txtUserName.Text = ""
    If chkPassword.Value = 0 Then txtPassword.Text = ""
    If chkNotes.Value = 0 Then txtNotes.Text = ""
    
    
    txtDescription.Text = Trim(txtDescription.Text)
    txtServer.Text = Trim(txtServer.Text)
    txtUserName.Text = Trim(txtUserName.Text)
    txtPassword.Text = Trim(txtPassword.Text)
    txtNotes.Text = Trim(txtNotes.Text)
    
    SaveSetting MainTitle, UserRegSection, "searchDescription", chkDescription.Value
    SaveSetting MainTitle, UserRegSection, "searchServer", chkServer.Value
    SaveSetting MainTitle, UserRegSection, "searchUserName", chkUserName.Value
    SaveSetting MainTitle, UserRegSection, "searchPassword", chkPassword.Value
    SaveSetting MainTitle, UserRegSection, "searchNotes", chkNotes.Value
    
    If optAny.Value = True Then SaveSetting MainTitle, UserRegSection, "searchMode", "1"
    If optAll.Value = True Then SaveSetting MainTitle, UserRegSection, "searchMode", "2"
    
    cmdSearch.Tag = "1"
    frmSearch.Hide
    
End Sub

