VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Guard"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNotes 
      BackColor       =   &H00E0FEFE&
      Height          =   975
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   4320
      Width           =   7815
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00E0FEFE&
      Height          =   285
      Left            =   5760
      TabIndex        =   11
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00E0FEFE&
      Height          =   285
      Left            =   5760
      TabIndex        =   9
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00E0FEFE&
      Height          =   285
      Left            =   5760
      TabIndex        =   7
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox txtServer 
      BackColor       =   &H00E0FEFE&
      Height          =   285
      Left            =   5760
      TabIndex        =   4
      Top             =   2280
      Width           =   3375
   End
   Begin VB.PictureBox boxTools 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1130
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   9195
      TabIndex        =   3
      Top             =   0
      Width           =   9255
      Begin VB.Timer tmrPos 
         Left            =   7440
         Top             =   120
      End
      Begin MSComDlg.CommonDialog dlg 
         Left            =   7920
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList img 
         Left            =   8520
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0CCA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image imgMinimize 
         Height          =   480
         Left            =   4110
         Picture         =   "frmMain.frx":0E26
         Top             =   165
         Width           =   480
      End
      Begin VB.Label lblMinimize 
         AutoSize        =   -1  'True
         Caption         =   "Minimize"
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
         Left            =   3960
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.Image imgImport 
         Height          =   510
         Left            =   2520
         Picture         =   "frmMain.frx":1130
         Stretch         =   -1  'True
         Top             =   90
         Width           =   615
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         Caption         =   "About"
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
         Left            =   5160
         TabIndex        =   20
         Top             =   720
         Width           =   570
      End
      Begin VB.Image imgAbout 
         Height          =   465
         Left            =   5310
         Picture         =   "frmMain.frx":2386
         Stretch         =   -1  'True
         Top             =   165
         Width           =   210
      End
      Begin VB.Label lblImport 
         Alignment       =   2  'Center
         Caption         =   "Import Records from File"
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
         Left            =   2040
         TabIndex        =   18
         Top             =   645
         Width           =   1605
      End
      Begin VB.Label lblExport 
         Alignment       =   2  'Center
         Caption         =   "Export Records to File"
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
         Left            =   120
         TabIndex        =   17
         Top             =   645
         Width           =   1560
      End
      Begin VB.Image imgExport 
         Height          =   510
         Left            =   480
         Picture         =   "frmMain.frx":29D0
         Stretch         =   -1  'True
         Top             =   90
         Width           =   630
      End
   End
   Begin VB.PictureBox boxSide 
      Align           =   3  'Align Left
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4290
      Left            =   0
      ScaleHeight     =   4230
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   1125
      Width           =   1215
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         Caption         =   "Search"
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
         TabIndex        =   19
         Top             =   3120
         Width           =   675
      End
      Begin VB.Image imgSearch 
         Height          =   435
         Left            =   360
         Picture         =   "frmMain.frx":3CB2
         Top             =   2640
         Width           =   405
      End
      Begin VB.Image imgRemove 
         Height          =   255
         Left            =   480
         Picture         =   "frmMain.frx":46F4
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   270
      End
      Begin VB.Image imgAdd 
         Height          =   465
         Left            =   240
         Picture         =   "frmMain.frx":4B76
         Stretch         =   -1  'True
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblRemove 
         Alignment       =   2  'Center
         Caption         =   "Remove Record"
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
         Left            =   200
         TabIndex        =   15
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label lblAdd 
         Alignment       =   2  'Center
         Caption         =   "Add Record"
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
         Left            =   200
         TabIndex        =   14
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.PictureBox boxStatus 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9195
      TabIndex        =   1
      Top             =   5415
      Width           =   9255
      Begin VB.PictureBox picProgress 
         Height          =   255
         Left            =   1800
         ScaleHeight     =   195
         ScaleWidth      =   2355
         TabIndex        =   22
         Top             =   30
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "StatusBar"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   60
         Width           =   840
      End
   End
   Begin MSComctlLib.ListView lstItem 
      Height          =   2175
      Left            =   1320
      TabIndex        =   0
      Top             =   1800
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3836
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "img"
      SmallIcons      =   "img"
      ColHdrIcons     =   "ImageList"
      ForeColor       =   -2147483640
      BackColor       =   14745342
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Image imgItems 
      Height          =   480
      Left            =   1320
      Picture         =   "frmMain.frx":5BC0
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image imgDescription 
      Height          =   300
      Left            =   5760
      Picture         =   "frmMain.frx":688A
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   255
   End
   Begin VB.Image imgPassword 
      Height          =   255
      Left            =   5800
      Picture         =   "frmMain.frx":6D6C
      Stretch         =   -1  'True
      Top             =   3410
      Width           =   195
   End
   Begin VB.Image imgServer 
      Height          =   290
      Left            =   5760
      Picture         =   "frmMain.frx":70CE
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   195
      Left            =   1320
      Picture         =   "frmMain.frx":7578
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   240
   End
   Begin VB.Image imgUserName 
      Height          =   285
      Left            =   5640
      Picture         =   "frmMain.frx":76AA
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   360
   End
   Begin VB.Label lblItemsCap 
      AutoSize        =   -1  'True
      Caption         =   "Items:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   1920
      TabIndex        =   16
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label lblNotesCap 
      AutoSize        =   -1  'True
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   1680
      TabIndex        =   12
      Top             =   4080
      Width           =   645
   End
   Begin VB.Label lblDescriptionCap 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   6120
      TabIndex        =   10
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label lblPasswordCap 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   6120
      TabIndex        =   8
      Top             =   3480
      Width           =   1050
   End
   Begin VB.Label lblUserCap 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   6120
      TabIndex        =   6
      Top             =   2760
      Width           =   1170
   End
   Begin VB.Label lblServerCap 
      AutoSize        =   -1  'True
      Caption         =   "Server :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   6120
      TabIndex        =   5
      Top             =   2040
      Width           =   810
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExport 
         Caption         =   "&Export Records to File..."
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import Records from File..."
      End
      Begin VB.Menu Null0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuAddRecord 
         Caption         =   "Add &New Record"
      End
      Begin VB.Menu mnuRemoveRecord 
         Caption         =   "&Remove Current Record"
      End
      Begin VB.Menu Null1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search..."
      End
      Begin VB.Menu mnuViewSearchResult 
         Caption         =   "&View last search's results..."
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "&User"
      Begin VB.Menu mnuUserIDSettings 
         Caption         =   "User ID &settings..."
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Log&out..."
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About..."
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CursorPosition As Point
Private Sub boxSide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(0)

End Sub

Private Sub boxStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(0)

End Sub

Private Sub boxTools_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(0)

End Sub

Private Sub Form_Activate()
   
   If frmUserID.chkPassword.Value = 0 Then
        txtPassword.PasswordChar = ""
    ElseIf frmUserID.chkPassword.Value = 1 Then
        txtPassword.PasswordChar = "*"
    End If
    
End Sub

Private Sub Form_Load()
    
    lblStatus.Caption = ""

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(0)

End Sub

Private Sub imgAbout_Click()
mnuAbout_Click

End Sub

Private Sub imgAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(3)

End Sub

Private Sub imgAdd_Click()
mnuAddRecord_Click

End Sub

Private Sub imgAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(4)

End Sub

Private Sub imgExport_Click()
mnuExport_Click

End Sub

Private Sub imgExport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(1)

End Sub

Private Sub imgImport_Click()
mnuImport_Click

End Sub

Private Sub imgImport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(2)

End Sub

Private Sub imgItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(0)

End Sub

Private Sub imgMinimize_Click()
frmMain.WindowState = 1

End Sub

Private Sub imgMinimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(7)

End Sub

Private Sub imgRemove_Click()
mnuRemoveRecord_Click

End Sub

Private Sub imgRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(5)

End Sub

Private Sub imgSearch_Click()
mnuSearch_Click

End Sub

Private Sub imgSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(6)

End Sub

Private Sub lblAbout_Click()
mnuAbout_Click

End Sub

Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(3)

End Sub

Private Sub lblAdd_Click()
mnuAddRecord_Click

End Sub

Private Sub lblAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(4)

End Sub

Private Sub lblExport_Click()
mnuExport_Click

End Sub

Private Sub lblExport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(1)

End Sub

Private Sub lblImport_Click()
mnuImport_Click

End Sub

Private Sub lblImport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(2)

End Sub

Private Sub lblItemsCap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(0)

End Sub

Private Sub lblMinimize_Click()
frmMain.WindowState = 1

End Sub

Private Sub lblMinimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(7)

End Sub

Private Sub lblRemove_Click()
mnuRemoveRecord_Click

End Sub

Private Sub lblRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(5)

End Sub

Private Sub lblSearch_Click()
mnuSearch_Click

End Sub

Private Sub lblSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(6)

End Sub

Private Sub lblStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(0)

End Sub

Private Sub lstItem_Click()
    If ItemCount = 0 Then Exit Sub
    lstItem.SelectedItem.Selected = True

End Sub

Private Sub lstItem_DblClick()
    If ItemCount = 0 Then Exit Sub
    lstItem.SelectedItem.Selected = True
    
End Sub

Private Sub lstItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
   
    If lstItem.SelectedItem.Key = currentKey Then Exit Sub ' User just clicked the selected
                                                                                     ' ListItem, exit sub...
                                                                                     
    ' Because always the event lstItem_ItemClick happens before the TextBox_LostFocus event,
    ' we need to make sure that the data are saved before they will be read again.
    
    Call UpdateRecord
    
ProcessItemClick:
    Dim recordOutput As DataRecord
    currentKey = lstItem.SelectedItem.Key
    currentRecord = lstItem.SelectedItem.Index
    
    ' Read selected record
    recordOutput = ReadRecord(currentKey)
    tDescription = recordOutput.Description
    tServer = recordOutput.Server
    tUserName = recordOutput.UserName
    tPassword = recordOutput.Password
    tNotes = recordOutput.Notes
    
    txtDescription.Text = tDescription
    txtServer.Text = tServer
    txtUserName.Text = tUserName
    txtPassword.Text = tPassword
    txtNotes.Text = tNotes
    
End Sub

Private Sub lstItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(0)

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1

End Sub

Private Sub LightLabel(LabelIndex As Integer)
    
    If LabelIndex = 0 Then
        lblExport.ForeColor = LightOff
        lblExport.Font.Underline = False
        lblImport.ForeColor = LightOff
        lblImport.Font.Underline = False
        lblAbout.ForeColor = LightOff
        lblAbout.Font.Underline = False
        lblAdd.ForeColor = LightOff
        lblAdd.Font.Underline = False
        lblRemove.ForeColor = LightOff
        lblRemove.Font.Underline = False
        lblSearch.ForeColor = LightOff
        lblSearch.Font.Underline = False
        lblMinimize.ForeColor = LightOff
        lblMinimize.Font.Underline = False
        
    ElseIf LabelIndex = 1 Then
        lblExport.ForeColor = LightOn
        lblExport.Font.Underline = True
        lblImport.ForeColor = LightOff
        lblImport.Font.Underline = False
        lblAbout.ForeColor = LightOff
        lblAbout.Font.Underline = False
        lblAdd.ForeColor = LightOff
        lblAdd.Font.Underline = False
        lblRemove.ForeColor = LightOff
        lblRemove.Font.Underline = False
        lblSearch.ForeColor = LightOff
        lblSearch.Font.Underline = False
        lblMinimize.ForeColor = LightOff
        lblMinimize.Font.Underline = False

    ElseIf LabelIndex = 2 Then
        lblExport.ForeColor = LightOff
        lblExport.Font.Underline = False
        lblImport.ForeColor = LightOn
        lblImport.Font.Underline = True
        lblAbout.ForeColor = LightOff
        lblAbout.Font.Underline = False
        lblAdd.ForeColor = LightOff
        lblAdd.Font.Underline = False
        lblRemove.ForeColor = LightOff
        lblRemove.Font.Underline = False
        lblSearch.ForeColor = LightOff
        lblSearch.Font.Underline = False
        lblMinimize.ForeColor = LightOff
        lblMinimize.Font.Underline = False
    
    ElseIf LabelIndex = 3 Then
        lblExport.ForeColor = LightOff
        lblExport.Font.Underline = False
        lblImport.ForeColor = LightOff
        lblImport.Font.Underline = False
        lblAbout.ForeColor = LightOn
        lblAbout.Font.Underline = True
        lblAdd.ForeColor = LightOff
        lblAdd.Font.Underline = False
        lblRemove.ForeColor = LightOff
        lblRemove.Font.Underline = False
        lblSearch.ForeColor = LightOff
        lblSearch.Font.Underline = False
        lblMinimize.ForeColor = LightOff
        lblMinimize.Font.Underline = False
    
    ElseIf LabelIndex = 4 Then
        lblExport.ForeColor = LightOff
        lblExport.Font.Underline = False
        lblImport.ForeColor = LightOff
        lblImport.Font.Underline = False
        lblAbout.ForeColor = LightOff
        lblAbout.Font.Underline = False
        lblAdd.ForeColor = LightOn
        lblAdd.Font.Underline = True
        lblRemove.ForeColor = LightOff
        lblRemove.Font.Underline = False
        lblSearch.ForeColor = LightOff
        lblSearch.Font.Underline = False
        lblMinimize.ForeColor = LightOff
        lblMinimize.Font.Underline = False
    
    ElseIf LabelIndex = 5 Then
        lblExport.ForeColor = LightOff
        lblExport.Font.Underline = False
        lblImport.ForeColor = LightOff
        lblImport.Font.Underline = False
        lblAbout.ForeColor = LightOff
        lblAbout.Font.Underline = False
        lblAdd.ForeColor = LightOff
        lblAdd.Font.Underline = False
        lblRemove.ForeColor = LightOn
        lblRemove.Font.Underline = True
        lblSearch.ForeColor = LightOff
        lblSearch.Font.Underline = False
        lblMinimize.ForeColor = LightOff
        lblMinimize.Font.Underline = False

    ElseIf LabelIndex = 6 Then
        lblExport.ForeColor = LightOff
        lblExport.Font.Underline = False
        lblImport.ForeColor = LightOff
        lblImport.Font.Underline = False
        lblAbout.ForeColor = LightOff
        lblAbout.Font.Underline = False
        lblAdd.ForeColor = LightOff
        lblAdd.Font.Underline = False
        lblRemove.ForeColor = LightOff
        lblRemove.Font.Underline = False
        lblSearch.ForeColor = LightOn
        lblSearch.Font.Underline = True
        lblMinimize.ForeColor = LightOff
        lblMinimize.Font.Underline = False
    
    ElseIf LabelIndex = 7 Then
        lblExport.ForeColor = LightOff
        lblExport.Font.Underline = False
        lblImport.ForeColor = LightOff
        lblImport.Font.Underline = False
        lblAbout.ForeColor = LightOff
        lblAbout.Font.Underline = False
        lblAdd.ForeColor = LightOff
        lblAdd.Font.Underline = False
        lblRemove.ForeColor = LightOff
        lblRemove.Font.Underline = False
        lblSearch.ForeColor = LightOff
        lblSearch.Font.Underline = False
        lblMinimize.ForeColor = LightOn
        lblMinimize.Font.Underline = True
        
    End If
    
End Sub

Private Sub mnuAddRecord_Click()
    
    Call UpdateRecord
    Call LogAction(LogWord(6))
    Dim outputKeyName As String
        
        ' Generate a random Key for storing the new record
GenerateOutputKeyName:
        outputKeyName = RandomPinString(4)
        
        ' Make sure generated key is not being used before
        For currentChr = 1 To Len(UserIndex) Step 8
            tmpString = Mid$(UserIndex, currentChr, 8)
            If tmpString = outputKeyName Then GoTo GenerateOutputKeyName
        Next
        
        If Len(ItemIndex) > 0 Then
            For currentChr = 1 To Len(ItemIndex) Step 8
                tmpString = Mid$(ItemIndex, currentChr, 8)
                If tmpString = outputKeyName Then GoTo GenerateOutputKeyName
            Next
        End If
        
    If ItemCount = 0 Then
        txtDescription.Enabled = True
        txtServer.Enabled = True
        txtUserName.Enabled = True
        txtPassword.Enabled = True
        txtNotes.Enabled = True
        mnuRemoveRecord.Enabled = True
        mnuSearch.Enabled = True
        imgRemove.Visible = True
        lblRemove.Visible = True
        imgSearch.Visible = True
        lblSearch.Visible = True
        
    End If
    
    ItemCount = ItemCount + 1
    ItemIndex = ItemIndex & outputKeyName
    tmpString = "New Recordƒƒƒƒƒ"
    ' Save in Registry
    SaveSetting MainTitle, UserRegSection, outputKeyName, crypt(tmpString, UserKeyword & Mid$(outputKeyName, 3, 1) & Mid$(outputKeyName, 5, 1))
    SaveSetting MainTitle, UserRegSection, "Item", crypt(ItemIndex, UserKeyword)
    
    tDescription = "New Record"
    tServer = ""
    tUserName = ""
    tPassword = ""
    tNotes = ""
    
    txtDescription.Text = "New Record"
    txtServer.Text = ""
    txtUserName.Text = ""
    txtPassword.Text = ""
    txtNotes.Text = ""
    currentKey = outputKeyName
    ' Add Record to ListView Control
    Set MainList = lstItem.ListItems.Add(, outputKeyName, "New Record", , 1)
    lstItem.ListItems(currentKey).Selected = True
    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription.Text)
    txtDescription.SetFocus

End Sub

Private Sub mnuExit_Click()
    Call UpdateRecord
    LogAction (LogWord(2))
    On Error Resume Next
    Close #2
    End
    
End Sub

Private Sub mnuExport_Click()
    
    Call UpdateRecord
    
    If ItemCount = 0 Then
        MsgBox "Nothing to export!", 48, MainTitle
        Exit Sub
    End If
    
    dlg.FileName = ""
    dlg.DialogTitle = "Export Records to File"
    dlg.Filter = "Password Gaurd Data Files (*.pgd)|*.pgd|All Files (*.*)|*.*"
    dlg.ShowSave
    If dlg.FileName = "" Then Exit Sub
    
    If FileExists(dlg.FileName + "") Then
        Title = MainTitle
        Msg = "The file " & LCase$(dlg.FileName) & " already exists." & vbCrLf & "Do you want to replace it?"
        DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

        Response = MsgBox(Msg, DgDef, Title)
        If Response = IDYES Then
            Kill dlg.FileName
        Else
            Exit Sub
        End If
    End If
    
    On Error GoTo ErrorFileExport
    Open dlg.FileName For Output As #3
    Close #3
    Keyword = ObtainFilePassword("Export", dlg.FileName)
    If Keyword = "" Then Exit Sub
    Call ExportData(dlg.FileName, Keyword)
    Call LogAction(LogWord(10), UCase$(dlg.FileName))
    Exit Sub
    
ErrorFileExport:
    Screen.MousePointer = 0
    MsgBox "Invalid file's path:" & vbCrLf & dlg.FileName, 48, MainTitle
    Exit Sub
    
    
End Sub

Private Sub mnuImport_Click()
    
    Dim InRecord1 As DataRecord
    Dim RecoverR1 As Boolean
    
    Call UpdateRecord
    If ItemCount >= 1 Then RecoverR1 = False Else RecoverR1 = True
    
    dlg.FileName = ""
    dlg.Filter = "Password Gaurd Data Files (*.pgd)|*.pgd|All Files (*.*)|*.*"
    dlg.DialogTitle = "Import Records from File"
    dlg.ShowOpen
    If dlg.FileName = "" Then Exit Sub
    
    If Not FileExists(dlg.FileName + "") Then
        MsgBox "Can not open file " & LCase$(dlg.FileName) & ". Please make sure that file exists and not being used by another application.", 48, "Error openning file"
        Exit Sub
    End If
    Keyword = ObtainFilePassword("Import", dlg.FileName)
    If Keyword = "" Then Exit Sub
    Call ImportData(dlg.FileName, Keyword)
    Call LogAction(LogWord(9), UCase$(dlg.FileName))
    
    If RecoverR1 = True Then
        currentKey = lstItem.ListItems(1).Key
        lstItem.ListItems(1).Selected = True
        InRecord1 = ReadRecord(currentKey)
        
        tDescription = InRecord1.Description
        tServer = InRecord1.Server
        tUserName = InRecord1.UserName
        tPassword = InRecord1.Password
        tNotes = InRecord1.Notes
        
        txtDescription.Text = InRecord1.Description
        txtServer.Text = InRecord1.Server
        txtUserName.Text = InRecord1.UserName
        txtPassword.Text = InRecord1.Password
        txtNotes.Text = InRecord1.Notes
        
        txtDescription.Enabled = True
        txtServer.Enabled = True
        txtUserName.Enabled = True
        txtPassword.Enabled = True
        txtNotes.Enabled = True
        mnuRemoveRecord.Enabled = True
        mnuSearch.Enabled = True
        imgRemove.Visible = True
        lblRemove.Visible = True
        imgSearch.Visible = True
        lblSearch.Visible = True

    End If
    
End Sub

Private Sub mnuLogout_Click()
    Call UpdateRecord
    Title = MainTitle
    Msg = "Are you sure you would like to Logout?"
    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

    Response = MsgBox(Msg, DgDef, Title)
    If Response = IDYES Then
    Else
        Exit Sub
    End If
    
    Call LogAction(LogWord(2))
    frmLogIn.lstUserID.Clear
    Dim mndex As String
    Dim mserIDIndex As String
    Dim mptionsIndex As String
    Dim mptions As String

    For currentIndex = 1 To Len(Index) Step 6
        UserCount = UserCount + 1
        UserRegSection = Mid$(Index, currentIndex, 6)
        regString = GetSetting(MainTitle, UserRegSection, "Index")
        mIndex = decrypt(regString, Key2 & UserRegSection)
        mOptionsIndex = Mid$(mIndex, 17, 8)
        regString = GetSetting(MainTitle, UserRegSection, mOptionsIndex)
        mOptions = decrypt(regString, Key1 & Mid$(mOptionsIndex, 3, 1) & Mid$(mOptionsIndex, 5, 1))
        
        If Left$(mOptions, 1) = "1" Then
            mUserIndex = Left$(mIndex, 8)
            regString = GetSetting(MainTitle, UserRegSection, mUserIndex)
            frmLogIn.lstUserID.AddItem decrypt(regString, Left$(UserRegSection, 2) & Right$(mUserIndex, 2) & Key1)
        End If
    Next

    Call LogAction("User logged out at " & GetTime & ".")
    frmMain.Hide
    frmLogIn.lstUserID.Text = ""
    frmLogIn.txtMasterPassword.Text = ""
    frmLogIn.Show
    
End Sub

Private Sub mnuRemoveRecord_Click()
    
    Call UpdateRecord
    
    If frmUserID.chkRemove.Value = 0 Then GoTo RemoveItem
    
    Title = "Remove Record"
    Msg = "Are you sure you want to remove the record " & lstItem.SelectedItem.Text & " ?"
    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

    Response = MsgBox(Msg, DgDef, Title)
    If Response = IDYES Then
    Else
        Exit Sub
    End If

RemoveItem:
    Call LogAction(LogWord(8), lstItem.SelectedItem.Text)
    Dim Key2Remove As String
    Dim outputItemIndex As String
    Dim tmpBackRec As DataRecord
    
    Key2Remove = lstItem.SelectedItem.Key
    outputItemIndex = ""
    For currentChr = 1 To Len(ItemIndex) Step 8
        tmpString = Mid$(ItemIndex, currentChr, 8)
        If Not tmpString = Key2Remove Then outputItemIndex = outputItemIndex & tmpString
    Next
    ItemIndex = outputItemIndex
    SaveSetting MainTitle, UserRegSection, "Item", crypt(ItemIndex, UserKeyword)
    lstItem.ListItems.Remove lstItem.SelectedItem.Index         ' Remove item from ListView control
    DeleteSetting MainTitle, UserRegSection, Key2Remove     ' Remove record's key from Registry
    ItemCount = ItemCount - 1
    
    If ItemCount = 0 Then
        mnuRemoveRecord.Enabled = False
        mnuSearch.Enabled = False
        imgRemove.Visible = False
        lblRemove.Visible = False
        imgSearch.Visible = False
        lblSearch.Visible = False
        tDescription = ""
        tServer = ""
        tUserName = ""
        tPassword = ""
        tNotes = ""
        
        txtDescription.Text = ""
        txtServer.Text = ""
        txtUserName.Text = ""
        txtPassword.Text = ""
        txtNotes.Text = ""
        
        txtDescription.Enabled = False
        txtServer.Enabled = False
        txtUserName.Enabled = False
        txtPassword.Enabled = False
        txtNotes.Enabled = False

        Exit Sub
    End If
    
    currentKey = lstItem.SelectedItem.Key
    lstItem.ListItems(currentKey).Selected = True
    tmpBackRec = ReadRecord(currentKey)
    txtDescription.Text = tmpBackRec.Description
    txtServer.Text = tmpBackRec.Server
    txtUserName.Text = tmpBackRec.UserName
    txtPassword.Text = tmpBackRec.Password
    txtNotes.Text = tmpBackRec.Notes
    
End Sub

Private Sub mnuSearch_Click()
        
    Call UpdateRecord
    frmSearch.chkDescription.Value = GetSetting(MainTitle, UserRegSection, "searchDescription", "1")
    frmSearch.chkServer.Value = GetSetting(MainTitle, UserRegSection, "searchServer", "0")
    frmSearch.chkUserName.Value = GetSetting(MainTitle, UserRegSection, "searchUserName", "0")
    frmSearch.chkPassword.Value = GetSetting(MainTitle, UserRegSection, "searchPassword", "0")
    frmSearch.chkNotes.Value = GetSetting(MainTitle, UserRegSection, "searchNotes", "0")

    regString = GetSetting(MainTitle, UserRegSection, "searchMode", "1")
    If regString = "1" Then frmSearch.optAny.Value = True
    If regString = "2" Then frmSearch.optAll.Value = True
    frmSearch.txtDescription.Text = ""
    frmSearch.txtServer.Text = ""
    frmSearch.txtUserName.Text = ""
    frmSearch.txtPassword.Text = ""
    frmSearch.txtNotes.Text = ""

    frmSearch.chkDescription_Click
    frmSearch.chkServer_Click
    frmSearch.chkUserName_Click
    frmSearch.chkPassword_Click
    frmSearch.chkNotes_Click
    
    If frmUserID.chkPassword.Value = 0 Then frmSearch.txtPassword.PasswordChar = "" Else frmSearch.txtPassword.PasswordChar = "*"
    frmSearch.Show 1
    If frmSearch.cmdSearch.Tag = "0" Then Exit Sub
    Screen.MousePointer = 11
    lblStatus.Caption = "Searching..."
    lblStatus.Refresh
    If frmSearch.optAny.Value = True Then Call Search(1, frmSearch.txtDescription.Text, frmSearch.txtServer.Text, frmSearch.txtUserName.Text, frmSearch.txtPassword.Text, frmSearch.txtNotes.Text)
    If frmSearch.optAll.Value = True Then Call Search(2, frmSearch.txtDescription.Text, frmSearch.txtServer.Text, frmSearch.txtUserName.Text, frmSearch.txtPassword.Text, frmSearch.txtNotes.Text)
    lblStatus.Caption = ""
    Screen.MousePointer = 0
    
End Sub

Private Sub mnuUserIDSettings_Click()
    
    Call UpdateRecord
    frmUserID.Show 1
    
    If Not frmUserID.txtUserID.Text = "" Then Exit Sub
    Unload frmUserID
    frmLogIn.lstUserID.Clear
    
    If Index = "/NEWRUN/" Then
        On Error Resume Next
        DeleteSetting MainTitle, "Settings", "Index"
        GoTo ErrHandler01
    End If
    
    Dim xIndex As String
    Dim xUserIDIndex As String
    Dim xOptionsIndex As String
    Dim xOptions As String
    
    For currentIndex = 1 To Len(Index) Step 6
        UserCount = UserCount + 1
        UserRegSection = Mid$(Index, currentIndex, 6)
        regString = GetSetting(MainTitle, UserRegSection, "Index")
        xIndex = decrypt(regString, Key2 & UserRegSection)
        xOptionsIndex = Mid$(xIndex, 17, 8)
        regString = GetSetting(MainTitle, UserRegSection, xOptionsIndex)
        xOptions = decrypt(regString, Key1 & Mid$(xOptionsIndex, 3, 1) & Mid$(xOptionsIndex, 5, 1))
        
        If Left$(xOptions, 1) = "1" Then
            xUserIndex = Left$(xIndex, 8)
            regString = GetSetting(MainTitle, UserRegSection, xUserIndex)
            frmLogIn.lstUserID.AddItem decrypt(regString, Left$(UserRegSection, 2) & Right$(xUserIndex, 2) & Key1)
        End If
    Next

ErrHandler01:
    frmMain.Hide
    frmLogIn.Show
    

End Sub

Private Sub mnuViewSearchResult_Click()
    frmSearchResults.Show
    
End Sub

Private Sub tmrPos_Timer()
    Dim CurHwnd As Long
    
    Call GetCursorPos(CursorPosition)
    CurHwnd = WindowFromPoint(CursorPosition.X, CursorPosition.Y)
    If Not CurHwnd = frmMain.hwnd Then Call LightLabel(0)
    
End Sub

Private Sub txtDescription_Change()
    If Not Trim(txtDescription.Text) = tDescription Then Saved = False
    
End Sub

Private Sub txtDescription_LostFocus()
    If Trim(txtDescription.Text) = "" Then
        Beep
        txtDescription.Text = tDescription
        Exit Sub
    End If
    
    If Saved = False Then
        Call SaveRecord(currentKey)
    End If
    lstItem.ListItems(currentKey).Text = Trim(txtDescription.Text)
        
End Sub

Private Sub txtDescription_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(0)

End Sub

Private Sub txtNotes_Change()
    If Not Trim(txtNotes.Text) = tNotes Then Saved = False

End Sub

Private Sub txtNotes_LostFocus()
    If Saved = False Then
        Call SaveRecord(currentKey)
    End If
    
End Sub


Private Sub txtNotes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(0)

End Sub

Private Sub txtPassword_Change()
    If Not Trim(txtPassword.Text) = tPassword Then Saved = False

End Sub

Private Sub txtPassword_LostFocus()
    If Saved = False Then
        Call SaveRecord(currentKey)
    End If
    

End Sub


Private Sub txtPassword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(0)

End Sub

Private Sub txtServer_Change()
    If Not Trim(txtServer.Text) = tServer Then Saved = False
    
End Sub

Private Sub txtServer_LostFocus()
    If Saved = False Then
        Call SaveRecord(currentKey)
    End If
    
End Sub


Private Sub txtServer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LightLabel(0)

End Sub

Private Sub txtUserName_Change()
    If Not Trim(txtUserName.Text) = tUserName Then Saved = False

End Sub

Private Sub txtUserName_LostFocus()
    If Saved = False Then
        Call SaveRecord(currentKey)
    End If
    
End Sub

Public Sub SaveRecord(RecordKey As String)
       
    tDescription = Trim(txtDescription.Text)
    tServer = Trim(txtServer.Text)
    tUserName = Trim(txtUserName.Text)
    tPassword = Trim(txtPassword.Text)
    tNotes = Trim(txtNotes.Text)
    
    tmpString = tDescription & sDivide & tServer & sDivide & tUserName & sDivide & tPassword & sDivide & tNotes & sDivide
    SaveSetting MainTitle, UserRegSection, RecordKey, crypt(tmpString, UserKeyword & Mid$(currentKey, 3, 1) & Mid$(currentKey, 5, 1))
       
End Sub


Public Sub ExportData(FileName As String, FileKeyword As String)
    
    Screen.MousePointer = 11
    Dim curIndexExport As Long
    Dim curKeyExport
    Dim curDataExport As String
    
    Open FileName For Output As #3
    Print #3, crypt(UserRegSection, FileKeyword) & crypt(FileTitle, Left$(UserRegSection, 1) & Mid$(UserRegSection, 3, 1) & FileKeyword & Mid$(UserRegSection, 5, 1))
    
    For curIndexExport = 1 To Len(ItemIndex) Step 8
        curKeyExport = Mid$(ItemIndex, curIndexExport, 8)
        regString = GetSetting(MainTitle, UserRegSection, curKeyExport)
        curDataExport = decrypt(regString, UserKeyword & Mid$(curKeyExport, 3, 1) & Mid$(curKeyExport, 5, 1))
        Print #3, crypt(curDataExport, FileKeyword)
    Next
    
    Close #3
    Screen.MousePointer = 0
End Sub

Public Sub ImportData(FileName As String, FileKeyword As String)
    
    Dim tmpUserRegSection As String
    Dim tmpFileRec As String
    Dim tmpImportRecord As DataRecord
    
    Open FileName For Input As #3
    If LOF(3) = 0 Then
        MsgBox "Sorry, unable to read from source file. Source file contains 0 byte."
        Close #3
        Exit Sub
    End If
    
    Line Input #3, tmpString
    tmpUserRegSection = decrypt(Left$(tmpString, 12), FileKeyword)
    tmpFileRec = decrypt(Mid$(tmpString, 13), Left$(tmpUserRegSection, 1) & Mid$(tmpUserRegSection, 3, 1) & FileKeyword & Mid$(tmpUserRegSection, 5, 1))
    
    If Not tmpFileRec = FileTitle Then
        MsgBox "Sorry, unable to read from source file. Please make sure that you've entered the correct Keyword.", 48, "Error reading file"
        Close #3
        Exit Sub
    End If
    
    Do Until EOF(3)
        Line Input #3, tmpString
        tmpString2 = decrypt(tmpString, FileKeyword)
NewImportKey:
        tmpString3 = RandomPinString(4)
        
            For currentChr = 1 To Len(UserIndex) Step 8
                If Mid$(UserIndex, currentChr, 8) = tmpString3 Then GoTo NewImportKey
            Next
            For currentChr = 1 To Len(ItemIndex) Step 8
                If Mid$(ItemIndex, currentChr, 8) = tmpString3 Then GoTo NewImportKey
            Next
        SaveSetting MainTitle, UserRegSection, tmpString3, crypt(tmpString2, UserKeyword & Mid$(tmpString3, 3, 1) & Mid$(tmpString3, 5, 1))
        tmpImportRecord = ReadRecord(tmpString3)
        Set MainList = lstItem.ListItems.Add(, tmpString3, tmpImportRecord.Description, , 1)
        ItemCount = ItemCount + 1
        ItemIndex = ItemIndex & tmpString3
    Loop
    Close #3
    SaveSetting MainTitle, UserRegSection, "Item", crypt(ItemIndex, UserKeyword)
    
End Sub

Private Sub txtUserName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LightLabel(0)

End Sub

Public Sub UpdateRecord()
' This Sub handles the AutoSave* feature
        
        If ItemCount <= 0 Or Saved = True Then Exit Sub
        Call LogAction(LogWord(7), lstItem.ListItems(currentKey).Text)
        
        If Not tDescription = Trim(txtDescription.Text) Then
            If Trim(txtDescription.Text) = "" Then txtDescription.Text = tDescription: Saved = True: Beep: Exit Sub Else lstItem.ListItems(currentKey).Text = Trim(txtDescription.Text)
        End If
        Call SaveRecord(currentKey)
        Saved = True

End Sub
