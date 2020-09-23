VERSION 5.00
Begin VB.Form frmSearchResults 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Search Results"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3750
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
   ScaleHeight     =   1905
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstSearch 
      Height          =   1620
      Left            =   0
      TabIndex        =   1
      Top             =   280
      Width           =   3735
   End
   Begin VB.CheckBox chkOnTop 
      Caption         =   "Always on top"
      Height          =   255
      Left            =   40
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmSearchResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub chkOnTop_Click()

If chkOnTop.Value = 1 Then
               SetWindowPos frmSearchResults.hwnd, HWND_TOPMOST, frmSearchResults.Left / 15, _
                            frmSearchResults.Top / 15, frmSearchResults.Width / 15, _
                            frmSearchResults.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
Else
               SetWindowPos frmSearchResults.hwnd, HWND_NOTOPMOST, frmSearchResults.Left / 15, _
                            frmSearchResults.Top / 15, frmSearchResults.Width / 15, _
                            frmSearchResults.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End If

End Sub

Private Sub Form_Resize()

    If Me.WindowState = 1 Then Exit Sub    ' Window is minimized, exit
    On Error Resume Next
    Me.lstSearch.Height = Me.Height - 720
    On Error Resume Next
    Me.lstSearch.Width = Me.Width - 125
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    frmSearchResults.Hide
    
End Sub

Private Sub lstSearch_Click()
    
    Dim destResultKey As String
    Dim destArrayElement As Long
    Dim destResultRecord As DataRecord
    If lstSearch.ListCount = 0 Then Exit Sub
    destArrayElement = lstSearch.ListIndex
    destResultKey = SearchResult(destArrayElement)
    destResultRecord = ReadRecord(destResultKey)
    
    On Error GoTo ElementNOTFound
    Set frmMain.lstItem.SelectedItem = frmMain.lstItem.ListItems(destResultKey)
    tDescription = destResultRecord.Description
    tServer = destResultRecord.Server
    tUserName = destResultRecord.UserName
    tPassword = destResultRecord.Password
    tNotes = destResultRecord.Notes
    
    frmMain.txtDescription.Text = tDescription
    frmMain.txtServer.Text = tServer
    frmMain.txtUserName.Text = tUserName
    frmMain.txtPassword.Text = tPassword
    frmMain.txtNotes.Text = tNotes
    Exit Sub
    
ElementNOTFound:
    Exit Sub
End Sub
