VERSION 5.00
Begin VB.Form frmKeyword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Keyword"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
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
   ScaleHeight     =   1605
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Okay"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtKeyword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label lblAction 
      AutoSize        =   -1  'True
      Caption         =   "Action"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3600
      TabIndex        =   5
      Top             =   360
      Width           =   525
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "FilePath"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   660
   End
   Begin VB.Image img 
      Height          =   525
      Left            =   240
      Top             =   360
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please type the Password to            the file:"
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   3795
   End
End
Attribute VB_Name = "frmKeyword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    txtKeyword.Text = ""
    frmKeyword.Hide
    
End Sub

Private Sub cmdOK_Click()
    frmKeyword.Hide
    
End Sub

Private Sub txtKeyword_Change()
    
    If Len(Trim(txtKeyword.Text)) > 0 Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
    
End Sub
