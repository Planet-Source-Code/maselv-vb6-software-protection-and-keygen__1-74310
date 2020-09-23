VERSION 5.00
Begin VB.Form Frm_Login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2490
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4710
   Icon            =   "Frm_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1471.174
   ScaleMode       =   0  'User
   ScaleWidth      =   4422.435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1016
      Width           =   1095
   End
   Begin VB.CheckBox ChkRememberMe 
      BackColor       =   &H00CFE1E2&
      Caption         =   "&Remember Me"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1960
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Frame Fra_Login 
      BackColor       =   &H00CFE1E2&
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   4575
      Begin VB.TextBox TxtUserName 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtPassword 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label LblLogin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   840
      End
      Begin VB.Label LblLogin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   750
      End
   End
   Begin VB.Image ImgUser 
      Height          =   600
      Index           =   0
      Left            =   3840
      Picture         =   "Frm_Login.frx":0ECA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image ImgFooter 
      Height          =   495
      Left            =   0
      Picture         =   "Frm_Login.frx":1D94
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label LblTrials 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 Attempts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   3480
      TabIndex        =   11
      Top             =   2070
      Width           =   960
   End
   Begin VB.Image ImgUser 
      Height          =   240
      Index           =   1
      Left            =   4080
      Picture         =   "Frm_Login.frx":258A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   240
   End
   Begin VB.Label LblLogin 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter user name and password to connect to the server ..."
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label LblLogin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Information"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   1560
   End
   Begin VB.Label lblForgotPassword 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot your password?"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Click here if you have forgotten your account password"
      Top             =   2160
      Width           =   1560
   End
   Begin VB.Image ImgHeader 
      Height          =   855
      Left            =   0
      Picture         =   "Frm_Login.frx":2E54
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    LoginSucceeded = True
    Frm_Main.Show
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not LoginSucceeded Then Cancel = (MsgBox("Are you sure you want to Quit this Application?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbNo)
End Sub
