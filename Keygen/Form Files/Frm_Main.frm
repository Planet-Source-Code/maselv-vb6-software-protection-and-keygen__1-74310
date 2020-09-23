VERSION 5.00
Begin VB.Form Frm_Main 
   Caption         =   "App Title : Software KeyGen"
   ClientHeight    =   8010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11895
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   11895
   StartUpPosition =   1  'CenterOwner
   Begin SoftwareKeyGen.AutoSizer AutoSizer 
      Left            =   6840
      Top             =   4920
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.Frame Fra_Main 
      BackColor       =   &H00FFFFFF&
      Height          =   4800
      Left            =   2880
      TabIndex        =   2
      Tag             =   "AutoSizer:WH"
      Top             =   2175
      Width           =   8895
      Begin VB.Image ImgDeveloper 
         Height          =   1245
         Left            =   120
         Picture         =   "Frm_Main.frx":08CA
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   960
      End
      Begin VB.Label lblDeveloper 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Developer details"
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
         Left            =   1200
         TabIndex        =   41
         Top             =   3720
         Width           =   1245
      End
      Begin VB.Label lblSchoolInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please don't forget to vote for this code at planet source code"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Index           =   6
         Left            =   120
         TabIndex        =   40
         Tag             =   "AutoSizer:W"
         Top             =   1920
         Width           =   8730
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSchoolInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm_Main.frx":3FA5
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1380
         Index           =   5
         Left            =   120
         TabIndex        =   39
         Tag             =   "AutoSizer:W"
         Top             =   360
         Width           =   8640
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Fra_Photo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   6000
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Shape shpFraPhotoBorder 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   1455
         Left            =   0
         Top             =   0
         Width           =   1215
      End
      Begin VB.Image ImgDBPhoto 
         Height          =   1215
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   975
      End
      Begin VB.Image ImgVirtualPhoto 
         Height          =   135
         Left            =   240
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Timer TimerDateTime 
      Interval        =   100
      Left            =   5760
      Top             =   7320
   End
   Begin VB.Frame Fra_UserPhoto 
      BackColor       =   &H00CFE1E2&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      Begin VB.Image ImgUserPhoto 
         Height          =   975
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Timer TimerConnection 
      Interval        =   700
      Left            =   5400
      Top             =   3720
   End
   Begin VB.Shape ShpOutline 
      Height          =   1455
      Left            =   2880
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   0
      Left            =   240
      Picture         =   "Frm_Main.frx":40C7
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      Height          =   615
      Index           =   1
      Left            =   240
      Picture         =   "Frm_Main.frx":525C
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      Height          =   615
      Index           =   2
      Left            =   240
      Picture         =   "Frm_Main.frx":63F1
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      Height          =   615
      Index           =   3
      Left            =   240
      Picture         =   "Frm_Main.frx":7586
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      Height          =   615
      Index           =   4
      Left            =   240
      Picture         =   "Frm_Main.frx":871B
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      Height          =   615
      Index           =   5
      Left            =   240
      Picture         =   "Frm_Main.frx":98B0
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      Height          =   615
      Index           =   6
      Left            =   240
      Picture         =   "Frm_Main.frx":AA45
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Students"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   38
      Top             =   3000
      Width           =   630
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Parents/Guardians"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   37
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Classes"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   36
      Top             =   4200
      Width           =   540
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Dormitories"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   3
      Left            =   960
      TabIndex        =   35
      Top             =   4800
      Width           =   780
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student &Guardians"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   4
      Left            =   960
      TabIndex        =   34
      Top             =   5400
      Width           =   1320
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Classes"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   5
      Left            =   960
      TabIndex        =   33
      Top             =   6000
      Width           =   1140
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Software &Users"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   6
      Left            =   960
      TabIndex        =   32
      Top             =   6600
      Width           =   1080
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   0
      Left            =   480
      Picture         =   "Frm_Main.frx":BBDA
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   1
      Left            =   480
      Picture         =   "Frm_Main.frx":CAA4
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   2
      Left            =   480
      Picture         =   "Frm_Main.frx":CE2E
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   3
      Left            =   480
      Picture         =   "Frm_Main.frx":D1B8
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   4
      Left            =   480
      Picture         =   "Frm_Main.frx":E082
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   5
      Left            =   480
      Picture         =   "Frm_Main.frx":EA6C
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   6
      Left            =   480
      Picture         =   "Frm_Main.frx":F936
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label LblHour 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   870
      Left            =   240
      TabIndex        =   31
      ToolTipText     =   "The hour of the day"
      Top             =   1680
      Width           =   645
   End
   Begin VB.Label LblMinute 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   420
      Left            =   840
      TabIndex        =   30
      ToolTipText     =   "The current minute"
      Top             =   1770
      Width           =   315
   End
   Begin VB.Label LblSecond 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   885
      TabIndex        =   29
      ToolTipText     =   "The current second"
      Top             =   2070
      Width           =   225
   End
   Begin VB.Label LblAMPM 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PM"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   1110
      TabIndex        =   28
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblDateToday 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wed 23 Sep 2010"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   1365
      TabIndex        =   27
      Top             =   1860
      Width           =   1365
   End
   Begin VB.Label lblWeekDay 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Day 3 of 7"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1695
      TabIndex        =   26
      Top             =   2115
      Width           =   720
   End
   Begin VB.Label lblSchoolInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maselv High School"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   0
      Left            =   4440
      TabIndex        =   25
      Tag             =   "AutoSizer:W"
      Top             =   120
      Width           =   7410
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSchoolInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.O Box 1234, Nairobi, 00100, KENYA    Tel: (254) - 724 688 172   Fax: 020 123456"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   4440
      TabIndex        =   24
      Tag             =   "AutoSizer:W"
      Top             =   780
      Width           =   7290
   End
   Begin VB.Label lblSchoolInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail: info@maselvhigh.com"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   4440
      TabIndex        =   23
      Tag             =   "AutoSizer:W"
      Top             =   1035
      Width           =   7305
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H009CC1C5&
      X1              =   2880
      X2              =   2880
      Y1              =   1560
      Y2              =   120
   End
   Begin VB.Label LblLicenseTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ABCDE-FGHIJ-KLMNO-PQRST-UVWXY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   22
      Tag             =   "AutoSizer:Y"
      Top             =   7200
      Width           =   3060
   End
   Begin VB.Label LblLicenseTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Code:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Tag             =   "AutoSizer:Y"
      Top             =   7200
      Width           =   990
   End
   Begin VB.Label LblLicenseTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "09/03/1986"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   1440
      TabIndex        =   20
      Tag             =   "AutoSizer:Y"
      Top             =   7440
      Width           =   1020
   End
   Begin VB.Label LblLicenseTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Date:"
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
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Tag             =   "AutoSizer:Y"
      Top             =   7440
      Width           =   1020
   End
   Begin VB.Label LblLicenseTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   1200
      TabIndex        =   18
      Tag             =   "AutoSizer:Y"
      Top             =   7680
      Width           =   420
   End
   Begin VB.Label LblLicenseTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Key:"
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
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Tag             =   "AutoSizer:Y"
      Top             =   7680
      Width           =   885
   End
   Begin VB.Label LblLicenseTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max Users:"
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
      Index           =   6
      Left            =   1800
      TabIndex        =   16
      Tag             =   "AutoSizer:Y"
      Top             =   7680
      Width           =   930
   End
   Begin VB.Label LblLicenseTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   2880
      TabIndex        =   15
      Tag             =   "AutoSizer:Y"
      Top             =   7680
      Width           =   210
   End
   Begin VB.Label lblCurrentUser 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please don't forget to vote for this code at planet source code"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   7965
      TabIndex        =   14
      Tag             =   "AutoSizer:WY"
      Top             =   7200
      Width           =   3750
   End
   Begin VB.Label lblSchoolInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "As eagles we soar high"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   4
      Left            =   4440
      TabIndex        =   13
      Tag             =   "AutoSizer:W"
      Top             =   480
      Width           =   7425
   End
   Begin VB.Image imgStretcher 
      Height          =   15
      Index           =   1
      Left            =   240
      Picture         =   "Frm_Main.frx":FEC0
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:H"
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Label lblUserFullName 
      BackStyle       =   0  'Transparent
      Caption         =   "Full name"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1440
      TabIndex        =   12
      Tag             =   "W"
      Top             =   750
      Width           =   2235
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Tag             =   "W"
      Top             =   315
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Tag             =   "W"
      Top             =   120
      Width           =   510
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name:"
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
      Index           =   2
      Left            =   1440
      TabIndex        =   9
      Tag             =   "W"
      Top             =   555
      Width           =   855
   End
   Begin VB.Label lblUserGender 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Tag             =   "W"
      Top             =   960
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
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
      Index           =   3
      Left            =   1440
      TabIndex        =   7
      Tag             =   "W"
      Top             =   960
      Width           =   660
   End
   Begin VB.Label lblUserNationalID 
      BackStyle       =   0  'Transparent
      Caption         =   "National ID"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Tag             =   "W"
      Top             =   1395
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "National ID:"
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
      Index           =   4
      Left            =   1440
      TabIndex        =   5
      Tag             =   "W"
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblDeveloper 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by Masika Elvas elvasmasika@lexeme-kenya.com"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   0
      Left            =   7710
      TabIndex        =   4
      Tag             =   "AutoSizer:WY"
      Top             =   7680
      Width           =   3960
   End
   Begin VB.Label lblSchoolInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Website: http://www.maselvhigh.com"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   3
      Left            =   4440
      TabIndex        =   3
      Tag             =   "AutoSizer:W"
      Top             =   1305
      Width           =   7305
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgHeader 
      Height          =   1695
      Left            =   120
      Picture         =   "Frm_Main.frx":11055
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:W"
      Top             =   0
      Width           =   11895
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   5535
      Left            =   120
      Tag             =   "AutoSizer:WH"
      Top             =   1560
      Width           =   11775
   End
   Begin VB.Image imgFooter 
      Height          =   855
      Left            =   0
      Picture         =   "Frm_Main.frx":121EA
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:WY"
      Top             =   7080
      Width           =   11895
   End
   Begin VB.Menu mnuRegister 
      Caption         =   "&Register"
   End
   Begin VB.Menu mnuCancelRegistration 
      Caption         =   "&Cancel Registration"
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IsFrmLoadingComplete, vStartupComplete As Boolean

Private Sub Form_Activate()
    IsFrmLoadingComplete = True
    vStartupComplete = True
End Sub

Private Sub Form_Load()
    
    Me.Caption = App.Title & " : Switchboard"
    
    LblLicenseTo(&H0).Visible = VBA.IIf(VBA.DateDiff("d", VBA.Date, Licence.Expiry_Date) < 61 And Not vRegistered, True, False)
    
    mnuRegister.Visible = LblLicenseTo(&H0).Visible
    
    LblLicenseTo(&H1).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H2).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H3).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H4).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H5).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H6).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H7).Visible = LblLicenseTo(&H0).Visible
    
    LblLicenseTo(&H1).Caption = Licence.License_Code
    LblLicenseTo(&H5).Caption = Licence.Key
    LblLicenseTo(&H7).Caption = Licence.Max_Users
    
    Call TimerDateTime_Timer
    
    lblDeveloper(&H1).Caption = "DEVELOPER DETAILS:" & VBA.vbCrLf & _
                            "Name:       Masika .S. Elvas" & VBA.vbCrLf & _
                            "Address:   P.O Box 137, BUNGOMA 50200, KENYA" & VBA.vbCrLf & _
                            "Cell:          (254)724 688 172 / (254)751 041 184" & VBA.vbCrLf & _
                            "E-mail:      elvasmasika@lexeme-kenya.com"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If closure should not require User confirmation then confirm application exit
    If Not vSilentClosure Then Cancel = (MsgBox("Are you sure you want to Quit this Application?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbNo)
End Sub

Private Sub mnuCancelRegistration_Click()
    
    If MsgBox("Are you sure you want to revoke this Application's current Licence?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbNo Then Exit Sub
    VBA.DeleteSetting App.Title, "Copyright Protection"
    vSilentClosure = True: Unload Me
    
End Sub

Private Sub mnuRegister_Click()
    
    Me.Enabled = False
    
    vRegistering = True: vRegistered = True
    Frm_SoftwarePatent.Show , Me
    
    Do While vRegistering
        VBA.DoEvents: VBA.DoEvents
    Loop
    
    LblLicenseTo(&H0).Visible = VBA.IIf(VBA.DateDiff("d", VBA.Date, Licence.Expiry_Date) < 61 And Not vRegistered, True, False)
    
    mnuRegister.Visible = LblLicenseTo(&H0).Visible
    
    LblLicenseTo(&H1).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H2).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H3).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H4).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H5).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H6).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H7).Visible = LblLicenseTo(&H0).Visible
    
    LblLicenseTo(&H1).Caption = Licence.License_Code
    LblLicenseTo(&H5).Caption = Licence.Key
    LblLicenseTo(&H7).Caption = Licence.Max_Users
    
    Me.Enabled = True
    
End Sub

Private Sub TimerDateTime_Timer()
On Local Error GoTo Handle_TimerDateTime_Timer_Error
    
    'Display System Date & time
    LblSecond.Caption = VBA.Format$(VBA.Time$, "ss")
    LblMinute.Caption = VBA.Format$(VBA.Time$, "nn")
    LblHour.Caption = VBA.Format$(VBA.Time$, "HH")
    LblAMPM.Caption = VBA.Format$(VBA.Time$, "AMPM")
    lblDateToday.Tag = VBA.DateSerial(VBA.Year(VBA.Date), VBA.Month(VBA.Date), VBA.Day(VBA.Date))
    lblDateToday.Caption = VBA.Format$(lblDateToday.Tag, "ddd dd MMM yyyy")
    lblWeekDay.Caption = "Day " & VBA.Weekday(lblDateToday.Tag) & " of 7"
    
    'Format Expiry date & time appropriately
    LblLicenseTo(&H3).Caption = VBA.Format$(Licence.Expiry_Date, "ddd dd MMM yyyy hh:nn:ss AMPM") & VBA.IIf(VBA.DateDiff("d", VBA.Date, Licence.Expiry_Date) < &HB, " - Remaining " & VBA.DateDiff("d", VBA.Date, Licence.Expiry_Date) - &H1 & " day" & VBA.IIf(VBA.DateDiff("d", VBA.Date, Licence.Expiry_Date) - &H1 = &H1, "", "s") & ". This Software will automatically Shut down after expiring.", VBA.vbNullString)
    
    'If the expiry period is less than 31 days then display Serial Code in red
    LblLicenseTo(&H3).ForeColor = VBA.IIf(VBA.DateDiff("d", VBA.Date, Licence.Expiry_Date) + &H1 < 31, &HFF&, &H800000)
    
    'If the Form is loading or the Application is being registered then quit this sub procedure
    If Not IsFrmLoadingComplete Or Not vStartupComplete Or vRegistering Then Exit Sub
    
    'If the Software's Licence has expired then...
    If (VBA.DateDiff("d", VBA.Date, Licence.Expiry_Date) < &H1) And (Licence.Expiry_Date <> VBA.DateSerial(1986, 3, 9)) Then
        
1:
        
        'If the expiry date..
        Select Case VBA.DateDiff("d", VBA.Date, Licence.Expiry_Date)
            
            '..was the current day's date then...
            Case &H0: vBuffer(&H0) = "has expired today."
            
            '..was the previous day's date then...
            Case &H1: vBuffer(&H0) = "expired yesterday."
            
            '..was some day before then...
            Case Else: vBuffer(&H0) = "expired on " & VBA.Format$(Licence.Expiry_Date, "ddd dd MMM yyyy") & "."
            
        End Select
        
        vBuffer(&H0) = vBuffer(&H0) & " Please provide the following to the Software Administrator for renewal:-" & VBA.vbCrLf & _
                    "1. Software Name : " & App.Title & VBA.vbCrLf & _
                    "2. Device Serial : " & Licence.Device_Serial_No & VBA.vbCrLf & _
                    "3. Licence Key   : " & Licence.Key & VBA.vbCrLf & _
                    "4. " & LblLicenseTo(&H6).Caption & "  : " & Licence.Max_Users
                    
        'Warn User to renew the Licence. If the User accepts to renew Licence then...
        If MsgBox("The " & App.Title & " Software's Licence period " & vBuffer(&H0) & VBA.vbCrLf & " Do you want to Register Application?", vbCritical + vbYesNo + vbDefaultButton1, App.Title & " : Licence Expired") = vbYes Then
            
            Me.Enabled = False
            
            vRegistering = True: vRegistered = True
            Frm_SoftwarePatent.Show , Me
            
            Do While vRegistering
                VBA.DoEvents: VBA.DoEvents
            Loop
            
            LblLicenseTo(&H0).Visible = VBA.IIf(VBA.DateDiff("d", VBA.Date, Licence.Expiry_Date) < 61 And Not vRegistered, True, False)
            
            mnuRegister.Visible = LblLicenseTo(&H0).Visible
            
            LblLicenseTo(&H1).Visible = LblLicenseTo(&H0).Visible
            LblLicenseTo(&H2).Visible = LblLicenseTo(&H0).Visible
            LblLicenseTo(&H3).Visible = LblLicenseTo(&H0).Visible
            LblLicenseTo(&H4).Visible = LblLicenseTo(&H0).Visible
            LblLicenseTo(&H5).Visible = LblLicenseTo(&H0).Visible
            LblLicenseTo(&H6).Visible = LblLicenseTo(&H0).Visible
            LblLicenseTo(&H7).Visible = LblLicenseTo(&H0).Visible
            
            LblLicenseTo(&H1).Caption = Licence.License_Code
            LblLicenseTo(&H5).Caption = Licence.Key
            LblLicenseTo(&H7).Caption = Licence.Max_Users
            
            Me.Enabled = True
            
            'If the software licence has not expired then quit this procedure
            If Licence.Expiry_Date > VBA.Date Then Exit Sub
            
        End If
        
        'Warn User
        MsgBox "Software's Licence period not successfully verified. Application will not exit", vbExclamation, App.Title & " : Software Copyright"
        
        'Automatically Close the Software
        vSilentClosure = True: Unload Me 'Unload this Form from the memory
        
    End If 'Close respective IF..THEN block statement
    
Exit_TimerDateTime_Timer:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = vbDefault
    
    Exit Sub 'Quit this Procedure
    
Handle_TimerDateTime_Timer_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    MsgBox Err.Description, vbExclamation, App.Title & " : Date Time Error - " & Err.Number
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_TimerDateTime_Timer
    
End Sub
