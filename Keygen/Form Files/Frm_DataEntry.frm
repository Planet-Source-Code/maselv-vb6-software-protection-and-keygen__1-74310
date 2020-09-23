VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Frm_DataEntry 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Data Entry"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   Icon            =   "Frm_DataEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1530
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1530
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   -480
      Top             =   1455
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Fra_DataEntry 
      BackColor       =   &H00CFE1E2&
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   4215
      Begin VB.OptionButton OptEncryption 
         BackColor       =   &H00CFE1E2&
         Caption         =   "&Decrypt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   192
         Index           =   2
         Left            =   1920
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.OptionButton OptEncryption 
         BackColor       =   &H00CFE1E2&
         Caption         =   "&Encrypt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   192
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.OptionButton OptEncryption 
         BackColor       =   &H00CFE1E2&
         Caption         =   "&Normal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   192
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.TextBox txtEntry 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label LblInput 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry:"
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
         Height          =   210
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Label LblTrials 
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
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   960
   End
   Begin VB.Image ImgHeader 
      Height          =   255
      Left            =   0
      Picture         =   "Frm_DataEntry.frx":09EA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image ImgFooter 
      Height          =   495
      Left            =   0
      Picture         =   "Frm_DataEntry.frx":128A
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   4575
   End
End
Attribute VB_Name = "Frm_DataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'****************************************************************************************************
'*  Author:             Masika .S. Elvas                                                            *
'*  Gender:             Male                                                                        *
'*  Postal Address:     P.O Box 137, BUNGOMA 50200, KENYA                                           *
'*  Phone No:           (254) 724 688 172 / (254) 751 041 184                                       *
'*  E-mail Address:     maselv_e@yahoo.co.uk / masika_elvas@programmer.net / masika_elvas@live.com  *
'*  Location            BUNGOMA, KENYA                                                              *
'****************************************************************************************************

Option Explicit

Public strFilter$, strDefault$
Public mMax&, mMin&, mTrials&, mDialogAction
Public IsPassword, mIsNumeric, mIsDecimal, ConfirmClosure As Boolean

Private mCancelled As Boolean

Private Sub CmdBrowse_Click()
    
    With Dlg
        
        .Flags = &H4 'Hide Read-Only checkbox
        
        'Set Dialog to only show the specified Files.
        'If no file type has been specified then show all the files
        .Filter = VBA.IIf(VBA.Trim$(strFilter) = VBA.vbNullString, "All Files (*.*)|*.*", strFilter)
        
        '0 No Action.
        '1 Displays Open dialog box.
        '2 Displays Save As dialog box.
        '3 Displays Color dialog box.
        '4 Displays Font dialog box.
        '5 Displays Printer dialog box.
        If mDialogAction + &H1 = &H2 Then .FileName = txtEntry.Tag
        .Action = mDialogAction + &H1 'Display the CommonDialog control's defined dialog box.
        
        If mDialogAction + &H1 = &H2 Then txtEntry.Text = .FileName: Call cmdOK_Click
        
    End With
    
End Sub

Private Sub ImgFooter_DblClick()
    
    If LblInput.Caption <> "Enter Unlock Password:" Then Exit Sub
    
    Dim xStrKey$
    Dim vIndex&, xStrDate&
    
    xStrDate = VBA.Date 'VBA.DateAdd("d", 1, VBA.Date)
    
    xStrDate = VBA.Weekday(xStrDate) & VBA.Year(xStrDate) & VBA.Format$(VBA.Day(xStrDate), "00") & VBA.Format$(VBA.Month(xStrDate), "00")
    
    xStrKey = VBA.vbNullString
    
    For vIndex = &H1 To VBA.Len(xStrDate) Step &H3
        xStrKey = xStrKey & VBA.String$(&H3 - VBA.Len(VBA.Hex(VBA.Val(VBA.Mid(xStrDate, vIndex, &H3)))), "0") & VBA.Hex(VBA.Val(VBA.Mid(xStrDate, vIndex, &H3)))
    Next vIndex
    
    VB.Clipboard.Clear: VB.Clipboard.SetText xStrKey
    
End Sub

Private Sub cmdCancel_Click()
    vBuffer(&H0) = VBA.vbNullString: mCancelled = True: Unload Me
End Sub

Private Sub cmdOK_Click()
    
    cmdOK.SetFocus: mCancelled = False
    
    'If no entry has been made then...
    If txtEntry.Text = VBA.vbNullString Then
        
        'Inform User
        MsgBox "Please enter the requested data", vbInformation, App.Title & " : Blank Entry"
        txtEntry.SetFocus 'Move focus to Name textbox
        Exit Sub 'Quit this Saving Procedure
        
    End If 'End IF..THEN block function
    
    vBuffer(&H0) = txtEntry.Text
    
    If OptEncryption(&H1).Value And OptEncryption(&H1).Visible Then vBuffer(&H0) = SmartEncrypt(vBuffer(&H0)) Else If OptEncryption(&H2).Value And OptEncryption(&H2).Visible Then vBuffer(&H0) = SmartDecrypt(vBuffer(&H0), False)
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    Me.Caption = App.Title & " : " & VBA.Trim$(VBA.Replace(VBA.Replace(LblInput.Caption, "Enter", VBA.vbNullString), ":", VBA.vbNullString))
    
    vBuffer(&H0) = VBA.vbNullString
    
    txtEntry.Text = strDefault
    txtEntry.PasswordChar = VBA.IIf(IsPassword, "*", VBA.vbNullString)
    txtEntry.BackColor = VBA.IIf(txtEntry.Locked, &H71DFA3, &HC0FFFF)
    
    LblTrials.Caption = mTrials & " Attempts"
    LblTrials.Visible = (mTrials > &H0)
    
    OptEncryption(&H1).Visible = OptEncryption(&H0).Visible
    OptEncryption(&H2).Visible = OptEncryption(&H0).Visible
    
    txtEntry.SetFocus
    
End Sub

Private Sub Form_Load()
    Me.Caption = App.Title & " : Data Entry": ConfirmClosure = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mCancelled Then vBuffer(&H0) = "Cancelled": Cancel = (MsgBox("Are you sure you want to close this Form?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbNo): Exit Sub
    If ConfirmClosure Then If vBuffer(&H0) = VBA.vbNullString Then Cancel = (MsgBox("Are you sure you want to close this Form?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbNo)
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
    
    'If only numeric entries are required then Discard non-numeric entries
    If mIsNumeric Then
        
        If (((KeyAscii >= 32 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 126)) Or (Not mIsDecimal And KeyAscii = VBA.Asc("."))) And (KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack) Then KeyAscii = Empty
        KeyAscii = VBA.IIf(((KeyAscii >= 32 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 126)), VBA.IIf(Not mIsDecimal And KeyAscii = VBA.Asc("."), VBA.IIf(KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack, KeyAscii = Empty, KeyAscii), KeyAscii), KeyAscii)
        If KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then If Not VBA.IsNumeric(VBA.Left$(txtEntry.Text, txtEntry.SelStart) & VBA.Chr$(KeyAscii) & VBA.Right$(txtEntry.Text, VBA.Len(txtEntry.Text) - txtEntry.SelStart)) Then KeyAscii = Empty
        If Not mIsDecimal And KeyAscii = VBA.Asc(".") Then KeyAscii = Empty
                                                                                                                                                                    
    End If 'End IF..THEN block function
    
End Sub

Private Sub txtEntry_Validate(Cancel As Boolean)
    
    'If no entry has been made then quit this procedure
    If txtEntry.Text = VBA.vbNullString Then Exit Sub
    
    'If the entry is numeric then...
    If mIsNumeric Then
        
        'If then entered value exceeds the specified Maximum value
        If VBA.Val(VBA.Replace(txtEntry.Text, ",", VBA.vbNullString)) > mMax And mMax <> &H0 Then
            
            'Warn User to stick to the limits
            MsgBox "The entered value exceeds the Maximum value expected {" & mMax & "}", vbExclamation, App.Title & " : Invalid value"
            txtEntry.SetFocus: txtEntry.SelStart = &H0: txtEntry.SelLength = VBA.Len(txtEntry.Text)
            Cancel = True
            
        End If 'End IF..THEN block function
        
        'If then entered value exceeds the specified Minimum value
        If VBA.Val(VBA.Replace(txtEntry.Text, ",", VBA.vbNullString)) < mMin And mMin <> &H0 Then
            
            'Warn User to stick to the limits
            MsgBox "The entered value exceeds the Minimum value expected {" & mMin & "}", vbExclamation, App.Title & " : Invalid value"
            txtEntry.SetFocus: txtEntry.SelStart = &H0: txtEntry.SelLength = VBA.Len(txtEntry.Text)
            Cancel = True
            
        End If 'End IF..THEN block function
        
    End If 'End IF..THEN block function
    
End Sub
