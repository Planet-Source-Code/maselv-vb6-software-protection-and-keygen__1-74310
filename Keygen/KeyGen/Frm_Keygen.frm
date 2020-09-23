VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Keygen 
   BackColor       =   &H00E2AD96&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Software Protection"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   Icon            =   "Frm_Keygen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fra_Login 
      BackColor       =   &H00E2AD96&
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   5655
      Begin VB.TextBox txtMaxUsers 
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
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   4440
         TabIndex        =   3
         Top             =   728
         Width           =   1095
      End
      Begin VB.TextBox txtSerialCode 
         BackColor       =   &H00E2AD96&
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox txtAppTitle 
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
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtPassword 
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
         ForeColor       =   &H00000040&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4440
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtExpiryDate 
         Height          =   300
         Left            =   4200
         TabIndex        =   6
         Top             =   1140
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   40591
      End
      Begin VB.TextBox txtSerialKey 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   1320
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox txtDeviceSerial 
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
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   1320
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtLicenseCode 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   3120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1140
         Width           =   372
      End
      Begin VB.Label lblSerial 
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
         Index           =   7
         Left            =   3480
         TabIndex        =   19
         Top             =   773
         Width           =   930
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Program Title:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pwd:"
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
         Index           =   5
         Left            =   3960
         TabIndex        =   17
         Top             =   285
         Width           =   390
      End
      Begin VB.Label lblSerial 
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
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1605
         Width           =   990
      End
      Begin VB.Label lblSerial 
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
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Device Serial:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "License Code:"
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
         Index           =   3
         Left            =   1920
         TabIndex        =   12
         Top             =   1185
         Width           =   1140
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry:"
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
         Index           =   4
         Left            =   3600
         TabIndex        =   11
         Top             =   1185
         Width           =   570
      End
   End
   Begin VB.CheckBox chkRegisterThisMachine 
      BackColor       =   &H00E2AD96&
      Caption         =   "Register this machine"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   2400
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox chkAlwaysOnTop 
      BackColor       =   &H00E2AD96&
      Caption         =   "Always On Top"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image ImgFooter 
      Height          =   735
      Left            =   0
      Picture         =   "Frm_Keygen.frx":08CA
      Stretch         =   -1  'True
      Top             =   2660
      Width           =   6015
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
      TabIndex        =   16
      Top             =   2910
      Width           =   960
   End
   Begin VB.Image ImgHeader 
      Height          =   495
      Left            =   0
      Picture         =   "Frm_Keygen.frx":164A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "Frm_Keygen"
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

'****************************************************************************************************
'*                                  COPY PROTECTION                                                 *
'****************************************************************************************************
'*  This is a software “lock” placed on the program by the developer to prevent the product from    *
'*  being copied and distributed without approval or authorization.                                 *
'****************************************************************************************************

Option Explicit
Option Compare Binary

Private Trials%
Private myArray() As String
Private IsLoading As Boolean
Private sBuffer$, Device_Serial_No$
Private vFso As New FileSystemObject
Private CopyrightVerified As Boolean

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Function GetWindowsDir() As String
    
    Dim Temp As String
    Dim Ret As Long
    
    Const MAX_LENGTH = 145

    Temp = VBA.String$(MAX_LENGTH, &H0)
    Ret = GetWindowsDirectory(Temp, MAX_LENGTH)
    Temp = VBA.Left$(Temp, Ret)
    GetWindowsDir = VBA.IIf(Temp <> VBA.vbNullString And VBA.Right$(Temp, &H1) <> "\", Temp & "\", Temp)
    
End Function

'------------------------------------------------------------------------------------------
Public Function SmartEncrypt(StringToEncrypt As String, Optional AlphaEncoding As Boolean = True) As String
On Error GoTo ErrorHandler
    
    Dim I&
    Dim Char$, strEncrypt$
    
    If StringToEncrypt = VBA.vbNullString Then Exit Function
    
    For I = &H1 To VBA.Len(StringToEncrypt) Step &H1
        Char = VBA.Asc(VBA.Mid(StringToEncrypt, I, &H1))
        SmartEncrypt = SmartEncrypt & VBA.Len(Char) & Char
    Next I
    
    If AlphaEncoding Then
    
        strEncrypt = SmartEncrypt
        SmartEncrypt = VBA.vbNullString
        
        For I = &H1 To VBA.Len(strEncrypt) Step &H1
            SmartEncrypt = SmartEncrypt & VBA.Chr(VBA.Mid(strEncrypt, I, &H1) + &H93)
        Next I
        
    End If
    
    Exit Function
    
ErrorHandler:
    
    SmartEncrypt = "Error encrypting string"
    
End Function

Public Function SmartDecrypt(StringToDecrypt As String, Optional AlphaDecoding As Boolean = True) As String
On Error GoTo ErrorHandler
    
    Dim I&
    Dim CharPos%
    Dim Char$, CharCode$, strEncrypt$
    
    If StringToDecrypt = VBA.vbNullString Then Exit Function
        
    If AlphaDecoding Then
    
        SmartDecrypt = StringToDecrypt
        
        For I = &H1 To VBA.Len(SmartDecrypt) Step &H1
            strEncrypt = strEncrypt & (VBA.Asc(VBA.Mid(SmartDecrypt, I, &H1)) - &H93)
        Next I
        
    End If
    
    SmartDecrypt = VBA.vbNullString
    
    If strEncrypt = VBA.vbNullString Then strEncrypt = StringToDecrypt
    
    Do While strEncrypt <> VBA.vbNullString
        
        CharPos = VBA.Left(strEncrypt, &H1)
        strEncrypt = VBA.Mid(strEncrypt, &H2)
        CharCode = VBA.Left(strEncrypt, CharPos)
        strEncrypt = VBA.Mid(strEncrypt, VBA.Len(CharCode) + &H1)
        SmartDecrypt = SmartDecrypt & VBA.Chr(CharCode)
                
    Loop
    
    Exit Function
    
ErrorHandler:
    
End Function

Private Function EncodeSerial(vDate As Date, vWeekday%, vSerialKey&, vUsers&) As String
On Local Error GoTo Handle_EncodeSerial_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim sIndex&
    Dim sBuffer$, sBuffer1$
    
    'Users 00: Day 00: Month 00: Serial 0000: Year 00
    
    VBA.Randomize 'Initializes the random-number generator.
    
    'Txt(&H0).Text = VBA.Format$(vUsers, "00") & VBA.Format$(VBA.Day(vDate), "00") & VBA.Format$(VBA.Month(VBA.Date), "00") & vSerialKey & VBA.Right(VBA.Year(VBA.Date), &H2)
    sBuffer1 = VBA.StrReverse(VBA.Format$(vUsers, "00")) & VBA.StrReverse(VBA.Format$(VBA.Day(vDate), "00")) & VBA.StrReverse(VBA.Format$(VBA.Month(vDate), "00")) & VBA.StrReverse(vSerialKey) & VBA.StrReverse(VBA.Right(VBA.Year(vDate), &H2))
    
    'For each character in the provided Serial Code
    For sIndex = &H1 To VBA.Len(sBuffer1) Step &H1
        
        'Returns an Integer representing the code for each character
        sBuffer = sBuffer & VBA.Asc(VBA.Mid$(sBuffer1, sIndex, &H1))
        
    Next sIndex 'Increment counter variable by value in the STEP option
    
    'Attach the System's weekday to the generated value
    sBuffer = VBA.StrReverse(sBuffer) & vWeekday
    
    sBuffer1 = VBA.vbNullString 'Initialize variable
    
    'For each character in the generated value
    For sIndex = &H1 To VBA.Len(sBuffer) Step &H1
        
        'Return a String containing the character associated with the each character code
        sBuffer1 = sBuffer1 & VBA.Chr$(65 + (VBA.Mid$(sBuffer, sIndex, &H1)))
        
    Next sIndex 'Increment counter variable by value in the STEP option
    
    'Reverse the character order of the specified string
    sBuffer1 = VBA.StrReverse(sBuffer1)
    
    sBuffer = VBA.vbNullString 'Initialize variable
    
    'For each character in the generated value
    For sIndex = &H1 To VBA.Len(sBuffer1) Step &H5
        
        'Split the value into 5 characters separated by -
        sBuffer = sBuffer & VBA.Mid$(sBuffer1, sIndex, &H5) & "-"
        
    Next sIndex 'Increment counter variable by value in the STEP option
    
    'Assign the final generated value to this Function
    EncodeSerial = VBA.Left(sBuffer, VBA.Len(sBuffer) - &H1)
    
'    VB.Clipboard.Clear: VB.Clipboard.SetText EncodeSerial
    
Exit_EncodeSerial:
    
    'Set Mouse pointer to indicate end of this process or operation
    Screen.MousePointer = vbDefault
    
    Exit Function 'Quit this Function
    
Handle_EncodeSerial_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    MsgBox Err.Description, vbExclamation, App.Title & " : Error Encoding - " & Err.Number
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_EncodeSerial
    
End Function

Private Function DecodeSerial(SerialCode$) As String
On Local Error GoTo Handle_DecodeSerial_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim sIndex&
    Dim sBuffer$, sBuffer1$
    Dim vWeekday&, vDay&, vMonth&, vYear&, vSerial&, vUsers&
    
    'Reverse the character order of the specified string and combine characters in the value by removing by -
    sBuffer = VBA.Replace(VBA.StrReverse(SerialCode), "-", VBA.vbNullString)
    
    'For each character in the specified value
    For sIndex = &H1 To VBA.Len(sBuffer) Step &H1
        
        'Return a String containing the character codes associated with the each character in the value
        sBuffer1 = sBuffer1 & (VBA.Asc((VBA.Mid$(sBuffer, sIndex, &H1))) - 65)
        
    Next sIndex 'Increment counter variable by value in the STEP option
    
    vWeekday = VBA.Right$(sBuffer1, &H1) 'Extract the System's Weekday from the regenerated Serial Code
    
    'Reverse the character order of the regenerated Serial Code without the System's Weekday
    sBuffer1 = VBA.StrReverse(VBA.Left$(sBuffer1, VBA.Len(sBuffer1) - &H1))
    
    sBuffer = VBA.vbNullString 'Initialize variable
    
    'For each character in the specified value
    For sIndex = &H1 To VBA.Len(sBuffer1) Step &H2
        
        'Return a String containing the character associated with the each double-character code
        sBuffer = sBuffer & VBA.Chr$(VBA.Mid$(sBuffer1, sIndex, &H2))
        
    Next sIndex 'Increment counter variable by value in the STEP option
    
    'Assign the final generated value to this Function
    DecodeSerial = VBA.StrReverse(VBA.Left(sBuffer, &H2)) & "|" & _
                    VBA.StrReverse(VBA.Mid$(sBuffer, &H3, &H2)) & "|" & _
                    VBA.StrReverse(VBA.Mid$(sBuffer, &H5, &H2)) & "|" & _
                    VBA.StrReverse(VBA.Mid$(sBuffer, &H7, &H4)) & "|" & _
                    VBA.StrReverse(VBA.Mid$(sBuffer, &HB, &H2)) & "|" & vWeekday
    
Exit_DecodeSerial:
    
    'Set Mouse pointer to indicate end of this process or operation
    Screen.MousePointer = vbDefault
    
    Exit Function 'Quit this Function
    
Handle_DecodeSerial_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    MsgBox Err.Description, vbExclamation, App.Title & " : Error Decoding - " & Err.Number
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_DecodeSerial
    
End Function

Private Sub chkAlwaysOnTop_Click()

    'Set the Form to be always on top of other open windows
    Call SetWindowPos(Me.hwnd, VBA.IIf(chkAlwaysOnTop.Value = vbChecked, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, FLAGS)
    
End Sub

Private Sub CmdCancel_Click()
    Unload Me 'Unload this Form from the Memory
End Sub

Private Sub cmdOK_Click()
On Local Error GoTo Handle_CmdOK_Click_Error
    
    'If an invalid password has been specified then...
    If txtPassword.Text <> VBA.String$(VBA.Weekday(VBA.Date), VBA.CStr(VBA.Weekday(VBA.Date))) Then
        
        'Warn User
        MsgBox "Invalid Password!!!", vbExclamation, App.Title & " : Operation Denied"
        txtPassword.SetFocus: txtPassword.SelStart = &H0: txtPassword.SelLength = VBA.Len(txtPassword.Text)
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim sMachineSerial$
    
    'If the device being unlocked is the one running this program then...
    If VBA.Left(VBA.StrReverse(txtDeviceSerial.Tag), VBA.Len(VBA.Replace(txtDeviceSerial.Text, "-", VBA.vbNullString))) = VBA.StrReverse(txtDeviceSerial.Tag) Then
        
        myArray = VBA.Split(SmartDecrypt(VBA.GetSetting(txtAppTitle.Text, "Copyright Protection", "Device Serial", VBA.vbNullString)), "-")
        
        'If the program is running for the first time on the Computer then...
        If UBound(myArray) < &H0 Then
            
            'Warn User
            MsgBox "Please run the Software first before using this program.", vbExclamation, App.Title & " : No Software Details"
            GoTo Exit_CmdOK_Click 'Quit this Procedure
            
        End If 'Close respective IF..THEN block statement
        
        'Check if a serial key previously been generated for the Computer
        txtSerialKey.Text = VBA.StrReverse(SmartDecrypt(myArray(&H0), False))
        
    End If 'Close respective IF..THEN block statement
    
    txtSerialCode.Text = EncodeSerial(dtExpiryDate.Value, VBA.CInt(VBA.Val(txtLicenseCode.Text)), VBA.CLng(VBA.Val(txtSerialKey.Text)), VBA.CLng(VBA.Val(txtMaxUsers.Text)))
    
    'If the device being unlocked is not the one running this program then Quit this Procedure
    If chkRegisterThisMachine.Value = vbUnchecked Or VBA.Left(VBA.StrReverse(txtDeviceSerial.Tag), VBA.Len(VBA.Replace(txtDeviceSerial.Text, "-", VBA.vbNullString))) <> VBA.StrReverse(VBA.Replace(txtDeviceSerial.Text, "-", VBA.vbNullString)) Then GoTo Exit_CmdOK_Click
    
    'If the Software's Serial Code is valid then...
    If txtSerialCode.Tag = txtSerialCode.Text Then
        
        'Inform User
        MsgBox "The Software had already been unlocked.", vbInformation, txtAppTitle.Text & " : Unlock"
        
    Else 'If the Software's Serial Code is invalid then...
        
        VBA.SaveSetting txtAppTitle.Text, "Copyright Protection", "License Code", SmartEncrypt(VBA.StrReverse(VBA.Format$(txtLicenseCode.Text, "00")), False)
        VBA.SaveSetting txtAppTitle.Text, "Copyright Protection", "License Encrypted", SmartEncrypt(txtSerialCode.Text & "|" & VBA.Replace(txtDeviceSerial.Text, "-", VBA.vbNullString))
        VBA.SaveSetting txtAppTitle.Text, "Copyright Protection", "Device Serial", SmartEncrypt(SmartEncrypt(VBA.StrReverse(txtSerialKey.Text), False) & "-" & VBA.StrReverse(VBA.Format$(VBA.Replace(txtDeviceSerial.Text, "-", VBA.vbNullString), "000000000000")))
        
        'Inform User
        MsgBox txtAppTitle.Text & " Software successfully unlocked. Please re-run the Software.", vbInformation, App.Title & " : " & txtAppTitle.Text & " Program Unlocked!!"
        
    End If 'Close respective IF..THEN block statement
    
Exit_CmdOK_Click:
    
    'Set Mouse pointer to indicate end of this process or operation
    Screen.MousePointer = vbDefault
    
    Exit Sub 'Quit this Procedure
    
Handle_CmdOK_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    MsgBox Err.Description, vbExclamation, App.Title & " : Error Activating Form - " & Err.Number
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_CmdOK_Click
    
End Sub

Private Sub Form_Load()
    
    Dim vDrive, vFso
    
    Me.Caption = App.Title & " v." & App.Major & "." & App.Minor & "." & App.Revision & " : Software Protection"
    Set vFso = CreateObject("Scripting.FileSystemObject")
    Set vDrive = vFso.GetDrive(vFso.GetDriveName(GetWindowsDir))
    txtDeviceSerial.Tag = VBA.Format$(VBA.Replace(vDrive.SerialNumber, "-", VBA.vbNullString), "000000000000")
    txtDeviceSerial.Text = vDrive.SerialNumber: Set vDrive = Nothing
    dtExpiryDate.Value = VBA.Date
    
    'Assign the serial key of the Computer
    txtAppTitle.Text = VBA.GetSetting(App.Title, "Settings", "Last Software Loaded", "")
    
End Sub

Private Sub ImgFooter_DblClick()
    
    Dim xStrKey$, xDate$
    Dim vIndex&, xStrDate&
    
    xDate = VBA.Date ' "31/01/2012"
    xStrDate = VBA.Weekday(xDate) & VBA.Year(xDate) & VBA.Format$(VBA.Day(xDate), "00") & VBA.Format$(VBA.Month(xDate), "00")
    
    xStrKey = VBA.vbNullString
    
    For vIndex = &H1 To VBA.Len(xStrDate) Step &H3
        xStrKey = xStrKey & VBA.String$(&H3 - VBA.Len(VBA.Hex(VBA.Val(VBA.Mid(xStrDate, vIndex, &H3)))), "0") & VBA.Hex(VBA.Val(VBA.Mid(xStrDate, vIndex, &H3)))
    Next vIndex
    
    VB.Clipboard.Clear: VB.Clipboard.SetText xStrKey
    
End Sub

Private Sub txtAppTitle_Validate(Cancel As Boolean)
    
    'Capitalize first Letter of each Title Word entered
    txtAppTitle.Text = VBA.StrConv(txtAppTitle.Text, vbProperCase)
    
CheckSettings:
    
    'If the device being unlocked is the one running this program then...
    If VBA.Left(VBA.StrReverse(txtDeviceSerial.Tag), VBA.Len(VBA.Replace(txtDeviceSerial.Text, "-", VBA.vbNullString))) = VBA.StrReverse(VBA.Replace(txtDeviceSerial.Text, "-", VBA.vbNullString)) And txtAppTitle.Text <> VBA.vbNullString Then
        
        myArray = VBA.Split(SmartDecrypt(VBA.GetSetting(txtAppTitle.Text, "Copyright Protection", "Device Serial", VBA.vbNullString)), "-")
        
        'If the program is running for the first time on the Computer then...
        If UBound(myArray) < &H0 Then
            
            'Warn User
            MsgBox "Please run the Software first before using this program.", vbExclamation, App.Title & " : No Software Details"
            Exit Sub 'Quit this Procedure
            
        End If 'Close respective IF..THEN block statement
        
        'Assign the serial key of the Computer
        VBA.SaveSetting App.Title, "Settings", "Last Software Loaded", txtAppTitle.Text
        
        'If the program is running for the first time on the Computer then...
        If UBound(myArray) < &H1 Then
            
            'Assign the serial key of the Computer
            VBA.SaveSetting txtAppTitle.Text, "Copyright Protection", "Device Serial", SmartEncrypt(SmartEncrypt(VBA.StrReverse(VBA.Int((1000 - 9999 + 1000) * VBA.Rnd + 9999)), False) & "-" & VBA.StrReverse(Device_Serial_No))
            
        Else
            
            If VBA.StrReverse(myArray(&H1)) <> txtDeviceSerial.Tag Then
                
                'Assign the serial key of the Computer
                VBA.SaveSetting txtAppTitle.Text, "Copyright Protection", "Device Serial", SmartEncrypt(SmartEncrypt(VBA.StrReverse(VBA.Int((1000 - 9999 + 1000) * VBA.Rnd + 9999)), False) & "-" & VBA.StrReverse(Device_Serial_No))
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        myArray = VBA.Split(SmartDecrypt(VBA.GetSetting(txtAppTitle.Text, "Copyright Protection", "Device Serial", VBA.vbNullString)), "-")
        
        txtLicenseCode.Text = VBA.StrReverse(SmartDecrypt(VBA.GetSetting(txtAppTitle.Text, "Copyright Protection", "License Code", VBA.vbNullString), False))
        
        txtSerialKey.Text = VBA.StrReverse(SmartDecrypt(myArray(&H0), False))
        
        Dim sStr$
        
        'Validate existing License Code if it exists
        sStr = SmartDecrypt(VBA.GetSetting(txtAppTitle.Text, "Copyright Protection", "License Encrypted", VBA.vbNullString))
        
        If sStr <> VBA.vbNullString Then
            
            'If the existing code is not up to to required number of characters then...
            If VBA.Len(VBA.Replace(VBA.Replace(sStr, "-", VBA.vbNullString), "|" & VBA.Replace(txtDeviceSerial.Text, "-", VBA.vbNullString), VBA.vbNullString)) < 25 Then GoTo FakeSerialCode
            
            txtSerialCode.Text = VBA.Replace(VBA.Replace(sStr, "|-", "|"), "|" & VBA.Replace(txtDeviceSerial.Text, "-", VBA.vbNullString), VBA.vbNullString)
            
            myArray = VBA.Split(DecodeSerial(txtSerialCode.Text), "|")
            
            If txtLicenseCode.Text = VBA.vbNullString Then GoTo FakeSerialCode
            If UBound(myArray) < &H5 Then GoTo FakeSerialCode
            
            If Not VBA.IsDate(VBA.DateSerial(myArray(&H4), myArray(&H2), myArray(&H1))) Then GoTo FakeSerialCode Else dtExpiryDate.Tag = VBA.DateSerial(myArray(&H4), myArray(&H2), myArray(&H1))
            If Not VBA.IsNumeric(myArray(&H0)) Then GoTo FakeSerialCode Else txtMaxUsers.Text = myArray(&H0) 'Max Users
            If Not VBA.IsNumeric(VBA.StrReverse(myArray(&H3))) Then GoTo FakeSerialCode
            
            If VBA.Val(myArray(&H5)) <> VBA.Val(txtLicenseCode.Text) Then GoTo FakeSerialCode
            If VBA.Val(myArray(&H3)) <> VBA.Val(txtSerialKey.Text) Then GoTo FakeSerialCode
            
            dtExpiryDate.Value = dtExpiryDate.Tag
            
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
    If txtLicenseCode.Text = VBA.vbNullString Then txtLicenseCode.Text = VBA.Format$(VBA.Weekday(VBA.Date), "00")
    
    Exit Sub
    
FakeSerialCode:
    
    VBA.DeleteSetting txtAppTitle.Text, "Copyright Protection"
    GoTo CheckSettings
    
End Sub

Private Sub txtDeviceSerial_Change()
    chkRegisterThisMachine.Visible = (VBA.Left(VBA.StrReverse(txtDeviceSerial.Tag), VBA.Len(VBA.Replace(txtDeviceSerial.Text, "-", VBA.vbNullString))) = VBA.StrReverse(VBA.Replace(txtDeviceSerial.Text, "-", VBA.vbNullString)))
End Sub

Private Sub txtSerialCode_GotFocus()
    
    'Highlight the entered Serial Code contents
    txtSerialCode.SetFocus 'Move focus to the specified control
    txtSerialCode.SelStart = &H0: txtSerialCode.SelLength = VBA.Len(txtSerialCode.Text)
    
End Sub

