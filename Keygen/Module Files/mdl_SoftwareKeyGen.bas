Attribute VB_Name = "mdl_SoftwareKeyGen"
Option Explicit

Private Const WM_COMMAND = &H111
Private Const MIN_ALL = &H1A3
Private Const MIN_ALL_UNDO = &H1A0

Public vRegistered As Boolean 'Denotes that the Software has been registered
Public vRegistering As Boolean 'Denotes that the Software is being registered
Public vSilentClosure As Boolean 'Denotes that the Form should close without User confirmation

Public Type SoftwareLicences
    
    License_Code As String
    License_Encrypted As String
    Expiry_Date As Date
    Max_Users As Long
    Key As String
    
    Device_Name As String
    Device_Account_Name As String
    Device_Serial_No As String
    
End Type

Public Licence As SoftwareLicences

Public vBuffer(&H1) As String

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Minimize all Open windows 'STATE=TRUE {Minimize All}, STATE=FALSE {Restore All}
Public Sub WindowsMinimizeAll(Optional State As Boolean = True)
On Error Resume Next
    
    Dim lngHwnd&
    
    lngHwnd = FindWindow("Shell_TrayWnd", VBA.vbNullString)
    Call PostMessage(lngHwnd, WM_COMMAND, VBA.IIf(State, MIN_ALL, MIN_ALL_UNDO), 0&)
  
End Sub

' This function creates an integer value based on the key
' provided.  The principle is simple.  The result is the
' absolute value of the difference between the averages of
' the odd and even characters.

Public Function CreateEncryptCode(Key As String) As Integer
    
    Dim Total(&H1 To &H2) As Integer
    Dim NbChars(&H1 To &H2) As Integer
    Dim vIndex, Index As Integer
    
    Total(&H1) = &H0: Total(&H2) = &H0
    NbChars(&H1) = &H0: NbChars(&H2) = &H0
    
    For vIndex = &H1 To VBA.LenB(Key) Step &H1
        
        Index = VBA.IIf(vIndex Mod &H2 = &H0, &H1, &H2) ' Characters in an even/odd position
        Total(Index) = Total(Index) + VBA.Asc(VBA.Mid(Key, vIndex, &H1))
        NbChars(Index) = NbChars(Index) + &H1
        
    Next vIndex
    
    ' A division by zero must be avoided.
    ' This will be the new value used for encryption
    ' Else If the key is less than 2 characters long, the code becomes 1
    CreateEncryptCode = VBA.IIf(NbChars(&H1) > &H0 And NbChars(&H2) > &H0, VBA.Abs((Total(&H1) / NbChars(&H1)) - (Total(&H2) / NbChars(&H2))), &H1)
    
End Function

' I prefer alternating between an addition and a subtraction
' to provide a more complex encryption method.  It is more
' difficult to crack due to the alternations and the
' encryption key.


' OrigStr : The original string value before encryption.
' Key     : The key used for encrypting/decrypting the string

Public Function EncryptStr(ByVal OrigStr As String, Optional Key As String = "Ketheline", Optional Decrypt As Boolean = False) As String
    
    'If no value has been supplied then quit this Function
    If VBA.LenB(VBA.Trim$(OrigStr)) = &H0 Then Exit Function
    
    Dim vIndex, EncCode As Integer
    
    ' First thing done is a calculation upon the encryption key
    ' to determine how the original string will be encrypted.
    EncCode = CreateEncryptCode(Key)
    EncryptStr = VBA.vbNullString
    
    ' Now the string will be changed according to the new encryption values
    For vIndex = &H1 To VBA.LenB(OrigStr) Step &H1
        EncryptStr = EncryptStr + VBA.IIf(Decrypt, VBA.IIf(vIndex Mod &H2 = &H0, VBA.Chr(VBA.Asc(VBA.Mid(OrigStr, vIndex, &H1)) - EncCode), VBA.Chr(VBA.Asc(VBA.Mid(OrigStr, vIndex, &H1)) + EncCode)), VBA.IIf(vIndex Mod &H2 = &H0, VBA.Chr(VBA.Asc(VBA.Mid(OrigStr, vIndex, &H1)) + EncCode), VBA.Chr(VBA.Asc(VBA.Mid(OrigStr, vIndex, &H1)) - EncCode)))
    Next vIndex
    
End Function

Public Function SmartDecrypt(StringToDecrypt As String, Optional AlphaDecoding As Boolean = True) As String
On Local Error GoTo ErrorHandler
    
    Dim i&
    Dim CharPos%
    Dim Char$, CharCode$, strEncrypt$
    
    If VBA.LenB(VBA.Trim$(StringToDecrypt)) = &H0 Then Exit Function
        
    If AlphaDecoding Then
    
        SmartDecrypt = StringToDecrypt
        
        For i = &H1 To VBA.Len(SmartDecrypt) Step &H1
            strEncrypt = strEncrypt & (VBA.Asc(VBA.Mid(SmartDecrypt, i, &H1)) - &H93)
        Next i
        
    End If
    
    SmartDecrypt = VBA.vbNullString
    
    If VBA.LenB(VBA.Trim$(strEncrypt)) = &H0 Then strEncrypt = StringToDecrypt
    
    Do While VBA.LenB(VBA.Trim$(strEncrypt)) <> &H0
        
        CharPos = VBA.Left(strEncrypt, &H1)
        strEncrypt = VBA.Mid(strEncrypt, &H2)
        CharCode = VBA.Left(strEncrypt, CharPos)
        strEncrypt = VBA.Mid(strEncrypt, VBA.Len(CharCode) + &H1)
        SmartDecrypt = SmartDecrypt & VBA.Chr(CharCode)
                
    Loop
    
    Exit Function
    
ErrorHandler:
    
End Function

'------------------------------------------------------------------------------------------
Public Function SmartEncrypt(StringToEncrypt As String, Optional AlphaEncoding As Boolean = True) As String
On Local Error GoTo ErrorHandler
    
    Dim i&
    Dim Char$, strEncrypt$
    
    If VBA.Len(VBA.Trim$(StringToEncrypt)) = &H0 Then Exit Function
    
    For i = &H1 To VBA.Len(StringToEncrypt) Step &H1
        Char = VBA.Asc(VBA.Mid(StringToEncrypt, i, &H1))
        SmartEncrypt = SmartEncrypt & VBA.Len(Char) & Char
    Next i
    
    If AlphaEncoding Then
    
        strEncrypt = SmartEncrypt
        SmartEncrypt = VBA.vbNullString
        
        For i = &H1 To VBA.Len(strEncrypt) Step &H1
            SmartEncrypt = SmartEncrypt & VBA.Chr(VBA.Mid(strEncrypt, i, &H1) + &H93)
        Next i
        
    End If
    
    Exit Function
    
ErrorHandler:
    
    SmartEncrypt = "Error encrypting string"
    
End Function

