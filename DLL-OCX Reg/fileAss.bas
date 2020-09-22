Attribute VB_Name = "FileAss"
'// You may use this code all you want on the condition you keep this simple comment
'// Anyone who improves the code please let me know.
'// Date     : 21/1/2000
'// Author   : Damien McGivern
'// E-Mail   : Damien@Dingo-Delights.co.uk
'// Web Site : www.dingo-delights.co.uk

'// Improved 23/1/200 - New parameters 'Switch', 'PromptOnError', better error handling

'// Parameters
'// Required    Extension       (Str) ie ".exe"
'// Required    FileType        (Str) ie "VB.Form"
'// Required    FileTYpeName    (Str) ie. "Visual Basic Form"
'// Required    Action          (Str) ie. "Open" or "Edit"
'// Optional    Switch          (Str) ie. "/u"                  Default = ""
'// Optional    SetIcon         (Bol)                           Default = False
'// Optional    DefaultIcon     (Str) ie. "C:\Myapp,0"
'// Optional    PromptOnError   (Bol)                           Default = False


Option Explicit

Private Const REG_SZ As Long = 1

Private Const HKEY_CLASSES_ROOT = &H80000000

Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long


Public Function CreateFileAss(Extension As String, FileType As String, _
        FileTypeName As String, Action As String, AppPath As String, Optional Switch As String = "", _
        Optional SetIcon As Boolean = False, Optional DefaultIcon As String, Optional PromptOnError As Boolean = False) As Boolean

    On Error GoTo ErrorHandler:
    
    '// Check that AppPath exists.
    If Dir(AppPath, vbNormal) = "" Then
        If PromptOnError Then MsgBox "The application path '" & AppPath & "' cannot be found.", vbCritical + vbOKOnly, "DLL/OCX Register"
        CreateFileAss = False
        Exit Function
    End If
    
    Dim ERROR_CHARS As String
    Dim i As Integer
    
    ERROR_CHARS = "\/:*?<>|" & Chr(34)
    '// Check for invalid chars
    For i = 1 To Len(Extension)
        If InStr(1, ERROR_CHARS, Mid(Extension, i, 1), vbTextCompare) Then
            If PromptOnError Then MsgBox "The file extension '" & Extension & "' contains an illegal char (\/:*?<>|" & Chr(34) & ").", vbCritical + vbOKOnly, "DLL/OCX Register"
            CreateFileAss = False
            Exit Function
        End If
    Next
    
    If Switch <> "" Then Switch = " " & Trim(Switch)

    '// Check extension has "." at front
    If Asc(Extension) <> 46 Then Extension = "." & Extension
    
    

    Dim sPath As String

    Call CreateSubKey(Extension, HKEY_CLASSES_ROOT)
    
    Call SetKeyDefault(Extension, "", FileType)
    
    Call CreateSubKey(FileType & "\DefaultIcon", HKEY_CLASSES_ROOT)
    
    If SetIcon Then
        If DefaultIcon = "" Then
            '// This line of code sets the application's own icon as the default file icon
            SetKeyDefault FileType & "\DefaultIcon", "", Trim(AppPath & ",0")
        Else
            SetKeyDefault FileType & "\DefaultIcon", "", Trim(DefaultIcon)
        End If
    End If

    Call CreateSubKey(FileType & "\shell\" & Action & "\command", HKEY_CLASSES_ROOT)

    Call SetKeyDefault(FileType, "", FileTypeName)

    Call SetKeyDefault(FileType & "\shell\" & Action & "\command", "", AppPath & Switch & " %1")
    
    CreateFileAss = True
    Exit Function
    
ErrorHandler:

    If PromptOnError Then MsgBox "An error occured while attempting to create the file extension '" & Extension & "'.", vbCritical + vbOKOnly, "DLL/OCX Register"
    CreateFileAss = False
    
End Function


Private Function SetValue(ByVal hKey As Long, sValueName As String, vValue As Variant) As Long
    Dim nValue As Long
    Dim sValue As String
    sValue = vValue & Chr$(0)
    SetValue = RegSetValueExString(hKey, sValueName, 0&, REG_SZ, sValue, Len(sValue))
End Function


Private Sub CreateSubKey(sNewKeyName As String, lPredefinedKey As Long)
    '// This function creates a new sub key
    Dim hKey As Long, regReply As Long
    regReply = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, regReply)
    Call RegCloseKey(hKey)
End Sub


Private Sub SetKeyDefault(sKeyName As String, sValueName As String, vValueSetting As Variant)
    '// This function sets the default vaule of the key
    Dim regReply As Long, hKey As Long
    regReply = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    regReply = SetValue(hKey, sValueName, vValueSetting)
    Call RegCloseKey(hKey)
End Sub

