Attribute VB_Name = "ocxDllReg"
'// I have enhanced code which I found at www.planetsourcecode.com
'// Date 23/1/200
'// Author Damien McGivern
'// E-Mail Damien@Dingo-Delights.co.uk

'// Required    File            (Str)
'// Optional    Process         (Reg/Unreg) Default = register
'// Optional    PromptOnError   (Bol)       Default = False

Option Explicit

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
        (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
        ByVal lpProcName As String) As Long
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, _
        ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, _
        ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, _
        lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Enum RegUnreg
    Register = 0
    Unregister = 1
End Enum

Const NO_ERROR = 0


Public Function RegDLL_OCX(File As String, Optional Process As RegUnreg = Register, _
      Optional PromptOnError As Boolean = False) As Boolean

    On Error Resume Next

    '// Check file exists
    If Dir(File, vbNormal) = "" Then
        If PromptOnError Then MsgBox "The file " & File & " doesn't exist", vbCritical, "DLL/OCX Register"
        RegDLL_OCX = False: Exit Function
    End If

    Dim LoadedLib As Long, EntryPoint As Long, ExitCode As Long
    Dim newThread As Long, newThreadID As Long



    LoadedLib = LoadLibrary(File)                          '// Load file

    If LoadedLib = 0 Then
        If PromptOnError Then MsgBox "An error occured while loading the file " & File, vbCritical, "DLL/OCX Register"
        RegDLL_OCX = False
        Exit Function
    End If

    '// Find right entery point
    If Process = Register Then
        EntryPoint = GetProcAddress(LoadedLib, "DllRegisterServer")
    ElseIf Process = Unregister Then
        EntryPoint = GetProcAddress(LoadedLib, "DllUnregisterServer")
    Else
        If PromptOnError Then MsgBox "An error occured while loading the file " & File, vbCritical, "DLL/OCX Register"
        RegDLL_OCX = False
        Exit Function
    End If


    If EntryPoint = vbNull Then
        If PromptOnError Then MsgBox "An error occured while locating the entery point for the file : " & vbNewLine & File, vbCritical, "DLL/OCX Register"
        FreeLibrary (LoadedLib)                            '// Unload libarary
        RegDLL_OCX = False
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    newThread = CreateThread(ByVal 0, 0, ByVal EntryPoint, ByVal 0, 0, newThreadID)    '// Create a new thread.

    If newThread = 0 Then
        Screen.MousePointer = vbDefault
        If PromptOnError Then MsgBox "An error occured while attempting to create a new thread.", vbCritical, "DLL/OCX Register"
        FreeLibrary (LoadedLib)                            '// Unload libarary
        Exit Function
    End If

    If WaitForSingleObject(newThread, 10000) <> 0 Then
        Screen.MousePointer = vbDefault
        If PromptOnError Then MsgBox "An error occured while attempting to register/unregister the file : " & vbNewLine & File, vbCritical, "DLL/OCX Register"
        ExitCode = GetExitCodeThread(newThread, ExitCode)
        ExitThread (ExitCode)
        FreeLibrary (LoadedLib)
        RegDLL_OCX = False
        Exit Function
    End If


    CloseHandle (newThread)                                '// Close thread
    FreeLibrary (LoadedLib)                                '// Unload libarary
    Screen.MousePointer = vbDefault                        '// Reset cursor
    RegDLL_OCX = True

End Function
