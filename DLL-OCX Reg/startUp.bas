Attribute VB_Name = "startUp"
Option Explicit

'// If the application is opened by it's self it creates the file associations.
'// If opened by a .dll or .ocx file it retrives the file's path from the Command statement
'// looks for a switch (/u to unregister) and registers or unregisters.


Public Sub Main()

    If Command = "" Then
        '// Create file associations
        Dim AppPath As String
        AppPath = App.path & "\" & App.EXEName & ".exe"
        Call FileAss.CreateFileAss(".dll", "DLL_FileType", "Application Extension", "Register", AppPath, , True, AppPath & ",1", True)
        Call FileAss.CreateFileAss(".dll", "DLL_FileType", "Application Extension", "Unregister", AppPath, "/u", True, AppPath & ",1", True)
        Call FileAss.CreateFileAss(".ocx", "OCX_FileType", "Active X Control", "Register", AppPath, , True, AppPath & ",2")
        Call FileAss.CreateFileAss(".ocx", "OCX_FileType", "Active X Control", "Unregister", AppPath, "/u", True, AppPath & ",2")
        MsgBox "File associations created", vbInformation, "DLL/OCX Register"
        Exit Sub
    End If
    

    If InStr(1, Command, "/u", vbTextCompare) Then
        '// Unregister
        Dim NewCommand As String
        NewCommand = Mid(Command, 4, Len(Command))
        If RegDLL_OCX(NewCommand, Unregister, True) Then
            MsgBox NewCommand & " has been unregistered successfully.", vbInformation, "DLL/OCX Register"
        End If
    Else
        '// Register
        If RegDLL_OCX(Command, Register, True) Then
            MsgBox Command & " has been registered successfully.", vbInformation, "DLL/OCX Register"
        End If
    End If
    
End Sub
