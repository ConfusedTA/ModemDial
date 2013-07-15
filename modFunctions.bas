Attribute VB_Name = "modFunctions"
Public Declare Function GetDefaultCommConfig Lib "kernel32" _
  Alias "GetDefaultCommConfigA" (ByVal lpszName _
  As String, lpCC As COMMCONFIG, _
   lpdwSize As Long) As Long
   Public Type DCB
        DCBlength As Long
        BaudRate As Long
        fBitFields As Long
        wReserved As Integer
        XonLim As Integer
        XoffLim As Integer
        ByteSize As Byte
        Parity As Byte
        StopBits As Byte
        XonChar As Byte
        XoffChar As Byte
        ErrorChar As Byte
        EofChar As Byte
        EvtChar As Byte
        wReserved1 As Integer
End Type

Public Type COMMCONFIG
    dwSize As Long
    wVersion As Integer
    wReserved As Integer
    dcbx As DCB
    dwProviderSubType As Long
    dwProviderOffset As Long
    dwProviderSize As Long
    wcProviderData As Byte
End Type


Public Function ComPortExists(ByVal ComPort As Integer) _
   As Boolean
'*****************************************************
'EXAMPLE
    'Dim bAns As Boolean
    'bAns = ComPortExists(1)
    'If bans then
        'msgbox "Com Port 1 is available
    'Else
        'msgbox "Com Port 1 is not available
    'End if
'*************************************************
    Dim udtComConfig As COMMCONFIG
    Dim lUDTSize As Long
    Dim lRet As Long
    
On Error GoTo ComPortExists_Error

    lUDTSize = LenB(udtComConfig)
    lRet = GetDefaultCommConfig("COM" + Trim(Str(ComPort)) + _
        Chr(0), udtComConfig, lUDTSize)
    ComPortExists = lRet <> 0

    On Error GoTo 0
    Exit Function

ComPortExists_Error:

    RaiseError Err.Number, Err.Description, Erl, "ComPortExists", "Module", "modFunctions", Now

End Function

Public Function FileExists(ByVal FileName As String) As Boolean

        
    
    Dim FileInfo            As Variant

    'Set Default
On Error GoTo FileExists_Error

    FileExists = True
    
    'Set up error handler
    On Error Resume Next

    'Attempt to grab date and time
    FileInfo = FileDateTime(FileName)

    'Process errors

    Select Case Err

        Case 53, 76, 68   'File Does Not Exist
            FileExists = False
            Err = 0

        Case Else

            If Err <> 0 Then
                    
                End

            End If

    End Select
    

    On Error GoTo 0
    Exit Function

FileExists_Error:

    RaiseError Err.Number, Err.Description, Erl, "FileExists", "Module", "modFunctions", Now
    
End Function
Public Sub RaiseError(ErrNumber As String, ErrDescription As String, Erl As String, ProcedureName As String, ModuleType As String, ModuleName As String, TimeDate As Date)

    If ShowErrors = True Then
        MsgBox "Error " & ErrNumber & "(" & ErrDescription & ") on line " & Erl & " in procedure " & ProcedureName & " of " & ModuleType & " " & ModuleName, vbOKOnly, "Error"
    End If
    
    If LogErrors = True Then
        Open App.Path & "\error.log" For Append As #1
            Print #1, Format(TimeDate, "dd/mm/yyyy") & " " & Format(TimeDate, "hh:mm:ss")
            Print #1, ModuleType & ": " & ModuleName
            Print #1, "Procedure: " & ProcedureName
            Print #1, "Error: " & ErrNumber & " on line " & Erl & " - " & ErrDescription
            Print #1, "============================================================" & vbCrLf
        Close #1
    End If
    
End Sub


