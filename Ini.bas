Attribute VB_Name = "modIni"
' AIMEE Ini Access Module
' =======================
'
' AIMEE (c) 2004, Written by Garry Mitchell (Confused) and Andrew Bickers (Danceheaven)
'
' This code is supplied without any support or warranties, either explicit or implied.


Option Explicit

Public Const MString = 0
Public Const MBool = 1
Public Const MLong = 2
Public Const MInteger = 3
Public Const MVariant = 4
Public Const MSingle = 5
Public Const MDec = 6



Public Declare Function ShellExecute Lib "SHELL32.DLL" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal _
    lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As _
    String, ByVal nShowCmd As Long) As Long

'filesystem handlers
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
   "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
   ByVal lpKeyName As String, ByVal lpDefault As String, _
   ByVal lpReturnedString As String, ByVal nSize As Long, _
   ByVal lpFileName As String) As Long
   
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
   "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
   ByVal lpKeyName As String, ByVal lpString As Any, _
   ByVal lpFileName As String) As Long

Function ReadINI(ByVal FileName As String, ByVal Section, ByVal KeyName As String, ByVal VarType As Long) As Variant


    Dim sRet As String
On Error GoTo ReadINI_Error

    sRet = String(255, Chr(0))
    ReadINI = left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), FileName))

    'If ReadINI = "" Then
    '    ErrH "The ReadINI function returned a null value from " & filename & " " & Section & " " & KeyName
    'End If
    Select Case VarType
    
        Case 0
            Exit Function
        
        Case 1
            If ReadINI = "" Then
fix1:
                ReadINI = False
                Exit Function
            Else
                On Error GoTo fix1
                ReadINI = CBool(ReadINI)
            End If
            
        Case 2
            If ReadINI = "" Then
fix2:
                ReadINI = "0"
                Exit Function
            Else
                On Error GoTo fix2
                ReadINI = CLng(ReadINI)
            End If
    
        Case 3
            If ReadINI = "" Then
fix3:
                ReadINI = "0"
                Exit Function
            Else
                On Error GoTo fix3
                ReadINI = CInt(ReadINI)
            End If
        
        Case 4
            If ReadINI = "" Then
fix4:
                ReadINI = "0"
                Exit Function
            Else
                On Error GoTo fix4
                ReadINI = CVar(ReadINI)
            End If
        
        Case 5
            If ReadINI = "" Then
fix5:
                ReadINI = "0"
                Exit Function
            Else
                On Error GoTo fix5
                ReadINI = CSng(ReadINI)
            End If
        
        Case 6
            If ReadINI = "" Then
fix6:
                ReadINI = "0"
                Exit Function
            Else
                On Error GoTo fix6
                ReadINI = CDec(ReadINI)
            End If
        
        Case Else
            On Error Resume Next
            ReadINI = CStr(ReadINI)
    
    End Select
    



    On Error GoTo 0
    Exit Function

ReadINI_Error:

    RaiseError Err.Number, Err.Description, Erl, "ReadINI", "Module", "modIni", Now

End Function

Function WriteINI(ByVal sFilename, ByVal sSection As String, ByVal sKeyName As String, ByVal sNewString As String) As Integer


    Dim r
On Error GoTo WriteINI_Error

    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFilename)
    WriteINI = r




    On Error GoTo 0
    Exit Function

WriteINI_Error:

    RaiseError Err.Number, Err.Description, Erl, "WriteINI", "Module", "modIni", Now

End Function

Public Sub DeleteSection(strFile As String, strSection As String)



On Error GoTo DeleteSection_Error

    WritePrivateProfileString strSection, vbNullString, vbNullString, strFile






    On Error GoTo 0
    Exit Sub

DeleteSection_Error:

    RaiseError Err.Number, Err.Description, Erl, "DeleteSection", "Module", "modIni", Now

End Sub

Public Sub DeleteKey(strFile As String, strSection As String, strKey As String)



On Error GoTo DeleteKey_Error

    WritePrivateProfileString strSection, strKey, vbNullString, strFile


    On Error GoTo 0
    Exit Sub

DeleteKey_Error:

    RaiseError Err.Number, Err.Description, Erl, "DeleteKey", "Module", "modIni", Now

End Sub

