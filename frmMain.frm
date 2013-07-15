VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MrBlobby's Modem Dialler"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrConnectCom 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1680
      Top             =   3000
   End
   Begin VB.Timer tmrTimeCheck 
      Interval        =   1000
      Left            =   1200
      Top             =   3000
   End
   Begin MSComDlg.CommonDialog cdgMain 
      Left            =   480
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Start"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4335
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3493
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3493
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3493
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2160
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame pnlSettings 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5775
      Begin VB.ComboBox cboComPortSpeed 
         Height          =   315
         Left            =   4560
         TabIndex        =   22
         Text            =   "cboComPortSpeed"
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optWaitTime 
         Caption         =   "Minutes"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   21
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton optWaitTime 
         Caption         =   "Seconds"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   20
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Frame fraConnectedTime 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   4200
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
         Begin VB.OptionButton optConnectedTime 
            Caption         =   "Seconds"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   19
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optConnectedTime 
            Caption         =   "Minutes"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.ComboBox cboComPort 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveSettings 
         Caption         =   "Save Settings"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdLoadSettings 
         Caption         =   "Load Settings"
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtPhoneNumber 
         Height          =   285
         Left            =   3240
         TabIndex        =   8
         Text            =   "txtPhoneNumber"
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox chkShowErrors 
         Caption         =   "Show Errors"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox chkLogErrors 
         Caption         =   "Log Errors"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtConnectedTime 
         Height          =   285
         Left            =   3240
         TabIndex        =   5
         Text            =   "txtConnectedTime"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtWaitTime 
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Text            =   "txtWaitTime"
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox chkAutoStart 
         Caption         =   "Auto dial when app loaded"
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label lblComPort 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Select Modem COM port"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "2. Enter telephone number"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "3. Time to stay connected"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "4. Time to wait before redialling"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "5. Set options"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DisconnectTime As Date
Dim ConnectTime As Date
Dim CurrentStatus As String
Dim DesiredStatus As String
Dim LogCount As Long

Private Sub cmdLoadSettings_Click()

On Error GoTo cmdLoadSettings_Click_Error

    Dim mbResult
    mbResult = MsgBox("Load Settings?", vbYesNo, "Load")
    If mbResult = vbYes Then
        LoadSettings
    End If
    
    Exit Sub
    
cmdLoadSettings_Click_Error:

    RaiseError Err.Number, Err.Description, Erl, "cmdLoadSettings_Click", "Form", "frmMain", Now
    
End Sub

Private Sub cmdSaveSettings_Click()

On Error GoTo cmdSaveSettings_Click_Error

    WriteINI App.Path & "\settings.ini", "Settings", "COMPortIndex", cboComPort.ListIndex
    WriteINI App.Path & "\settings.ini", "Settings", "COMPortSpeedIndex", cboComPortSpeed.ListIndex
    WriteINI App.Path & "\settings.ini", "Settings", "PhoneNumber", txtPhoneNumber.Text
    WriteINI App.Path & "\settings.ini", "Settings", "ConnectedTime", txtConnectedTime.Text
    If optConnectedTime(0).Value = True Then
        WriteINI App.Path & "\settings.ini", "Settings", "ConnectedTimeUnit", 0
    Else
        WriteINI App.Path & "\settings.ini", "Settings", "ConnectedTimeUnit", 1
    End If
    WriteINI App.Path & "\settings.ini", "Settings", "WaitTime", txtWaitTime.Text
    If optWaitTime(0).Value = True Then
        WriteINI App.Path & "\settings.ini", "Settings", "WaitTimeUnit", 0
    Else
        WriteINI App.Path & "\settings.ini", "Settings", "WaitTimeUnit", 1
    End If
    WriteINI App.Path & "\settings.ini", "Settings", "ShowErrors", chkShowErrors.Value
    WriteINI App.Path & "\settings.ini", "Settings", "LogErrors", chkLogErrors.Value
    WriteINI App.Path & "\settings.ini", "Settings", "AutoStart", chkAutoStart.Value
        
    StatusBar1.Panels(1).Text = "Settings saved"

Exit Sub

cmdSaveSettings_Click_Error:

    RaiseError Err.Number, Err.Description, Erl, "cmdSaveSettings_Click", "Form", "frmMain", Now


End Sub

Private Sub cmdStartStop_Click()

On Error GoTo cmdStartStop_Click_Error

If cmdStartStop.Tag = "Start" Then
    
    EnableControls (False)
    
    
    DesiredStatus = "Running"
    CurrentStatus = "Stopped"
    tmrTimeCheck.Enabled = True
    ConnectCom
    tmrConnectCom.Enabled = True
    
    StatusBar1.Panels(1).Text = "Running"
    cmdStartStop.Tag = "Stop"
    cmdStartStop.Caption = "Stop"
        
Else

    EnableControls (True)

    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
    DesiredStatus = "Stopped"
    StatusBar1.Panels(1).Text = "Stopped"
    StatusBar1.Panels(2).Text = ""
    StatusBar1.Panels(3).Text = ""
    
    tmrTimeCheck.Enabled = False
    tmrConnectCom.Enabled = False

    cmdStartStop.Tag = "Start"
    cmdStartStop.Caption = "Start"
    
End If

Exit Sub


cmdStartStop_Click_Error:

    RaiseError Err.Number, Err.Description, Erl, "cmdStartStop_Click", "Form", "frmMain", Now
    
End Sub

Private Function EnableControls(Enable As Boolean)

On Error GoTo EnableControls_Error

    cboComPort.Enabled = Enable
    txtPhoneNumber.Enabled = Enable
    txtConnectedTime.Enabled = Enable
    optConnectedTime(0).Enabled = Enable
    optConnectedTime(1).Enabled = Enable
    txtWaitTime.Enabled = Enable
    optWaitTime(0).Enabled = Enable
    optWaitTime(1).Enabled = Enable
    chkAutoStart.Enabled = Enable
    chkShowErrors.Enabled = Enable
    chkLogErrors.Enabled = Enable
    cmdSaveSettings.Enabled = Enable
    cmdLoadSettings.Enabled = Enable
    
EnableControls_Error:

    RaiseError Err.Number, Err.Description, Erl, "EnableControls", "Form", "frmMain", Now

End Function

Private Sub Form_Load()

On Error GoTo Load_Error

Dim i As Integer

    For i = 1 To 100
        If ComPortExists(i) Then
            cboComPort.AddItem (i)
        End If
    Next i

    cboComPortSpeed.AddItem ("1200")
    cboComPortSpeed.AddItem ("4800")
    cboComPortSpeed.AddItem ("9600")
    cboComPortSpeed.AddItem ("19200")
    cboComPortSpeed.AddItem ("38400")
    cboComPortSpeed.AddItem ("57600")
    cboComPortSpeed.AddItem ("115200")
    
    ConnectTime = 0
    DisconnectTime = 0

    
    cmdStartStop.Tag = "Start"
    LoadSettings (True)
    
    If chkAutoStart.Value = 1 Then
        cmdStartStop_Click
    End If

    Exit Sub

Load_Error:

    RaiseError Err.Number, Err.Description, Erl, "frmMain_Load", "Form", "frmMain", Now

End Sub

Private Function LoadSettings(Optional Default As Boolean = False)

On Error GoTo LoadSettings_Error


    If FileExists(App.Path & "\settings.ini") Then
        cboComPort.ListIndex = ReadINI(App.Path & "\settings.ini", "Settings", "COMPortIndex", vbString)
        cboComPortSpeed.ListIndex = ReadINI(App.Path & "\settings.ini", "Settings", "COMPortSpeedIndex", vbString)
        txtPhoneNumber.Text = ReadINI(App.Path & "\settings.ini", "Settings", "PhoneNumber", vbString)
        txtConnectedTime.Text = ReadINI(App.Path & "\settings.ini", "Settings", "ConnectedTime", vbInteger)
        optConnectedTime(ReadINI(App.Path & "\settings.ini", "Settings", "ConnectedTimeUnit", vbInteger)).Value = True
        txtWaitTime.Text = ReadINI(App.Path & "\settings.ini", "Settings", "WaitTime", vbInteger)
        optWaitTime(ReadINI(App.Path & "\settings.ini", "Settings", "WaitTimeUnit", vbInteger)).Value = True
        chkShowErrors.Value = ReadINI(App.Path & "\settings.ini", "Settings", "ShowErrors", vbBoolean)
        chkLogErrors.Value = ReadINI(App.Path & "\settings.ini", "Settings", "LogErrors", vbBoolean)
        chkAutoStart.Value = ReadINI(App.Path & "\settings.ini", "Settings", "AutoStart", vbBoolean)
        StatusBar1.Panels(1).Text = "Settings loaded"
    Else
        If Default = False Then
            Dim mbResult
            mbResult = MsgBox("settings.ini file not found, load default values instead?", vbYesNo, "Error")
            If mbResult = vbYes Then
                Default = True
            End If
        End If
        
        If Default = True Then
            cboComPort.ListIndex = -1
            cboComPortSpeed.ListIndex = -1
            txtPhoneNumber.Text = "01234567890"
            txtConnectedTime.Text = 60
            optConnectedTime(0).Value = True
            txtWaitTime.Text = 60
            optWaitTime(1).Value = True
            chkShowErrors.Value = 1
            chkLogErrors.Value = 1
            chkAutoStart.Value = 0
            StatusBar1.Panels(1).Text = "Default settings loaded"
        End If
    End If
    
    Exit Function
    
LoadSettings_Error:

    RaiseError Err.Number, Err.Description, Erl, "LoadSettings", "Form", "frmMain", Now

End Function

Private Sub tmrTimeCheck_Timer()

On Error GoTo tmrTimeCheck_Timer_Error

    If ConnectTime <> 0 Then
    
        If ConnectTime < Now() Then
        
            If MSComm1.PortOpen = True Then
                
                DialNumber
                
            End If
            
        Else
           
           StatusBar1.Panels(3).Text = "Reconnect in: " & Format(ConnectTime - Now(), "hh:mm:ss")
            
        End If
            
        Exit Sub
            
    End If
    
    If DisconnectTime <> 0 Then
    
        If DisconnectTime < Now() Then
        
            If MSComm1.PortOpen = True Then
                
                HangUp
                
            End If
            
        Else
        
            StatusBar1.Panels(3).Text = "Connected: " & Format(DisconnectTime - Now(), "hh:mm:ss")
            
        End If
        
    End If
    
    Exit Sub
    
tmrTimeCheck_Timer_Error:

    RaiseError Err.Number, Err.Description, Erl, "tmrTimeCheck_Timer", "Form", "frmMain", Now
    

End Sub

Public Sub ConnectCom()
      
      
On Error GoTo ConnectCom_Error

On Error GoTo LocalHandler

Dim Port As Integer
Dim PortSpeed As Integer
Dim Protocol As Integer
Dim ret

Port = cboComPort.Text
PortSpeed = cboComPortSpeed.Text
Protocol = 0

If ComPortExists(Port) Then
    
        If MSComm1.PortOpen = True Then
            Exit Sub
        End If

         With MSComm1
            .CommPort = Port
            .Handshaking = 0 - comNone
            .RThreshold = 1
            .RTSEnable = False
            .Settings = Str(PortSpeed) + ",n,8,1"
            .SThreshold = 0
            .PortOpen = True

        End With
        
        tmrConnectCom.Enabled = False
        
        StatusBar1.Panels(2).Text = "Com port connected"
        
        DialNumber
        
        Exit Sub

End If

    StatusBar1.Panels(2).Text = "Com port not found"
    
LocalHandler:

    StatusBar1.Panels(2).Text = "Com port not found"

    On Error GoTo 0
    Exit Sub

ConnectCom_Error:

    RaiseError Err.Number, Err.Description, Erl, "ConnectCom", "Form", "frmMain", Now
    
End Sub

Private Sub DialNumber()

On Error GoTo DialNumber_Error
    
    If MSComm1.PortOpen = True Then
    
        StatusBar1.Panels(3).Text = "Dialling"

        MSComm1.Output = "ATDT " & txtPhoneNumber.Text & ";"
        Debug.Print "(" & LogCount & ") ATDT " & txtPhoneNumber.Text & ";"
        LogCount = LogCount + 1
        
        ConnectTime = 0
        
        If optConnectedTime(0).Value = True Then
            DisconnectTime = DateAdd("n", txtConnectedTime.Text, Now())
        Else
            DisconnectTime = DateAdd("s", txtConnectedTime.Text, Now())
        End If
        
    End If
    
    Exit Sub
    
DialNumber_Error:

    RaiseError Err.Number, Err.Description, Erl, "DialError", "Form", "frmMain", Now
    
End Sub


Private Sub HangUp()

On Error GoTo HangUp_Error
    
    If MSComm1.PortOpen = True Then
    
        StatusBar1.Panels(3).Text = "Hanging up"
    
        MSComm1.Output = "ATH;"
        Debug.Print "(" & LogCount & ") ATH;"
        LogCount = LogCount + 1
        
        DisconnectTime = 0
    
        If optWaitTime(0).Value = True Then
            ConnectTime = DateAdd("n", txtWaitTime.Text, Now())
        Else
            ConnectTime = DateAdd("s", txtWaitTime.Text, Now())
        End If
    
    End If
    
    Exit Sub
    
HangUp_Error:

    RaiseError Err.Number, Err.Description, Erl, "HangUp", "Form", "frmMain", Now
    
End Sub

Private Sub tmrConnectCom_Timer()

On Error GoTo tmrConnectCom_Timer_Error

    ConnectCom

    On Error GoTo 0
    Exit Sub

tmrConnectCom_Timer_Error:

    RaiseError Err.Number, Err.Description, Erl, "tmrConnectCom_Timer", "Form", "frmMain", Now
    
End Sub

