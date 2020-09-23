VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Scanner"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   BeginProperty Font 
      Name            =   "Unispace"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton comLocal 
      Caption         =   "Scan for Local Ports in Use"
      Height          =   495
      Left            =   4440
      TabIndex        =   18
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Frame frm5 
      BackColor       =   &H00404040&
      Caption         =   "Options"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   1920
      Width           =   3495
      Begin VB.CheckBox optPortName 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   2640
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Identify Known Ports"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   270
         Width           =   2175
      End
   End
   Begin VB.CommandButton comStop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton comScan 
      Caption         =   "Scan"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton comClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame frm4 
      BackColor       =   &H00404040&
      Caption         =   "Open Ports"
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   7455
      Begin VB.ListBox lstPorts 
         BeginProperty Font 
            Name            =   "Unispace"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   7215
      End
      Begin VB.ListBox lstPorts2 
         Height          =   1035
         Left            =   360
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame frm3 
      BackColor       =   &H00404040&
      Caption         =   "Max Connections"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   3495
      Begin VB.TextBox txtMaxConnections 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "50"
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame frm2 
      BackColor       =   &H00404040&
      Caption         =   "Ports to Scan"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   240
      Width           =   3375
      Begin VB.TextBox txtLow 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtHigh 
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Text            =   "32767"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.Frame frm1 
      BackColor       =   &H00404040&
      Caption         =   "IP Address"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   120
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Timer timStatusUpdate 
      Interval        =   1000
      Left            =   8160
      Top             =   4440
   End
   Begin MSWinsockLib.Winsock sckScan 
      Index           =   0
      Left            =   8160
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar barStatus2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   6585
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   7144
            Text            =   "Port 0"
            TextSave        =   "Port 0"
            Object.ToolTipText     =   "Current port being scanned..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   7144
            Text            =   "0 ports scanned"
            TextSave        =   "0 ports scanned"
            Object.ToolTipText     =   "Total ports scanned..."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Unispace"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar barStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   6210
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   4763
            Text            =   "Idle"
            TextSave        =   "Idle"
            Object.ToolTipText     =   "Status of Port Scanner"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   4763
            Text            =   "0 ports per second"
            TextSave        =   "0 ports per second"
            Object.ToolTipText     =   "Speed of Port Scanning"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   4763
            Text            =   "0 seconds elapsed"
            TextSave        =   "0 seconds elapsed"
            Object.ToolTipText     =   "Time Since Beginning of Port Scan"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Unispace"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer timSck 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   8040
      Top             =   6840
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Scan
    MaxConn As Single
    Target As String
    PortLo As Double
    PortHi As Double
    Name As Boolean
End Type
Private Type ScanStats
    Elapsed As Integer
    Ports As Double
    Current As Double
    State As Integer
End Type

Const MaxPort As Integer = 32676
Const MaxConn As Integer = 150
Const TimeOut As Integer = 750

Dim States(1) As Variant
Dim Scan As Scan
Dim Stat As ScanStats
Dim Ports(MaxPort) As Variant
Dim GlobalTF As Boolean

Private Sub comClear_Click()
lstPorts.Clear          'Clears the lists.
lstPorts2.Clear
End Sub

Private Sub comLocal_Click()
Dim i As Integer

Scan.MaxConn = Val(txtMaxConnections.Text)     'Sets the scan values.
Scan.PortHi = Val(txtHigh.Text)
Scan.PortLo = Val(txtLow.Text)
Scan.Target = "127.0.0.1"
If optPortName.Value = vbChecked Then
    Scan.Name = True
Else
    Scan.Name = False
End If
Stat.Current = Scan.PortLo                     'Resets the stat values.
Stat.Elapsed = 0
Stat.Ports = 1
Stat.State = 1

If barStatus.Panels(3).Text <> "0 seconds elapsed" Then barStatus.Panels(3).Text = "0 seconds elapsed"
If barStatus.Panels(2).Text <> "0 ports per second" Then barStatus.Panels(2).Text = "0 ports per second"
txtLow.Enabled = False
txtHigh.Enabled = False
txtIP.Enabled = False
txtMaxConnections.Enabled = False              'Enables and disables the appro-
comClear.Enabled = False                       'priate text boxes, buttons, and
comScan.Enabled = False                        'options.
comStop.Enabled = True
optPortName.Enabled = False
comLocal.Enabled = False

DoEvents

For i = 1 To sckScan.ubound                    'Unloads all winsock elements
    Unload sckScan(i)                          'except element 0.
Next

For i = Scan.PortLo + 1 To Scan.PortHi + 1     'Begins the self-scan loop.
    If Stat.State <> 1 Then Exit For           'If the user clicked stop, STOP!
    Stat.Current = i                           'Sets the current port #.
    Stat.Ports = Stat.Ports + 1                'Adds to the "ttl ports scanned" value
    sckScan(0).LocalPort = i                   'Sets the port to check.
    GlobalTF = False                           'Sets the global true.false boolean.
    On Error GoTo Err
    sckScan(0).Listen                          'Attempts to listen on the local port,
                                               'if an error is encountered, that means
                                               'that the port is already locally in
                                               'use.
    sckScan(0).Close                           'Closes the port (if it was opened).
    If GlobalTF = True And sckScan(0).LocalPort <> sckScan(0).RemotePort Then
        If Scan.Name = True And Ports(sckScan(0).LocalPort) <> "" Then
            lstPorts.AddItem "(local)  Port #" & sckScan(0).LocalPort & ", " & Ports(sckScan(0).LocalPort)
            lstPorts2.AddItem "[" & Time & "] Local  Scan] " & sckScan(0).LocalIP & " Port] " & sckScan(0).LocalPort & " Port Use] " & Ports(sckScan(0).LocalPort)
        Else
            lstPorts.AddItem "(local)  Port #" & sckScan(0).LocalPort
            lstPorts2.AddItem "[" & Time & "] Local  Scan] " & sckScan(0).LocalIP & " Port] " & sckScan(0).LocalPort
        End If
    End If
    DoEvents
Next

If barStatus.Panels(3).Text <> "0 seconds elapsed" Then barStatus.Panels(3).Text = "0 seconds elapsed"
If barStatus.Panels(2).Text <> "0 ports per second" Then barStatus.Panels(2).Text = "0 ports per second"
txtLow.Enabled = True
txtHigh.Enabled = True
txtIP.Enabled = True                           'Resets stat values and enables and
txtMaxConnections.Enabled = True               'disables the appropriate text boxes,
comClear.Enabled = True                        'options, and buttons.
comScan.Enabled = True
comStop.Enabled = False
optPortName.Enabled = True
comLocal.Enabled = True
Stat.State = 0

Exit Sub
Err:
GlobalTF = True
Resume Next
End Sub

Private Sub comScan_Click()
Dim i As Integer

Scan.MaxConn = Val(txtMaxConnections.Text)      'Set all the scan values
Scan.PortHi = Val(txtHigh.Text)                 'to what was entered
Scan.PortLo = Val(txtLow.Text)                  'into the text boxes.
Scan.Target = txtIP.Text
If optPortName.Value = vbChecked Then
    Scan.Name = True
Else
    Scan.Name = False
End If
Stat.Current = Scan.PortLo                      'Resets the stat values.
Stat.Elapsed = 0
Stat.Ports = 1
Stat.State = 1

For i = 1 To sckScan.ubound                     'Unloads all winsock elements
    On Error Resume Next                        'except 0.
    Unload sckScan(i)
    Unload timSck(i)
Next

Stat.Current = Stat.Current - 1                 'Loads X amount of winsock elements
For i = 1 To Scan.MaxConn                       'where X is the Max Connection
    Load sckScan(i)                             'value.
    Stat.Current = Stat.Current + 1             'Starts the scan loop by connecting
    On Error GoTo Err
    sckScan(i).Connect Scan.Target, Stat.Current 'to the target.
    Load timSck(i)
    timSck(i).Interval = TimeOut
    timSck(i).Enabled = True
    DoEvents
Next

If barStatus.Panels(3).Text <> "0 seconds elapsed" Then barStatus.Panels(3).Text = "0 seconds elapsed"
If barStatus.Panels(2).Text <> "0 ports per second" Then barStatus.Panels(2).Text = "0 ports per second"
txtLow.Enabled = False
txtHigh.Enabled = False
txtIP.Enabled = False                           'Resets all the values and disables
txtMaxConnections.Enabled = False               'and enables the buttons, options,
comClear.Enabled = False                        'and text boxes.
comScan.Enabled = False
comStop.Enabled = True
optPortName.Enabled = False
comLocal.Enabled = False

Exit Sub
Err:
sckScan(i).Close
Resume Next
End Sub

Private Sub comStop_Click()
Dim i As Integer

For i = 1 To Scan.MaxConn                           'Unloads all the winsock elements
    On Error Resume Next                            'except element 0.
    sckScan(i).Close
    Unload sckScan(i)
    Unload timSck(i)
Next
Stat.State = 0
Stat.Elapsed = 0
Stat.Ports = 0
txtLow.Enabled = True
txtHigh.Enabled = True
txtIP.Enabled = True                                'Resets all stat values,
txtMaxConnections.Enabled = True                    'enabled and disables the appro-
comClear.Enabled = True                             'priate text boxes, buttons, and
comScan.Enabled = True                              'options
comStop.Enabled = False
optPortName.Enabled = True
comLocal.Enabled = True
Stat.Current = 0
Stat.Elapsed = 0
Stat.Ports = 0
If barStatus.Panels(3).Text <> "0 seconds elapsed" Then barStatus.Panels(3).Text = "0 seconds elapsed"
If barStatus.Panels(2).Text <> "0 ports per second" Then barStatus.Panels(2).Text = "0 ports per second"
If barStatus2.Panels(1).Text <> "Port 0" Then barStatus2.Panels(1).Text = "Port 0"
If barStatus2.Panels(2).Text <> "0 ports scanned" Then barStatus2.Panels(2).Text = "0 ports scanned"
End Sub

Private Sub Form_Load()
Dim i As Integer

Ports(21) = "File Transfer Protocol (FTP)"          'Sets the names of a couple
Ports(23) = "Telnet"                                'known ports.
Ports(53) = "Domain Name Server (DNS)"              'The port # is the index #.
Ports(79) = "Finger"
Ports(80) = "Hyper Text Transfer Protocol (HTTP)"
Ports(443) = "Hyper Text Transfer Protocol [Secure] (HTTPS)"
Ports(1214) = "KaZaA"
Ports(1433) = "Microsoft SQL Server"
Ports(6667) = "Internet Relay Chat (IRC)"

Stat.State = 0                                      'Sets the initial values
States(0) = "Idle"                                  'and ... yea.
States(1) = "Scanning"
txtMaxConnections.Text = 50
txtLow.Text = 0
txtHigh.Text = MaxPort
barStatus.Panels(1).Text = "Idle"
barStatus.Panels(2).Text = "0 ports per second"
barStatus.Panels(3).Text = "0 seconds elapsed"
For i = 1 To 3
    barStatus.Panels(i).Alignment = sbrLeft
Next
For i = 1 To 2
    barStatus2.Panels(i).Alignment = sbrLeft
Next
timStatusUpdate_Timer                               'Run the timer event.
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstPorts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then lstPorts.ListIndex = -1
End Sub

Private Sub sckScan_Connect(Index As Integer)
    'During a remote port scan, if a winsock element's connection request is accepted,
    'this sub is called. So that means that the port this winsock element connected to
    'was open and listening for connections.
If Scan.Name = True And Ports(sckScan(Index).RemotePort) <> "" Then
    lstPorts.AddItem "(remote) Port #" & sckScan(Index).RemotePort & ", " & Ports(sckScan(Index).RemotePort) & ", " & sckScan(Index).RemoteHostIP
    lstPorts2.AddItem "[" & Time & "] Remote Scan] " & sckScan(Index).RemoteHostIP & " Port] " & sckScan(Index).RemotePort & " Port Use] " & Ports(sckScan(Index).RemotePort)
Else
    lstPorts.AddItem "(remote) Port #" & sckScan(Index).RemotePort & ", " & sckScan(Index).RemoteHostIP
    lstPorts2.AddItem "[" & Time & "] Remote Scan] " & sckScan(Index).RemoteHostIP & " Port] " & sckScan(Index).RemotePort
End If
NextPort Index
End Sub

Private Sub sckScan_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'During a remote port scan, if a winsock element's connection request is rejected,
    'this sub is called. So that means that the port this winsock element attempted to
    'connect to was not open and refused the connection request.
On Error GoTo Err
NextPort Index
Exit Sub
Err:
MsgBox "An error was encountered." & vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbCritical
End
End Sub

Private Sub timSck_Timer(Index As Integer)
    'If the winsock element (with the same index as this timer element) has not
    'yet had the connection request accepted or rejected, this timer will count
    'it off as an error.
On Error GoTo Err
NextPort Index
Exit Sub
Err:
MsgBox "An error was encountered." & vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbCritical
End
End Sub

Private Sub timStatusUpdate_Timer()
Dim i, X As Integer
If barStatus.Panels(1).Text <> States(Stat.State) Then barStatus.Panels(1).Text = States(Stat.State)
If Stat.State = 1 Then
    Stat.Elapsed = Stat.Elapsed + 1
    barStatus.Panels(3).Text = Stat.Elapsed & " seconds elapsed"
    i = CInt(Stat.Ports / Stat.Elapsed)
    If i & " ports per second" <> barStatus.Panels(2).Text Then barStatus.Panels(2).Text = i & " ports per second"
    If Stat.Ports & " ports scanned" <> barStatus2.Panels(2).Text Then barStatus2.Panels(2) = Stat.Ports & " ports scanned"
    If "Port " & Stat.Current <> barStatus2.Panels(1).Text Then barStatus2.Panels(1).Text = "Port " & Stat.Current
    txtLow.Enabled = False
    txtHigh.Enabled = False                 'If a scan is in progress, this timer
    txtIP.Enabled = False                   'ensures that the appropriate buttons,
    txtMaxConnections.Enabled = False       'text boxes, and options are enabled
    comClear.Enabled = False                'and disabled, and also updates the
    comScan.Enabled = False                 'statistics in the stat bars.
    comStop.Enabled = True
    optPortName.Enabled = False
    comLocal.Enabled = False
Else
    txtLow.Enabled = True
    txtHigh.Enabled = True                  'If the program is idle, this timer
    txtIP.Enabled = True                    'ensures that the appropriate buttons,
    txtMaxConnections.Enabled = True        'text boxes, and options are enabled
    comClear.Enabled = True                 'and disabled, and also sets all the stat
    comScan.Enabled = True                  'values to 0.
    comStop.Enabled = False
    optPortName.Enabled = True
    comLocal.Enabled = True
    Stat.Current = 0
    Stat.Elapsed = 0
    Stat.Ports = 0
    If barStatus.Panels(3).Text <> "0 seconds elapsed" Then barStatus.Panels(3).Text = "0 seconds elapsed"
    If barStatus.Panels(2).Text <> "0 ports per second" Then barStatus.Panels(2).Text = "0 ports per second"
    If barStatus2.Panels(1).Text <> "Port 0" Then barStatus2.Panels(1).Text = "Port 0"
    If barStatus2.Panels(2).Text <> "0 ports scanned" Then barStatus2.Panels(2).Text = "0 ports scanned"
End If
End Sub

Private Sub txtHigh_LostFocus()
'This event ensures that the txtHigh text box value is legit.
txtHigh.Text = Replace(txtHigh.Text, " ", "")
If CheckVal(txtHigh.Text) = False Then txtHigh.Text = "65000": txtHigh.SetFocus
If Val(txtHigh.Text) <= Val(txtLow.Text) Then txtHigh.Text = Val(txtLow.Text) + 1: txtHigh.SetFocus
End Sub

Private Sub txtIP_LostFocus()
'This event ensures that the txtIP text box value is legit.
txtIP.Text = Replace(txtIP.Text, " ", "")
If CheckIP(txtIP.Text) = False Then txtIP.SelStart = 0: txtIP.SelLength = Len(txtIP.Text): txtIP.SetFocus
End Sub

Private Sub txtLow_LostFocus()
'This event ensures that the txtLow text box value is legit.
txtLow.Text = Replace(txtLow.Text, " ", "")
If CheckVal(txtLow.Text) = False Then txtLow.Text = "0": txtLow.SetFocus
If Val(txtLow.Text) >= Val(txtHigh.Text) Then txtLow.Text = Val(txtHigh.Text) - 1: txtLow.SetFocus
End Sub

Private Sub txtMaxConnections_LostFocus()
'This event ensures that the txtMaxConnections text box value is legit.
txtMaxConnections.Text = Replace(txtMaxConnections.Text, " ", "")
If CheckVal(txtMaxConnections.Text) = False Then txtMaxConnections.Text = "50": txtMaxConnections.SetFocus
If Val(txtMaxConnections.Text) > MaxConn Then txtMaxConnections.Text = MaxConn: txtMaxConnections.SetFocus
End Sub

Private Function CheckVal(Val As String) As Boolean
Dim i, X As Integer
Dim a, B As String

For i = 1 To Len(Val)
    If Asc(Mid(Val, i, 1)) > Asc("9") Or Asc(Mid(Val, i, 1)) < Asc("0") Then
        CheckVal = False
        Exit Function
    End If
Next
CheckVal = True
End Function

Private Function CheckIP(Val As String) As Boolean
Dim i, X As Integer
Dim a, B As String
Dim c As Variant

For i = 1 To Len(Val)
    If Asc(Mid(Val, i, 1)) > Asc("9") Or Asc(Mid(Val, i, 1)) < Asc("0") Then
        If Asc(Mid(Val, i, 1)) <> Asc(".") Then
            CheckIP = False
            Exit Function
        End If
    End If
Next

c = Split(Val, ".")
If UBound(c) <> 3 Then CheckIP = False: Exit Function

For i = 0 To 3
    If c(i) > 255 Then CheckIP = False: Exit Function
Next

CheckIP = True
End Function

Private Sub NextPort(Index As Integer)
Dim i As Integer

'When a winsock element's connection request is accepted or rejected, this sub is
'called and either ends the scan process or reuses the winsock element to attempt to
'connect to another port.

sckScan(Index).Close                                'Closes the winsock element.
Stat.Ports = Stat.Ports + 1                         'Increments the scanned ports val.
Stat.Current = Stat.Current + 1                     'Increments the current port val.
If Scan.PortHi >= Stat.Current Then                 'Makes sure the scan hasn't hit
                                                    'the high port val.
    sckScan(Index).Connect Scan.Target, Stat.Current 'Simply has the winsock element
    timSck(Index).Enabled = False                    'connect to another port.
    timSck(Index).Enabled = True
    On Error GoTo Err
    DoEvents
Else
    For i = 1 To sckScan.ubound                     'Unloads all winsock elements
        sckScan(i).Close                            'except element 0.
        Unload sckScan(i)
        Unload timSck(i)
    Next
    Stat.State = 0
    Stat.Elapsed = 0
    Stat.Ports = 0
    txtLow.Enabled = True
    txtHigh.Enabled = True
    txtIP.Enabled = True
    txtMaxConnections.Enabled = True                'Resets all stat values and
    comClear.Enabled = True                         'disables and enables the appro-
    comScan.Enabled = True                          'priate buttons, text boxes, and
    comStop.Enabled = False                         'options.
    optPortName.Enabled = True
    comLocal.Enabled = True
    Stat.Current = 0
    Stat.Elapsed = 0
    Stat.Ports = 0
    If barStatus.Panels(3).Text <> "0 seconds elapsed" Then barStatus.Panels(3).Text = "0 seconds elapsed"
    If barStatus.Panels(2).Text <> "0 ports per second" Then barStatus.Panels(2).Text = "0 ports per second"
    If barStatus2.Panels(1).Text <> "Port 0" Then barStatus2.Panels(1).Text = "Port 0"
    If barStatus2.Panels(2).Text <> "0 ports scanned" Then barStatus2.Panels(2).Text = "0 ports scanned"
End If
Exit Sub
Err:
Resume Next
End Sub
