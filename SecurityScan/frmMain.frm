VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Caption         =   "Security Port Scanner"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sckScan 
      Index           =   0
      Left            =   600
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   3975
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton Command4 
         Caption         =   "What?"
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         ToolTipText     =   "What does this mean?"
         Top             =   3180
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3180
         Width           =   735
      End
      Begin VB.ListBox lstPorts 
         Height          =   2595
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0 / 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   3180
         Width           =   2055
      End
   End
   Begin VB.Timer tmrScan 
      Interval        =   5000
      Left            =   3600
      Top             =   6120
   End
   Begin VB.Timer tmrBorder 
      Interval        =   500
      Left            =   240
      Top             =   6120
   End
   Begin VB.Frame Frame2 
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   3975
      Begin VB.CommandButton cmdVote 
         Caption         =   "Vote for this code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   3735
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H00FFC4B3&
         BorderWidth     =   5
         Height          =   375
         Left            =   120
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label lblScan 
         AutoSize        =   -1  'True
         Caption         =   "Scanning ? ports per second"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2025
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "I would like to scan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtPort2 
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Text            =   "65535"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtPort1 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Text            =   "1"
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtHost 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "###.###.###.### -or- www.domain.com"
         Top             =   840
         Width           =   3735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Another Computer"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "My Own Computer"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   14
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Port Numbers:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cancel As Boolean
Dim lngScanned As Long
Dim Checked As Boolean
Public Sub Scan()
Cancel = False

Dim i As Long
Dim n As Long
Dim strHost As String
If Option1.Value = True Then
strHost = "127.0.0.1"
Else
strHost = txtHost
End If

ProgressBar1.Value = 0
ProgressBar1.Max = txtPort2
ProgressBar1.Min = txtPort1
tmrScan.Enabled = True
For i = txtPort1.Text To txtPort2.Text
If Cancel = True Then Command1.Tag = i: Exit Sub

    For n = 0 To 5000
    If Cancel = True Then Command1.Tag = i: Exit Sub
    
        If sckScan(n).State = sckClosed Then
        sckScan(n).Connect strHost, i
        DoEvents
        GoTo nextport
        End If
    
    Next n
nextport:
lngScanned = lngScanned + 1
ProgressBar1.Value = i
lblStatus.Caption = i & " / " & txtPort2.Text
DoEvents
Next i
tmrScan.Enabled = False
Exit Sub
End Sub


Public Sub Warning()
If Option1.Value = True Then
    If Checked = False Then
    SaveSetting App.Title, "Settings", "Check", True
    Checked = True
    frmRecommend.Show vbModal, Me
    End If
End If
End Sub

Private Sub cmdVote_Click()
frmVote.Show , Me
End Sub

Private Sub Command1_Click()
Dim lngTest As Long
lngTest = 1
On Error Resume Next

If txtPort1.Text + lngTest <= 1 Or txtPort1.Text + lngTest > 65536 Then
MsgBox "The first port entry textbox must contain a number from 1 to 65535", vbExclamation
txtPort1.SetFocus
Exit Sub
End If

If txtPort2.Text + lngTest <= 1 Or txtPort2.Text + lngTest > 65536 Then
MsgBox "The second port entry textbox must contain a number from 1 to 65535", vbExclamation
txtPort2.SetFocus
Exit Sub
End If

Option1.Enabled = False
Option2.Enabled = False
txtHost.Enabled = False
Command1.Enabled = False
Command2.Enabled = True
txtPort1.Enabled = False
txtPort2.Enabled = False
Cancel = False

If Option1.Value = True Then
lstPorts.AddItem "Scanning your computer"
Else
lstPorts.AddItem "Scanning " & txtHost.Text
End If
lstPorts.AddItem "From ports " & txtPort1 & " to " & txtPort2
lstPorts.AddItem "Scan started at " & Time & " on " & Date
lstPorts.AddItem "--------------------------------------------------------------------"
Scan
End Sub

Private Sub Command2_Click()
tmrScan.Enabled = False
lblScan.Caption = "Scanning ? ports per second"
lstPorts.AddItem "--------------------------------------------------------------------"
lstPorts.AddItem "Scan aborted at " & Time & " on port " & Command1.Tag
Option1.Enabled = True
Option2.Enabled = True
If Option2.Value = True Then
txtHost.Enabled = True
End If
Command2.Enabled = False
Command1.Enabled = True
txtPort1.Enabled = True
txtPort2.Enabled = True
Cancel = True
End Sub


Private Sub Command4_Click()
frmRecommend.Show vbModal, Me
End Sub

Private Sub Form_Load()
Dim i As Integer
Checked = GetSetting(App.Title, "Settings", "Check", False)
frmLoad.Show , Me
For i = 1 To 5000
Load sckScan(i)
frmLoad.pbProgress.Value = i
DoEvents
Next i
frmLoad.Hide


End Sub


Private Sub Option1_Click()
txtHost.Enabled = False
txtHost.Text = "###.###.###.### -or- www.domain.com"
End Sub

Private Sub Option2_Click()
txtHost.Enabled = True
txtHost.SetFocus
End Sub


Private Sub sckScan_Close(Index As Integer)
sckScan(Index).Close
End Sub

Private Sub sckScan_Connect(Index As Integer)
Warning
lstPorts.AddItem sckScan(Index).RemotePort
sckScan(Index).Close
End Sub

Private Sub sckScan_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
sckScan(Index).Close
End Sub








Private Sub tmrBorder_Timer()
shpBorder.Visible = Not shpBorder.Visible
End Sub


Private Sub tmrScan_Timer()
lblScan.Caption = "Scanning at " & lngScanned / 5 & " per second."
lngScanned = 0
End Sub

Private Sub txtHost_GotFocus()
txtHost.SelStart = 0
txtHost.SelLength = Len(txtHost)
End Sub


Private Sub txtPort1_GotFocus()
txtPort1.SelStart = 0
txtPort1.SelLength = Len(txtPort1)
End Sub


Private Sub txtPort2_GotFocus()
txtPort2.SelStart = 0
txtPort2.SelLength = Len(txtPort2)
End Sub


