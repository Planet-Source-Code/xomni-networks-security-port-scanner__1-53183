VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoad 
   BorderStyle     =   0  'None
   Caption         =   "Loading..."
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   5000
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "The application is loading, please wait..."
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   1080
         Width           =   2880
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Created by Xomni Networks"
         Height          =   195
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Security Port Scan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
