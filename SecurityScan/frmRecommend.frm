VERSION 5.00
Begin VB.Form frmRecommend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "We recommend you install a firewall!"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6045
   Icon            =   "frmRecommend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Thank You"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   $"frmRecommend.frx":5D52
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5640
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   $"frmRecommend.frx":5E15
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5730
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRecommend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Hide
End Sub

