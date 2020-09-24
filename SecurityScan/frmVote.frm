VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmVote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Votes and Feedback Appreciated!"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7650
   Icon            =   "frmVote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      ExtentX         =   13573
      ExtentY         =   11668
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
WebBrowser.Navigate "http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=53183&lngWId=1"
End Sub

Private Sub WebBrowser_StatusTextChange(ByVal Text As String)

End Sub


