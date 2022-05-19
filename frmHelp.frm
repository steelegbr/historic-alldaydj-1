VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SHDocVwCtl.WebBrowser webHelp 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      ExtentX         =   18018
      ExtentY         =   9975
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
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   9480
      TabIndex        =   1
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   5880
      Width           =   855
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()

On Error Resume Next

webHelp.GoBack

End Sub

Private Sub cmdClose_Click()

Unload Me

End Sub
