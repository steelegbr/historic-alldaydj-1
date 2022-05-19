VERSION 5.00
Begin VB.Form frmAdminOptions 
   Caption         =   "Admin Options"
   ClientHeight    =   5490
   ClientLeft      =   4605
   ClientTop       =   2685
   ClientWidth     =   2355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   2355
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdLevel 
      Caption         =   "Fade Level"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton cmdAGC 
      Caption         =   "AGC"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdSoundCards 
      Caption         =   "Sound Cards"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdNetworking 
      Caption         =   "Networking"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdRotation 
      Caption         =   "Rotation"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdPlaylist 
      Caption         =   "Schedule Playlists"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmdAdvertisements 
      Caption         =   "Advertisements"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdNewsSettings 
      Caption         =   "News / Hourly Settings"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdStationDetails 
      Caption         =   "Station / Admin Settings"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit Database Entry"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdRemoveFromDB 
      Caption         =   "Remove From Database"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdAddToDB 
      Caption         =   "Add To Database"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cmdViewLogs 
      Caption         =   "View Logs"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmAdminOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAddToDB_Click()

' Display the screen and initialise variables

frmAddToDB.Visible = True
frmAddToDB.lblFileLocation.Caption = "NO FILE SELECTED"
frmAddToDB.lblStart.Caption = "0"
frmAddToDB.lblEnd.Caption = "0"
frmAddToDB.lblIntro.Caption = "0"
frmAddToDB.lblCurrentPosition.Caption = "00:00"
frmAddToDB.txtArtist.text = ""
frmAddToDB.txtTrack.text = ""

End Sub

Private Sub cmdAdvertisements_Click()

' Show the advertisements form

frmAdverts.Visible = True

End Sub

Private Sub cmdEdit_Click()

' Show edit search

frmSearchEdit.Visible = True
Unload Me

End Sub

Private Sub cmdExit_Click()

' Make this form and other admin options invisible

Me.Visible = False

End Sub

Private Sub cmdLevel_Click()

    frmFadeLevel.Visible = True

End Sub

Private Sub cmdNetworking_Click()

' Show networking form

frmNetworking.Visible = True

End Sub

Private Sub cmdNewsSettings_Click()

frmNewsV2.Visible = True

End Sub

Private Sub cmdAGC_Click()

' Show the AGC screen

frmAGC.Visible = True
Unload Me

End Sub

Private Sub cmdPlaylist_Click()

' Show the schedule playlists screen

frmPlaylists.Visible = True

End Sub

Private Sub cmdRemoveFromDB_Click()

' Show the screen

frmRemoveFromDBSearch.Visible = True

End Sub

Private Sub cmdRotation_Click()

' Load the form

Me.Visible = False
frmRotation.Visible = True

End Sub

Private Sub cmdSoundCards_Click()

' Show the screen

frmSoundCard.Visible = True
Unload Me

End Sub

Private Sub cmdStationDetails_Click()

' Load the form and data

frmAdminSettings.tmrLoadData.Enabled = True
frmAdminSettings.Visible = True

End Sub

Private Sub cmdViewLogs_Click()

' Make form visible and load the data

frmLogViewer.Visible = True
End Sub

