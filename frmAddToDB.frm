VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddToDB 
   Caption         =   "Add To Database"
   ClientHeight    =   4725
   ClientLeft      =   3945
   ClientTop       =   2685
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtplayer 
      Height          =   285
      Left            =   240
      TabIndex        =   28
      Text            =   "0"
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   1200
      TabIndex        =   27
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   255
      Left            =   600
      TabIndex        =   26
      Top             =   1080
      Width           =   615
   End
   Begin MSComctlLib.Slider sldTimer 
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSetIntro 
      Caption         =   "Intro"
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdTestIntro 
      Caption         =   "Test"
      Height          =   255
      Left            =   1920
      TabIndex        =   22
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtRecordCompany 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox txtComposer 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CheckBox chkFade 
      Alignment       =   1  'Right Justify
      Caption         =   "Fade?"
      BeginProperty DataFormat 
         Type            =   5
         Format          =   ""
         HaveTrueFalseNull=   1
         TrueValue       =   "True"
         FalseValue      =   "False"
         NullValue       =   ""
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   7
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdTestEnd 
      Caption         =   "Test"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdTestStart 
      Caption         =   "Test"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.Timer tmrGetCurrentPosition 
      Interval        =   1000
      Left            =   3480
      Top             =   1800
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   4200
      Width           =   975
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmAddToDB.frx":0000
      Left            =   2040
      List            =   "frmAddToDB.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox txtTrack 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txtArtist 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdSetEnd 
      Caption         =   "End"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdSetStart 
      Caption         =   "Start"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cdlCommon 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      Caption         =   "0:00"
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblRecordCompany 
      Alignment       =   1  'Right Justify
      Caption         =   "Record Company:"
      Height          =   255
      Left            =   480
      TabIndex        =   21
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblComposer 
      Alignment       =   1  'Right Justify
      Caption         =   "Composer:"
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblCurrentPosition 
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   19
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblFileLocation 
      Alignment       =   2  'Center
      Caption         =   "NO FILE SELECTED"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      Caption         =   "0:00"
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      Caption         =   "0:00"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      Caption         =   "Type:"
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label lblTrack 
      Alignment       =   1  'Right Justify
      Caption         =   "Track"
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblArtist 
      Alignment       =   1  'Right Justify
      Caption         =   "Artist:"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "frmAddToDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

' Initiate variables

Dim dbConnection As New ADODB.Connection
Dim strConnectionString As String
Dim strSQL As String
Dim strFileLocation As String
Dim dblStart As Double
Dim dblEnd As Double
Dim dblLength As Double
Dim strTrack As String
Dim strArtist As String
Dim strType As String
Dim boolFade As Boolean
Dim strMessage As String
Dim strRecordCompany As String
Dim strComposer As String
Dim dblIntro As Double

' Create connection string

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb"

' Connect

dbConnection.Open strConnectionString

' Create the command

strFileLocation = lblFileLocation.Caption

dblStart = lblStart.Caption
dblEnd = lblEnd.Caption
dblLength = dblEnd - dblStart
dblIntro = lblIntro.Caption

strArtist = txtArtist.text
strTrack = txtTrack.text
strType = LCase$(cmbType.text)
strComposer = txtComposer.text
strRecordCompany = txtRecordCompany.text

boolFade = chkFade

' Prepare in case invalid times have been entered

If dblStart >= dblEnd Then
MsgBox "The end time is earlier than the start time. Please enter a later end time.", vbCritical
dbConnection.Close
Exit Sub
End If

' Check that a type has been selected

If strType = "" Then
MsgBox "Please select a track type in the drop-down menu.", vbCritical
dbConnection.Close
Exit Sub
End If

' Prepare for names that are too long

If Len(strArtist) > 65535 Then
strArtist = Left$(strArtist, 63335)
strMessage = "The artist name you entered was too long." & Chr(10) & Chr(10) & "It has been trimmed to:" & Chr(10) & Chr(10) & strArtist
Call Error_Handler(strMessage, 1)
End If

If Len(strTrack) > 63335 Then
strArtist = Left$(strTrack, 63335)
strMessage = "The track name you entered was too long." & Chr(10) & Chr(10) & "It has been trimmed to:" & Chr(10) & Chr(10) & strTrack
Call Error_Handler(strMessage, 1)
End If

If Len(strComposer) > 63335 Then
strArtist = Left$(strArtist, 63335)
strMessage = "The composer name you entered was too long." & Chr(10) & Chr(10) & "It has been trimmed to:" & Chr(10) & Chr(10) & strArtist
Call Error_Handler(strMessage, 1)
End If

If Len(strRecordCompany) > 63335 Then
strArtist = Left$(strTrack, 63335)
strMessage = "The record company name you entered was too long." & Chr(10) & Chr(10) & "It has been trimmed to:" & Chr(10) & Chr(10) & strTrack
Call Error_Handler(strMessage, 1)
End If

' Prepare for names that are too short

If Len(strArtist) = 0 Then
strMessage = "The artist name you entered was too short." & Chr(10) & Chr(10) & "Please enter a longer name."
Call Error_Handler(strMessage, 1)
Exit Sub
End If

If Len(strTrack) = 0 Then
strMessage = "The track name you entered was too short." & Chr(10) & Chr(10) & "Please enter a longer name."
Call Error_Handler(strMessage, 1)
Exit Sub
End If

If Len(strComposer) = 0 Then
strMessage = "The composer name you entered was too short." & Chr(10) & Chr(10) & "Please enter a longer name."
Call Error_Handler(strMessage, 1)
Exit Sub
End If

If Len(strRecordCompany) = 0 Then
strMessage = "The record company name you entered was too short." & Chr(10) & Chr(10) & "Please enter a longer name."
Call Error_Handler(strMessage, 1)
Exit Sub
End If

' Check the intro is not too long

If dblIntro >= dblLength Then
Call Error_Handler("Intro is longer than the actual track." & Chr(10) & Chr(10) & "Please select a shoter intro.", 1)
Exit Sub
End If

' Prevent errors with voice tracking

If strType = "voice_tracks" Then
boolFade = True
End If

' Create the SQL command

strSQL = "INSERT INTO " & strType & " ([artist], [track], [length], [start], [finish], [file], [fades], [record_company], [composer], [extrainfo], [countdown]) VALUES(""" & strArtist & """, """ & strTrack & """, """ & dblLength & """, """ & dblStart & """, """ & dblEnd & """, """ & strFileLocation & """, " & boolFade & ", """ & strRecordCompany & """, """ & strComposer & """, ""Add your own information here and click save."", " & dblIntro & ")"

' Execute the SQL command and close the connection

dbConnection.Execute strSQL
dbConnection.Close
Call cmdStop_Click
Call BASS_SetDevice(GetSoundCard(PFL_SOUND_CARD))
Me.Visible = False
MsgBox "Track added to database.", vbInformation, "Track Added"

End Sub

Private Sub cmdBrowse_Click()

' Initiate Variables

Dim strFileToLoad As String
Dim strFileTitle As String
Dim intFileTitleLength As Integer
Dim dblTrackEnd As Double
Dim strArtist As String
Dim strTrack As String
Dim vntFileTitleSplit As Variant
Dim lngChannel As Long
Dim vntEnd As Variant

' Stop previous loads

Call cmdStop_Click

' Use common controls for opening file

On Error GoTo Handler

'cdlCommon.InitDir = App.Path
cdlCommon.Filter = "MP3 Files (*.mp3)|*.mp3|" & _
                   "WMA Files (*.wma)|*.wma|" & _
                   "WAV Files (*.wav)|*.wav|" & _
                   "OGG Files (*.ogg)|*.ogg|" & _
                   "AIFF Files (*.aiff)|*.aiff|"
cdlCommon.CancelError = True
cdlCommon.ShowOpen
strFileToLoad = cdlCommon.FileName
strFileTitle = cdlCommon.FileTitle
If Len(strFileTitle) > 4 Then
strFileTitle = Left$(strFileTitle, (Len(strFileTitle) - 4))
End If

' Load the track

Call BASS_SetDevice(GetSoundCard(PFL_SOUND_CARD))
If UCase$(Right$(strFileToLoad, 3)) = "WMA" Then
lngChannel = BASS_WMA_StreamCreateFile(BASSFALSE, strFileToLoad, 0, 0, BASS_STREAM_PRESCAN)
txtplayer.text = lngChannel
Else
lngChannel = BASS_StreamCreateFile(BASSFALSE, strFileToLoad, 0, 0, BASS_STREAM_PRESCAN)
txtplayer.text = lngChannel
End If

vntEnd = BASS_ChannelBytes2Seconds(lngChannel, BASS_ChannelGetLength(lngChannel))

' Show other data

lblEnd.Caption = vntEnd
sldTimer.max = vntEnd
lblFileLocation = strFileToLoad
txtTrack.text = strFileTitle
lblStart.Caption = 0

' Start the AGC and load global settings

Call BASS_FX_DSP_Set(lngChannel, BASS_FX_DSPFX_DAMP, 100)
Call BASS_FX_DSP_SetParameters(lngChannel, BASS_FX_DSPFX_DAMP, GlobalAGC)

' Error handling

Handler:

End Sub




Private Sub cmdCancel_Click()

' Clear all options and close screen

lblFileLocation.Caption = "NO FILE SELECTED"
lblStart.Caption = "0"
lblEnd.Caption = "0"
lblIntro.Caption = "0"
lblCurrentPosition.Caption = "00:00"
txtArtist.text = ""
txtTrack.text = ""
Call cmdStop_Click
Me.Visible = False

End Sub

Private Sub cmdPlay_Click()

Dim lngChannel As Long

lngChannel = txtplayer.text
Call BASS_SetDevice(GetSoundCard(PFL_SOUND_CARD))
Call BASS_ChannelPlay(txtplayer.text, BASSFALSE)
Call BASS_FX_DSP_Reset(lngChannel, BASS_FX_DSPFX_DAMP)
Call BASS_ChannelSetPosition(lngChannel, 0)
Call BASS_ChannelSetAttributes(lngChannel, -1, 100, -101)

End Sub

Private Sub cmdSetEnd_Click()

' Get the current position and assign it to the end

lblEnd.Caption = BASS_ChannelGetPosition(txtplayer.text)
lblEnd.Caption = BASS_ChannelBytes2Seconds(txtplayer.text, lblEnd.Caption)

End Sub

Private Sub cmdSetIntro_Click()

lblIntro.Caption = BASS_ChannelGetPosition(txtplayer.text)
lblIntro.Caption = BASS_ChannelBytes2Seconds(txtplayer.text, lblIntro.Caption)

End Sub

Private Sub cmdSetStart_Click()

' Get the current position and assign it to the start

lblStart.Caption = BASS_ChannelGetPosition(txtplayer.text)
lblStart.Caption = BASS_ChannelBytes2Seconds(txtplayer.text, lblStart.Caption)

End Sub


Private Sub cmdStop_Click()

Call BASS_SetDevice(GetSoundCard(PFL_SOUND_CARD))
BASS_ChannelStop (txtplayer.text)
Call BASS_ChannelSetPosition(txtplayer.text, 0)

End Sub

Private Sub cmdTestEnd_Click()

' Jump to position

Dim lngChannel As Long

lngChannel = txtplayer.text
Call BASS_ChannelSetPosition(lngChannel, BASS_ChannelSeconds2Bytes(lngChannel, lblEnd.Caption))

End Sub

Private Sub cmdTestIntro_Click()

' Jump to position

Dim lngChannel As Long

lngChannel = txtplayer.text
Call BASS_ChannelSetPosition(lngChannel, BASS_ChannelSeconds2Bytes(lngChannel, lblIntro.Caption))

End Sub

Private Sub cmdTestStart_Click()

' Jump to position

Dim lngChannel As Long

lngChannel = txtplayer.text
Call BASS_ChannelSetPosition(lngChannel, BASS_ChannelSeconds2Bytes(lngChannel, lblStart.Caption))

End Sub

Private Sub Form_Load()

' Start the repetitive task

tmrGetCurrentPosition.Enabled = True

End Sub

Private Sub sldTimer_Click()

Dim lngChannel As Long

lngChannel = txtplayer.text
Call BASS_ChannelSetPosition(lngChannel, BASS_ChannelSeconds2Bytes(lngChannel, sldTimer.value))

End Sub

Private Sub sldTimer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

tmrGetCurrentPosition.Enabled = False

End Sub

Private Sub sldTimer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

tmrGetCurrentPosition.Enabled = True

End Sub

Private Sub tmrGetCurrentPosition_Timer()

' Initiate the variables

Dim dblPosition As Double
Dim intPositionSeconds As Integer
Dim intPositionMinutes As Integer
Dim strPositionSeconds As String
Dim strPositionMinutes As String
Dim dblTrackEnd As Double

' Start the infinite loop that gets the current song position

Call BASS_SetDevice(GetSoundCard(PFL_SOUND_CARD))
dblPosition = BASS_ChannelGetPosition(txtplayer.text)
dblPosition = BASS_ChannelBytes2Seconds(txtplayer.text, dblPosition)
sldTimer.value = dblPosition
If dblPosition < 0 Then
dblPosition = 0
End If
intPositionMinutes = Int(dblPosition / 60)
intPositionSeconds = Int(dblPosition Mod 60)

If intPositionMinutes < 10 Then
strPositionMinutes = "0" & intPositionMinutes
Else
strPositionMinutes = intPositionMinutes
End If

If intPositionSeconds < 10 Then
strPositionSeconds = "0" & intPositionSeconds
Else
strPositionSeconds = intPositionSeconds
End If

lblCurrentPosition = strPositionMinutes & ":" & strPositionSeconds

End Sub
