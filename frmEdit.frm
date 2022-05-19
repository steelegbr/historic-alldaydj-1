VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEdit 
   Caption         =   "Edit Entry"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtplayer 
      Height          =   285
      Left            =   240
      TabIndex        =   24
      Text            =   "0"
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   255
      Left            =   960
      TabIndex        =   22
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdTestIntro 
      Caption         =   "Test"
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdSetIntro 
      Caption         =   "Intro"
      Height          =   255
      Left            =   3720
      TabIndex        =   18
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdSetEnd 
      Caption         =   "End"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Timer tmrGetCurrentPosition 
      Interval        =   250
      Left            =   5280
      Top             =   1920
   End
   Begin VB.CommandButton cmdTestEnd 
      Caption         =   "Test"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CheckBox chkFades 
      Alignment       =   1  'Right Justify
      Caption         =   "Fades?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtComposer 
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox txtCompany 
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txtTrack 
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtArtist 
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmdSetStart 
      Caption         =   "Start"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdTestStart 
      Caption         =   "Test"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin MSComctlLib.Slider sldTimer 
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      Caption         =   "0:00"
      Height          =   255
      Left            =   3720
      TabIndex        =   20
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      Caption         =   "0:00"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      Caption         =   "0:00"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblCurrentPosition 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   3720
      TabIndex        =   14
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblComposer 
      Alignment       =   1  'Right Justify
      Caption         =   "Composer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblRecordCompany 
      Alignment       =   1  'Right Justify
      Caption         =   "Record Company:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label lblTrack 
      Alignment       =   1  'Right Justify
      Caption         =   "Track:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblArtist 
      Alignment       =   1  'Right Justify
      Caption         =   "Artist:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2160
      Width           =   615
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdPlay_Click()

Dim lngChannel As Long

lngChannel = txtplayer.text
Call BASS_SetDevice(GetSoundCard(PFL_SOUND_CARD))
Call BASS_ChannelPlay(txtplayer.text, BASSFALSE)
Call BASS_FX_DSP_Reset(lngChannel, BASS_FX_DSPFX_DAMP)
Call BASS_ChannelSetPosition(lngChannel, 0)
Call BASS_ChannelSetAttributes(lngChannel, -1, 100, -101)

End Sub

Private Sub cmdSave_Click()

Dim intLooper As Integer
Dim intLooper2 As Integer
Dim dbConnection As New ADODB.Connection
Dim strConnectionString As String
Dim strSQL As String
Dim dblLength As Double
Dim dblStart As Double
Dim dblEnd As Double

' Check all the text data is valid

If Len(txtArtist.text) = 0 Then
Call Error_Handler("Artist name not long enough.", 1)
End If

If Len(txtTrack.text) = 0 Then
Call Error_Handler("Artist name not long enough.", 1)
End If

If Len(txtCompany.text) = 0 Then
Call Error_Handler("Record company name not long enough.", 1)
End If

If Len(txtComposer.text) = 0 Then
Call Error_Handler("Composer name not long enough.", 1)
End If

' Check that all the number data is valid

dblStart = lblStart.Caption
dblEnd = lblEnd.Caption
dblLength = dblEnd - dblStart

If dblLength < 0 Then
Call Error_Handler("Cannot have start time after the finish time.", 1)
Exit Sub
End If

If dblLength = 0 Then
Call Error_Handler("Cannot have start time at the same time as the finish time.", 1)
Exit Sub
End If

' Connect

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb;Persist Security Info=False"
dbConnection.Open strConnectionString

' Create the SQL command and execute

strSQL = "UPDATE " & LCase$(frmSearchEdit.cmbType.text) & " SET [artist] = """ & txtArtist.text & """, [track] = """ & txtTrack.text & """, [length] = " & dblLength & ", [Start] = " & lblStart.Caption & ", [finish] = " & lblEnd.Caption & ", [fades] = " & chkFades.value & ", [record_company] = """ & txtCompany.text & """, [composer] = """ & txtComposer.text & """, [countdown] = " & lblIntro.Caption & " WHERE [itemid] = " & txtID.text
dbConnection.Execute strSQL
dbConnection.Close

' Hide this screen and show admin menu

txtArtist.text = ""
txtTrack.text = ""
txtCompany.text = ""
txtComposer.text = ""
Call cmdStop_Click
frmAdminOptions.Visible = True
Unload Me

End Sub

Private Sub cmdSetEnd_Click()

' Get the current position and assign it to the end

lblEnd.Caption = BASS_ChannelGetPosition(txtplayer.text)
lblEnd.Caption = BASS_ChannelBytes2Seconds(txtplayer.text, lblEnd.Caption)

End Sub

Private Sub cmdSetIntro_Click()

' Get the current position and assign it to the intro

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

Private Sub tmrGetCurrentPosition_Timer()

' Initiate the variables

Dim dblPosition As Double
Dim intPositionSeconds As Integer
Dim intPositionMinutes As Integer
Dim strPositionSeconds As String
Dim strPositionMinutes As String

' Start the infinite loop that gets the current song position


Call BASS_SetDevice(GetSoundCard(PFL_SOUND_CARD))
dblPosition = BASS_ChannelGetPosition(txtplayer.text)
dblPosition = BASS_ChannelBytes2Seconds(txtplayer.text, dblPosition)
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

sldTimer.value = dblPosition

End Sub

Private Sub sldTimer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

tmrGetCurrentPosition.Enabled = False

End Sub

Private Sub sldTimer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

tmrGetCurrentPosition.Enabled = True

End Sub
