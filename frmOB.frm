VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmOB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AllDayDJ OB Server"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Timer tmrRecVol 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   1200
   End
   Begin MSWinsockLib.Winsock wsMain 
      Left            =   1560
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.Slider sldVolume 
      Height          =   1575
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   2778
      _Version        =   393216
      Orientation     =   1
      Max             =   100
      TickStyle       =   3
   End
   Begin VB.ComboBox cmbSource 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.ComboBox cmbSoundCard 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmOB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngRecordingChannel As Long         ' Recording channel
Dim byteRX(1 To 20000) As Byte          ' Recieved data
Dim byteTX(1 To 20000) As Byte          ' Transmitted data
Dim lngPlayoutChannel As Long           ' Playout channel
Dim boolTerminate As Boolean            ' Terminate the perpetual loop

Private Sub cmbSoundCard_Change()
Call LoadCard(cmbSoundCard.ListIndex)
End Sub

Private Sub cmbSoundCard_Click()
Call LoadCard(cmbSoundCard.ListIndex)
End Sub

Private Sub cmbSource_Change()
Call SelectSource(cmbSource.ListIndex)
End Sub

Private Sub cmbSource_Click()
Call SelectSource(cmbSource.ListIndex)
End Sub

Private Sub cmdClose_Click()

boolTerminate = True
wsMain.Close
Unload Me

End Sub

Private Sub Form_Load()

' Clear the lists

cmbSoundCard.Clear
cmbSource.Clear

' Setup the connection properties

wsMain.LocalPort = 4545

' List and initialize the soundcards

intLooper = 1
Do While BASS_GetDeviceDescription(intLooper)
    cmbSoundCard.AddItem VBStrFromAnsiPtr(BASS_GetDeviceDescription(intLooper))
    intLooper = intLooper + 1
Loop

' Do not terminate

boolTerminate = False

End Sub

Sub LoadCard(ByVal intCardNumber As Integer)

Dim intLooper As Integer
Dim intInputs As Integer

' Select the sound card

intCardNumber = intCardNumber + 1
Call BASS_SetDevice(intCardNumber)

' Get devices

cmbSource.Clear
intLooper = 0
While BASS_RecordGetInputName(intLooper) <> 0
    cmbSource.AddItem VBStrFromAnsiPtr(BASS_RecordGetInputName(intLooper))
    intLooper = intLooper + 1
Wend

' Set all levels to zero and select the first option

intInputs = intLooper

' Stop all inputs

For intLooper = -1 To (intInputs - 1)
    Call BASS_RecordSetInput(intLooper, BASS_INPUT_OFF)
Next intLooper

' Start the correct input

Call BASS_RecordSetInput(-1, BASS_INPUT_ON)

' Select the default device

If intLooper > 0 Then
cmbSource.ListIndex = 1
End If

End Sub

Sub SelectSource(ByVal intSelected As Integer)

Dim intLooper As Integer
Dim intInputs As Integer

' As we will be sent the item number from the combo box
' we need to remove 1 from the value

intSelected = intSelected - 1
intInputs = cmbSource.ListCount

' Set all levels to zero

For intLooper = -1 To (intInputs - 1)
    Call BASS_RecordSetInput(intLooper, BASS_INPUT_OFF)
Next intLooper

' Start the correct input

Call BASS_RecordSetInput(intSelected, BASS_INPUT_ON)
tmrRecVol.Enabled = True

End Sub

Private Sub tmrRecVol_Timer()

Dim intVolume As Integer
Dim intSource As Integer

' Set the recording volume

intSource = cmbSource.ListIndex
intVolume = 100 - sldVolume.value
Call BASS_RecordSetInput(intSource, BASS_INPUT_LEVEL Or intVolume)

End Sub

Private Sub wsMain_ConnectionRequest(ByVal requestID As Long)

Dim intLooper As Integer

' Close the old connection and open the new

wsMain.Close
wsMain.accept requestID

' Create the input and output audio channels

For intLooper = 1 To 20000
    byteTX(intLooper) = Null
    byteRX(intLooper) = Null
Next intLooper

' Start the recording
' Record to an array in mono 44100Hz quality
' This will give us good quality yet lower bandwidth

lngRecordingChannel = BASS_RecordStart(44100, 1, 0, 0, 0)

' Create the playout channel

lngPlayoutChannel = BASS_StreamCreateFile(BASSTRUE, byteRX(1), 0, 20000, 0)
Call BASS_ChannelPlay(lngPlayoutChannel, 1)

' Start the transmission

Call TransSub

End Sub

Sub TransSub()

Dim strData As String
ReDim byteTX(1 To 20000) As Byte

' This is a perpetual loop with a getout clause

StartPoint:
DoEvents

' Obtain the data from the recording source

Call BASS_ChannelGetData(lngRecordingChannel, byteTX(), 20000)
strData = "ADUI"
For intLooper = 1 To 20000
    strData = strData & byteTX(intLooper)
    byteTX(intLooper) = Null
Next intLooper

' Transmit the data from the recording source

wsMain.SendData strData

' Transmit the extra data

strData = "TMRR" & main.lblTimeRemaining.Caption
wsMain.SendData strData
strData = "TMRP" & main.lblTimePlayed.Caption
wsMain.SendData strData
strData = "BRKA" & main.chkBreakAfter.value
wsMain.SendData strData
strData = "NOWD" & main.txtArtist.text & " - " & main.txtSong.text
wsMain.SendData strData
strData = "NEXT" & main.lstPlaylist.ListItems.Item(1).text & " - " & main.lstPlaylist.ListItems.Item(1).SubItems(1)
wsMain.SendData strData

If main.lblTimeRemaining.ForeColor = vbBlack Then
strData = "TRMO1"
Else
strData = "TRMO0"
End If
wsMain.SendData strData

' Loop round

If boolTerminate = False Then
    GoTo StartPoint
End If

End Sub

Private Sub wsMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Transmission/Reception Error Occurred:" & vbCrLf & vbCrLf & Description, vbOKOnly + vbInformation, "Error"
Unload Me
End Sub
