Attribute VB_Name = "AllDayDJModule"
' Set the soundcard constants

Global Const MAIN_SOUND_CARD = 1
Global Const PFL_SOUND_CARD = 2

Public boolStarted As Boolean
Public InstantPlayers(0 To 8) As Long
Public MainPlayers(0 To 3) As Long
Public IsFading(0 To 3) As Boolean
Public GlobalAGC As BASS_FX_DSPDAMP
Public GlobalCompressor As BASS_FX_DSPCOMPRESSOR
Public GlobalVolume As clsVolume
Public gblFeed As String
Public boolUseFeed As Boolean
Public RotTimes As Integer
Public GlobalRotPos As Integer
Public GlobalPlayer As Integer
Public GlobalDate As String
Public FadeLevel As Integer

' Logged In?

Public AdminLoggedIn As Boolean

Sub Error_Handler(ByVal strMessage As String, ByVal intType As Integer)

' strMessage holds the message to be displayed to the user
' intType is the type of error message to be shown
' 0 = Critcal
' 1 = Information

If intType = 0 Then
MsgBox strMessage, vbOKOnly + vbCritical, "Critical Error"
End If

If intType = 1 Then
MsgBox strMessage, vbOKOnly + vbInformation, "Error"
End If

End Sub

Public Function AppPath()

Dim strLocation As String
Dim strPathOnFile As String

' Open network.txt file

strLocation = App.Path & "\networking.dat"
Open strLocation For Input As #20
Line Input #20, strPathOnFile
Close #20

' Check the value
' If it is App.Path then retrun App.Path
' Else return normal path

If strPathOnFile = "App.Path" Then
AppPath = App.Path
Else
AppPath = strPathOnFile
End If

End Function

Sub doAudioWall()

Dim intItems As Integer
Dim intLooper As Integer
Dim intLooper2 As Integer
Dim strText As String
Dim strTemp As String

' Retrieve the number of items

intItems = frmSearchResults.lstResults.ListItems.Count

' If it is zero then show a message and hide this screen

If intItems = 0 Then
Unload frmAudioWall
MsgBox "No results.", vbOKOnly + vbInformation, "Search Results"
frmSearch.Visible = True
Exit Sub
End If

' If less than 20 then hide other search boxes

If intItems < 20 Then
For intLooper = intItems To 19
frmAudioWall.cmdItem(intLooper).Visible = False
Next intLooper
End If

If intItems > 20 Then
intItems = 20
End If

' Load the items


For intLooper = 1 To intItems
strText = frmSearchResults.lstResults.ListItems.item(intLooper).SubItems(1) & Chr(10) & frmSearchResults.lstResults.ListItems.item(intLooper).SubItems(2)
For intLooper2 = 1 To Len(strText)
strTemp = Left$(strText, intLooper2)
If Right$(strTemp, 1) = "&" Then
strText = strTemp & "&" & Right$(strText, Len(strText) - intLooper2)
intLooper2 = intLooper2 + 1
End If
Next intLooper2
frmAudioWall.cmdItem(intLooper - 1).Caption = strText
Next intLooper

' Prepare the page

If intItems < 20 Then
frmAudioWall.cmdNext.Visible = False
End If
frmAudioWall.lblPage.Caption = "1"


End Sub

Sub doAddToPlaylist(ByVal intID As Integer, ByVal strType As String)

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strConnectionString As String
Dim vntData(1 To 13) As Variant
Dim strSQL As String
Dim intCounter As Integer
Dim x As Variant

' Prepare for scheduler bug

If Left$(strType, 1) = """" Then
    strType = Right$(strType, Len(strType) - 1)
End If

If Right$(strType, 1) = """" Then
    strType = Left$(strType, Len(strType) - 1)
End If

' Connect to the database

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb"
dbConnection.Open strConnectionString
strSQL = "SELECT * FROM " & strType & " WHERE [itemid] = " & intID
dbRecordset.Open strSQL, dbConnection

' Retrieve the data

intCounter = 0
Do While Not dbRecordset.EOF
For Each x In dbRecordset.Fields
intCounter = intCounter + 1
vntData(intCounter) = x
Next x
dbRecordset.MoveNext
Loop

vntData(13) = strType

' Remove #'s if they exist

If Left$(vntData(7), 1) = "#" Then
    vntData(7) = Right$(vntData(7), Len(vntData(7)) - 1)
End If
If Right$(vntData(7), 1) = "#" Then
    vntData(7) = Left$(vntData(7), Len(vntData(7)) - 1)
End If

' Place the data into the list

With frmPlayers.lstPlaylist.ListItems.Add(, , vntData(1))
For intCounter = 2 To 13
.SubItems(intCounter - 1) = vntData(intCounter)
Next intCounter
End With

' Place it into the visible playlist

With main.lstPlaylist.ListItems.Add(, , vntData(2))
.SubItems(1) = vntData(3)
End With

' Pre load next item

Call Pre_Load

End Sub


Sub Pre_Load()

Dim intPlayer As Integer
Dim strFileName As String
Dim dblIntro As Double
Dim lngChannel As Long

' Show the countdown timer

If frmPlayers.lstPlaylist.ListItems.Count > 0 Then
dblIntro = frmPlayers.lstPlaylist.ListItems.item(1).SubItems(11)
dblIntro = Int(dblIntro - frmPlayers.lstPlaylist.ListItems.item(1).SubItems(4))
main.lblRemaining_Time.Caption = dblIntro
main.lblRemaining_Time.BackColor = vbGreen
main.lblRemaining_Time.ForeColor = vbBlack
main.lblTimeRemaining.Caption = "Intro Countdown"
Else
main.lblRemaining_Time.BackColor = vbBlack
main.lblRemaining_Time.ForeColor = vbRed
main.lblTimeRemaining.Caption = "Time Remaining"
main.lblRemaining_Time.Caption = "00:00"
End If

' Clear the stream

intPlayer = GlobalPlayer
Select Case intPlayer
Case 0
intPlayer = 1
Case 1
intPlayer = 2
Case 2
intPlayer = 3
Case 3
intPlayer = 0
End Select

lngChannel = MainPlayers(intPlayer)
Call BASS_StreamFree(lngChannel)

End Sub

' Get the soundcard data from the text file and return it

Function GetSoundCard(intType As Integer) As Integer

Dim strFileName As String
Dim strData(1 To 2) As String

' Open the file, retrieve the data and close

strFileName = App.Path & "\soundcard.dat"
Open strFileName For Input As #1
Line Input #1, strData(1)
Line Input #1, strData(2)
Close #1

' Return the value

GetSoundCard = strData(intType)

End Function

' Bass.dll Routines
' Supplied with Bass.dll SDK

'check if any file exists
Public Function FileExists(ByVal fp As String) As Boolean
    FileExists = (Dir(fp) <> "")
End Function

' RPP = Return Proper Path
Public Function RPP(ByVal fp As String) As String
    RPP = IIf(Mid(fp, Len(fp), 1) = "\", fp, fp & "\")
End Function

'get file name from file path
Public Function GetFileName(ByVal fp As String) As String
    GetFileName = Mid(fp, InStrRev(fp, "\") + 1)
End Function

' Retrieve AGC settings

Public Sub GetAGCSettings()
    
    Dim strFileName As String
    Dim intFreeFile As Integer
    
    ' Open the file
    
    strFileName = RPP(App.Path) & "agc.dat"
    intFreeFile = FreeFile
    Open strFileName For Random As intFreeFile Len = Len(GlobalAGC)
    
    ' Retrieve the data
    
    Get intFreeFile, 1, GlobalAGC
    
    ' Close the file
    
    Close intFreeFile
    
    ' Set the AGC values
    
    'GlobalAGC.fTarget = 0.92
    'GlobalAGC.fQuiet = 0.02
    'GlobalAGC.fRate = 0.01
    'GlobalAGC.fDelay = 0.5
    'GlobalAGC.fGain = 1#
    
    ' Set the compressor settings
    
    GlobalCompressor.fThreshold = 0.8
    GlobalCompressor.fAttacktime = 1#
    GlobalCompressor.fReleasetime = 500#
    
End Sub

Public Sub ChangeVol(intUpDown As Integer)

    ' Change the news feed volume
    ' Feed names based on english sound card settings
    ' intUpDown: 0 = UP, 1 = DOWN
    
    If boolUseFeed = False Then
        Exit Sub
    End If
    
    If intUpDown = 0 Then
        Select Case gblFeed
            Case "CD"
                GlobalVolume.CD_Volume = GlobalVolume.MaxCDVolume
            Case "LINE IN"
                GlobalVolume.LineInVolume = GlobalVolume.MaxLineInVolume
            Case "MIC"
                GlobalVolume.MicroVolume = GlobalVolume.MaxMicVolume
            Case "MIDI"
                GlobalVolume.MIDIVolume = GlobalVolume.MaxMidVolume
            Case "WAVE IN"
                GlobalVolume.WaveInVolume = GlobalVolume.MaxWavInVolume
            Case "WAVE"
                GlobalVolume.WaveVolume = GlobalVolume.MaxWaveVolume
        End Select
    Else
        Select Case gblFeed
            Case "CD"
                GlobalVolume.CD_Volume = GlobalVolume.MinCDVolume
            Case "LINE IN"
                GlobalVolume.LineInVolume = GlobalVolume.MinLineInVolume
            Case "MIC"
                GlobalVolume.MicroVolume = GlobalVolume.MinMicVolume
            Case "MIDI"
                GlobalVolume.MIDIVolume = GlobalVolume.MinMidVolume
            Case "WAVE IN"
                GlobalVolume.WaveInVolume = GlobalVolume.MinWavInVolume
            Case "WAVE"
                GlobalVolume.WaveVolume = GlobalVolume.MinWaveVolume
        End Select
    End If

End Sub

Public Sub SetEq(lngChannel As Long)

    ' Adapted from BASS_FX.DLL example

    Dim eq As BASS_FX_DSPPEAKEQ
    Dim intLooper As Integer

    ' Get current sample rate of a Handle
    Dim f As Long
    Call BASS_ChannelGetAttributes(lngChannel, f, 0, 0)
    
    ' Set the EQ effect
    Call BASS_FX_DSP_Set(lngChannel, BASS_FX_DSPFX_PEAKEQ, 0)
    Call BASS_FX_DSP_Set(lngChannel, BASS_FX_DSPFX_PEAKEQ, 0)
    Call BASS_FX_DSP_Set(lngChannel, BASS_FX_DSPFX_PEAKEQ, 0)
    
    eq.lFreq = f
    eq.fBandwidth = 2.5
    eq.fQ = 0#
    eq.fGain = 0#
    
    ' set bass
    eq.lBand = 0
    eq.fCenter = 125
    Call BASS_FX_DSP_SetParameters(lngChannel, BASS_FX_DSPFX_PEAKEQ, eq)
    
    ' set mid
    eq.lBand = 1
    eq.fCenter = 1000
    Call BASS_FX_DSP_SetParameters(lngChannel, BASS_FX_DSPFX_PEAKEQ, eq)
    
    ' set treble
    eq.lBand = 2
    eq.fCenter = 8000
    Call BASS_FX_DSP_SetParameters(lngChannel, BASS_FX_DSPFX_PEAKEQ, eq)
    
    ' update dsp eq
    
    For intLooper = 1 To 3
        eq.lBand = intLooper
        Call BASS_FX_DSP_GetParameters(lngChannel, BASS_FX_DSPFX_PEAKEQ, eq)
        eq.fGain = 10
        Call BASS_FX_DSP_SetParameters(lngChannel, BASS_FX_DSPFX_PEAKEQ, eq)
    Next intLooper

End Sub

Public Sub Set_Compressor(lngChannel As Long)

    Call BASS_FX_DSP_Set(lngChannel, BASS_FX_DSPFX_COMPRESSOR, 100)
    Call BASS_FX_DSP_SetParameters(lngChannel, BASS_FX_DSPFX_COMPRESSOR, GlobalCompressor)
    
End Sub

Public Function SecondsToMins(intSecs As Integer) As String

    Dim intMins As Integer
    
    ' Zero the minutes
    
    intMins = 0
    
    ' Count up the minutes until seconds less than 60
    
    Do While intSecs >= 60
        intMins = intMins + 1
        intSecs = intSecs - 60
    Loop
    
    ' Create the string
    
    If intMins = 0 Then
        SecondsToMins = ""
    Else
        SecondsToMins = intMins & ":"
    End If
    
    If intSecs < 10 And intMins > 0 Then
        SecondsToMins = SecondsToMins & "0" & intSecs
    ElseIf intSecs < 10 And intMins = 0 Then
        SecondsToMins = SecondsToMins & intSecs
    Else
        SecondsToMins = SecondsToMins & intSecs
    End If

End Function

Public Function Slash(strLocation As String) As String

    ' Check and assign the slash
    
    If Right(strLocation, 1) = "\" Then
        Slash = strLocation
    Else
        Slash = strLocation & "\"
    End If

End Function

Public Sub loadFade()

    Dim strFileName As String
    Dim intFreeFile As Integer
    Dim intLevel As Integer
    Dim strTemp As String
    
    ' Get the file name and check if it exists
    ' And if it doesn't we'll create it
    
    If Right(App.Path, 1) = "\" Then
        strFileName = App.Path & "fade.dat"
    Else
        strFileName = App.Path & "\" & "fade.dat"
    End If
    
    If FileExists(strFileName) = False Then
        intFreeFile = FreeFile
        Open strFileName For Output As #intFreeFile
        Write #intFreeFile, 60
        Close #intFreeFile
        intLevel = 60
    Else
        intFreeFile = FreeFile
        Open strFileName For Input As #intFreeFile
        Line Input #intFreeFile, strTemp
        intLevel = strTemp
        Close #intFreeFile
    End If
    
    ' Set the onscreen
    
    FadeLevel = intLevel

End Sub
