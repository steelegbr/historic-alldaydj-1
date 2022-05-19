VERSION 5.00
Begin VB.Form frmAudioWall 
   Caption         =   "Audio Wall"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPlayer 
      Height          =   285
      Left            =   840
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   8160
      TabIndex        =   23
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   495
      Left            =   4800
      TabIndex        =   21
      Top             =   4560
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   495
      Left            =   4080
      TabIndex        =   20
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   19
      Left            =   7440
      TabIndex        =   19
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   18
      Left            =   5640
      TabIndex        =   18
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   17
      Left            =   3840
      TabIndex        =   17
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   16
      Left            =   2040
      TabIndex        =   16
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   15
      Left            =   240
      TabIndex        =   15
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   14
      Left            =   7440
      TabIndex        =   14
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   13
      Left            =   5640
      TabIndex        =   13
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   12
      Left            =   3840
      TabIndex        =   12
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   11
      Left            =   2040
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   10
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   9
      Left            =   7440
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   8
      Left            =   5640
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   7
      Left            =   3840
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   6
      Left            =   2040
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   4
      Left            =   7440
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   3
      Left            =   5640
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   2
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdItem 
      Height          =   975
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblPage 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   22
      Top             =   4560
      Width           =   375
   End
End
Attribute VB_Name = "frmAudioWall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

' Hide the search results

frmSearch.Visible = True
Unload Me

End Sub

Private Sub cmdItem_Click(Index As Integer)

Dim intItemID As Integer
Dim intPage As Integer
Dim strType As String

' Get the variables

intPage = lblPage.Caption
intItemID = frmSearchResults.lstResults.ListItems.item((intPage - 1) * 20 + Index + 1).text
strType = frmSearch.cmbType.text

' Add to the playlist

Call doAddToPlaylist(intItemID, strType)

End Sub

Private Sub cmdItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim vntData As Variant
Dim strFileToLoad As String
Dim strConnectionString As String
Dim strSQL As String
Dim lngChannel As Long

' Ignore left mouse button

If Button = 1 Then
Exit Sub
End If

' Retrieve the data

intPage = lblPage.Caption
intItemID = frmSearchResults.lstResults.ListItems.item((intPage - 1) * 20 + Index + 1).text
strType = frmSearch.cmbType.text

' Connect to database

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb"
dbConnection.Open strConnectionString
strSQL = "SELECT [file] FROM " & strType & " WHERE [itemid] = " & intItemID
dbRecordset.Open strSQL, dbConnection

' Obtain data from database

Do While Not dbRecordset.EOF
For Each vntData In dbRecordset.Fields
strFileToLoad = vntData
Next vntData
dbRecordset.MoveNext
Loop

' Close the database connection

dbRecordset.Close
dbConnection.Close

' Remove the #'s

If Left$(strFileToLoad, 1) = "#" Then
    strFileToLoad = Right$(strFileToLoad, Len(strFileToLoad) - 1)
End If

If Right$(strFileToLoad, 1) = "#" Then
    strFileToLoad = Left$(strFileToLoad, Len(strFileToLoad) - 1)
End If

Debug.Print strFileToLoad

' Load the sound file into the PFL sound card

Call BASS_SetDevice(GetSoundCard(PFL_SOUND_CARD))
If UCase$(Right$(strFileToLoad, 3)) = "WMA" Then
lngChannel = BASS_WMA_StreamCreateFile(BASSFALSE, strFileToLoad, 0, 0, BASS_STREAM_PRESCAN)
txtplayer.text = lngChannel
Else
lngChannel = BASS_StreamCreateFile(BASSFALSE, strFileToLoad, 0, 0, BASS_STREAM_PRESCAN)
txtplayer.text = lngChannel
End If

' Start the AGC and load global settings

Call BASS_FX_DSP_Set(lngChannel, BASS_FX_DSPFX_DAMP, 0)
Call BASS_FX_DSP_SetParameters(lngChannel, BASS_FX_DSPFX_DAMP, GlobalAGC)
Call BASS_FX_DSP_Reset(lngChannel, BASS_FX_DSPFX_DAMP)

' Play at full volume

Call BASS_ChannelPlay(lngChannel, 0)
Call BASS_ChannelSetAttributes(lngChannel, -1, 100, -101)

End Sub

Private Sub cmdItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim lngChannel As Long

' Ignore left mouse button

If Button = 1 Then
Exit Sub
End If

' Stop the player

lngChannel = txtplayer.text
Call BASS_SetDevice(GetSoundCard(PFL_SOUND_CARD))
Call BASS_ChannelStop(lngChannel)

End Sub

Private Sub cmdNext_Click()

Dim intPage As Integer
Dim intItems As Integer
Dim intPages As Integer
Dim intLooper As Integer
Dim strText As String
Dim strTemp As String

' Get the number of items

intItems = frmSearchResults.lstResults.ListItems.Count

' Get the number of pages available

intPages = Int(intItems / 20) + 1
intPage = lblPage.Caption

' Get the items to be loaded

intItems = intItems - (intPage * 20)

' Make all boxes visible

For intLooper = 0 To 19
cmdItem(intLooper).Visible = True
Next intLooper

' If less than 20 then hide other search boxes

If intItems < 20 Then
For intLooper = intItems To 19
cmdItem(intLooper).Visible = False
Next intLooper
End If

' Load the items

If intItems <= 20 Then
For intLooper = 1 To intItems
strText = frmSearchResults.lstResults.ListItems.item((intPage * 20) + intLooper).SubItems(1) & Chr(10) & frmSearchResults.lstResults.ListItems.item((intPage * 20) + intLooper).SubItems(2)
For intLooper2 = 1 To Len(strText)
strTemp = Left$(strText, intLooper2)
If Right$(strTemp, 1) = "&" Then
strText = strTemp & "&" & Right$(strText, Len(strText) - intLooper2)
intLooper2 = intLooper2 + 1
End If
Next intLooper2
cmdItem(intLooper - 1).Caption = strText
Next intLooper
cmdNext.Visible = False

Else
For intLooper = 1 To 20
strText = frmSearchResults.lstResults.ListItems.item((intPage * 20) + intLooper).SubItems(1) & Chr(10) & frmSearchResults.lstResults.ListItems.item((intPage * 20) + intLooper).SubItems(2)
For intLooper2 = 1 To Len(strText)
strTemp = Left$(strText, intLooper2)
If Right$(strTemp, 1) = "&" Then
strText = strTemp & "&" & Right$(strText, Len(strText) - intLooper2)
intLooper2 = intLooper2 + 1
End If
Next intLooper2
cmdItem(intLooper - 1).Caption = strText
Next intLooper
End If

' Hide the next buttton if appropriate

If intItems < 20 Then
cmdNext.Visible = False
End If

' Update the page number

intPage = intPage + 1
lblPage.Caption = intPage
cmdPrevious.Visible = True

End Sub

Private Sub cmdPrevious_Click()

Dim intPage As Integer
Dim intLooper As Integer

' Get the number of items

intItems = frmSearchResults.lstResults.ListItems.Count

' Get the current page number

intPage = lblPage.Caption
intPage = intPage - 1

' Load the items and make all boxes visible

For intLooper = 0 To 19
cmdItem(intLooper).Visible = True
strText = frmSearchResults.lstResults.ListItems.item(((intPage - 1) * 20) + intLooper + 1).SubItems(1) & Chr(10) & frmSearchResults.lstResults.ListItems.item(((intPage - 1) * 20) + intLooper + 1).SubItems(2)
For intLooper2 = 1 To Len(strText)
strTemp = Left$(strText, intLooper2)
If Right$(strTemp, 1) = "&" Then
strText = strTemp & "&" & Right$(strText, Len(strText) - intLooper2)
intLooper2 = intLooper2 + 1
End If
Next intLooper2
cmdItem(intLooper).Caption = strText
Next intLooper


' Update the page number

lblPage.Caption = intPage

If intPage = 1 Then
cmdPrevious.Visible = False
End If

cmdNext.Visible = True

End Sub


