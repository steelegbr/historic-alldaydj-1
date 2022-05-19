VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSearchResultsEdit 
   Caption         =   "Search Results (Edit Entry)"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin ComctlLib.ListView lstSearchResults 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Artist"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Track"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Length"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Start"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Finish"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Fades"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Composer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   9
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Record Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   10
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Extra Info"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   11
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Intro Countdown"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmSearchResultsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

frmSearchEdit.Visible = True
Unload Me

End Sub

Private Sub cmdOK_Click()

Dim vntTemp As Variant
Dim intSelected As Integer
Dim boolFades As Boolean
Dim lngChannel As Long
Dim strFile As String
Dim vntLength As Variant

' Hide this screen

Me.Visible = False

' Load the values

intSelected = lstSearchResults.SelectedItem.Index
If intSelected <= 0 Then
Exit Sub
End If

frmEdit.txtFile.text = lstSearchResults.ListItems.item(intSelected).SubItems(6)
frmEdit.txtArtist.text = lstSearchResults.ListItems.item(intSelected).SubItems(1)
frmEdit.txtTrack.text = lstSearchResults.ListItems.item(intSelected).SubItems(2)
frmEdit.txtCompany.text = lstSearchResults.ListItems.item(intSelected).SubItems(9)
frmEdit.txtComposer.text = lstSearchResults.ListItems.item(intSelected).SubItems(8)
frmEdit.lblStart.Caption = lstSearchResults.ListItems.item(intSelected).SubItems(4)
frmEdit.lblEnd.Caption = lstSearchResults.ListItems.item(intSelected).SubItems(5)
frmEdit.lblIntro.Caption = lstSearchResults.ListItems.item(intSelected).SubItems(11)
boolFades = lstSearchResults.ListItems.item(intSelected).SubItems(7)
If boolFades = True Then
frmEdit.chkFades.value = 1
Else
frmEdit.chkFades.value = 0
End If
frmEdit.txtID.text = lstSearchResults.SelectedItem.text

' Load the file

strFile = frmEdit.txtFile.text
If Left$(strFile, 1) = "#" Then
    strFile = Right$(strFile, Len(strFile) - 1)
End If
If Right$(strFile, 1) = "#" Then
    strFile = Left$(strFile, Len(strFile) - 1)
End If
    
Call BASS_SetDevice(GetSoundCard(PFL_SOUND_CARD))
If UCase$(Right$(strFile, 3)) = "WMA" Then
lngChannel = BASS_WMA_StreamCreateFile(BASSFALSE, strFile, 0, 0, BASS_STREAM_PRESCAN)
frmEdit.txtplayer.text = lngChannel
Else
lngChannel = BASS_StreamCreateFile(BASSFALSE, strFile, 0, 0, BASS_STREAM_PRESCAN)
frmEdit.txtplayer.text = lngChannel
End If
vntLength = BASS_ChannelBytes2Seconds(lngChannel, BASS_ChannelGetLength(lngChannel))
frmEdit.sldTimer.max = vntLength
frmEdit.txtplayer.text = lngChannel

' Start the AGC and load global settings

Call BASS_FX_DSP_Set(lngChannel, BASS_FX_DSPFX_DAMP, 100)
Call BASS_FX_DSP_SetParameters(lngChannel, BASS_FX_DSPFX_DAMP, GlobalAGC)

' Hide this form

frmEdit.Visible = True
Unload Me

End Sub
