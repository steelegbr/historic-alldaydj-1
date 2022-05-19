VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSearchResultsAddAdvert 
   Caption         =   "Select An Advert"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lstAdverts 
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ID"
         Object.Width           =   353
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Artist"
         Object.Width           =   3752
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Track"
         Object.Width           =   3752
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdSelectAdvert 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2640
      Width           =   855
   End
   Begin VB.Timer tmrGetAd 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   2640
   End
End
Attribute VB_Name = "frmSearchResultsAddAdvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

' Clear and hide

lstAdverts.ListItems.Clear
Me.Visible = False
frmAddAdBreak.Visible = True

End Sub

Private Sub cmdSelectAdvert_Click()

Dim strSelectedAdvert As String
Dim strMessage As String

' Get the select

strSelectedAdvert = lstAdverts.SelectedItem.text

' Check that it is not blank

If strSelectedAdvert = "" Then
strMessage = "No advert selected."
Call Error_Handler(strMessage, 1)
Exit Sub
End If

' Add item

With frmAddAdBreak.lstItemData.ListItems.Add(, , lstAdverts.SelectedItem.text)
.SubItems(1) = lstAdverts.SelectedItem.SubItems(1)
.SubItems(2) = lstAdverts.SelectedItem.SubItems(2)
End With

' Close this screen

lstAdverts.ListItems.Clear
Me.Visible = False
frmAddAdBreak.Visible = True

End Sub



Private Sub tmrGetAd_Timer()

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strConnectionString As String
Dim strSQL As String
Dim strData(3) As String
Dim intLooper As Integer
Dim strTmpData As String

' Connect

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb"
strSQL = "SELECT itemid, artist, track FROM adverts"
dbConnection.Open strConnectionString
dbRecordset.Open strSQL, strConnectionString

' Clear old data

lstAdverts.ListItems.Clear

' Load the data onto the screen

Do While Not dbRecordset.EOF
intLooper = 0
For Each X In dbRecordset.Fields
intLooper = intLooper + 1
strData(intLooper) = X
Next X
With lstAdverts.ListItems.Add(, , strData(1))
.SubItems(1) = strData(2)
.SubItems(2) = strData(3)
End With
dbRecordset.MoveNext
Loop

' Close connections

dbRecordset.Close
dbConnection.Close

' Stop repetition

tmrGetAd.Enabled = False

End Sub
