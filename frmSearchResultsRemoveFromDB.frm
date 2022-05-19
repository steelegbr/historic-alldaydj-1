VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSearchResultsRemoveFromDB 
   Caption         =   "Search Results"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin ComctlLib.ListView lstSearchResults 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4048
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
         Object.Width           =   6821
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Track"
         Object.Width           =   6821
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "frmSearchResultsRemoveFromDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

Me.Visible = False
frmRemoveFromDBSearch.Visible = True

End Sub

Private Sub cmdDelete_Click()

Dim dbConnection As New ADODB.Connection
Dim strConnectionString As String
Dim strConfirmString As String
Dim strSQL As String
Dim strSearchResultSelected As String
Dim strTrackID As String

' Confirm the user wishes to delete

strSearchResultSelected = lstSearchResults.SelectedItem.SubItems(1) & " - " & lstSearchResults.SelectedItem.SubItems(2)
strConfirmString = "Confirm delete: " & strSearchResultSelected
msgconfirm = MsgBox(strConfirmString, vbYesNo, "Confirm")

' If yes then connect to database and delete

If msgconfirm = vbYes Then

' Connect

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb"
dbConnection.Open (strConnectionString)

' Create delete SQL command

strTrackID = lstSearchResults.SelectedItem.text
strSQL = "DELETE From " & LCase$(frmRemoveFromDBSearch.cmbType.text) & " WHERE itemid = " & strTrackID

' Execute the SQL command

dbConnection.Execute strSQL

' Let the user know the records were deleted

MsgBox "Track removed from database.", vbInformation, "Track Removed"

' Hide this screen and return the user to the search

Me.Visible = False
frmRemoveFromDBSearch.Visible = False

End If

End Sub
