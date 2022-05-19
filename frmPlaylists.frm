VERSION 5.00
Begin VB.Form frmPlaylists 
   Caption         =   "Scheduled Playlists"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbPlaylistBlock 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox lstPlaylists 
      Height          =   1620
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   4815
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add New"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmPlaylists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbPlaylistBlock_Click()

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strConnectionString
Dim strSQL As String
Dim vntData As Variant
Dim strData As String
Dim intLooper As Integer

' Clear previous results

lstPlaylists.Clear

' Connect to the database

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb;Persist Security Info=False"
dbConnection.Open (strConnectionString)

' Open the recordset

strSQL = "SELECT * FROM [PL" & cmbPlaylistBlock.text & "]"
dbRecordset.Open strSQL, dbConnection

' Loop for the data

Do While Not dbRecordset.EOF
intLooper = 0
For Each vntData In dbRecordset.Fields
intLooper = intLooper + 1
strData = vntData
Next vntData
lstPlaylists.AddItem strData
dbRecordset.MoveNext
Loop

' Close the connection

dbRecordset.Close
dbConnection.Close


End Sub

Private Sub cmdAddNew_Click()

' Show the add playlist box

frmAddPlaylist.Visible = True
Unload Me

End Sub

Private Sub cmdClose_Click()

Me.Visible = False
lstPlaylists.Clear

End Sub

Private Sub cmdDelete_Click()

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strConnectionString As String
Dim strSQL As String
Dim strMessage As String
Dim vntData As Variant

' Check that a playlist has been selected

If cmbPlaylistBlock.text = "" Then
strMessage = "No playlist seleced." & Chr(10) & Chr(10) & "Please select one from the drop-down menu."
Call Error_Handler(strMessage, 1)
Exit Sub
End If

' If a playlist has been selected then go ahaed and connect to the database

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb;Persist Security Info=False"
dbConnection.Open (strConnectionString)

' Delete the table

strSQL = "DROP TABLE [PL" & cmbPlaylistBlock.text & "]"
dbConnection.Execute strSQL

' Delete the entry

strSQL = "DELETE FROM [PlaylistSchedule] WHERE [playlistblock] = '" & cmbPlaylistBlock.text & "'"
dbConnection.Execute strSQL

' Clear the screen

cmbPlaylistBlock.Clear
lstPlaylists.Clear

strSQL = "SELECT * FROM [PlaylistSchedule]"
dbRecordset.Open strSQL, dbConnection

' Get the data

Do While Not dbRecordset.EOF
For Each vntData In dbRecordset.Fields
cmbPlaylistBlock.AddItem vntData
Next vntData
dbRecordset.MoveNext
Loop

' Disconnect

dbConnection.Close

End Sub

Private Sub Form_Load()

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strConnectionString
Dim strSQL As String
Dim vntData As Variant

' Clear previous results

cmbPlaylistBlock.Clear

' Connect to the database

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb;Persist Security Info=False"
dbConnection.Open (strConnectionString)

' Send the SQL command

strSQL = "SELECT * FROM [PlaylistSchedule]"
dbRecordset.Open strSQL, dbConnection

' Get the data

Do While Not dbRecordset.EOF
For Each vntData In dbRecordset.Fields
cmbPlaylistBlock.AddItem vntData
Next vntData
dbRecordset.MoveNext
Loop

End Sub


