VERSION 5.00
Begin VB.Form frmAdverts 
   Caption         =   "Adverts"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add New"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.ListBox lstAdBreak 
      Height          =   1620
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   4815
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.ComboBox cmbAdBlock 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmAdverts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAdBlock_Click()

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strConnectionString
Dim strSQL As String
Dim vntData As Variant
Dim strData(2) As String
Dim intLooper As Integer
Dim strCompleteData As String

' Clear previous results

lstAdBreak.Clear

' Connect to the database

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb;Persist Security Info=False"
dbConnection.Open (strConnectionString)

' Open the recordset

strSQL = "SELECT * FROM [" & cmbAdBlock.text & "]"
dbRecordset.Open strSQL, dbConnection

' Loop for the data

Do While Not dbRecordset.EOF
intLooper = 0
For Each vntData In dbRecordset.Fields
intLooper = intLooper + 1
strData(intLooper) = vntData
Next vntData
strCompleteData = strData(1) & "* " & strData(2)
lstAdBreak.AddItem strCompleteData
dbRecordset.MoveNext
Loop

' Close the connection

dbRecordset.Close
dbConnection.Close


End Sub

Private Sub cmdAddNew_Click()

' Show the add ad break box

frmAddAdBreak.Visible = True
Unload Me

End Sub

Private Sub cmdClose_Click()

Me.Visible = False
lstAdBreak.Clear

End Sub

Private Sub cmdDelete_Click()

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strConnectionString As String
Dim strSQL As String
Dim strMessage As String
Dim vntData As Variant

' Check that an ad break has been selected

If cmbAdBlock.text = "" Then
strMessage = "No advert break seleced." & Chr(10) & Chr(10) & "Please select one from the drop-down menu."
Call Error_Handler(strMessage, 1)
Exit Sub
End If

' If an ad break has been selected then go ahaed and connect to the database

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb;Persist Security Info=False"
dbConnection.Open (strConnectionString)

' Delete the table

strSQL = "DROP TABLE [" & cmbAdBlock.text & "]"
dbConnection.Execute strSQL

' Delete the entry

strSQL = "DELETE FROM [AdSchedule] WHERE [adblock] = '" & cmbAdBlock.text & "'"
dbConnection.Execute strSQL

' Clear the screen

cmbAdBlock.Clear
lstAdBreak.Clear

strSQL = "SELECT * FROM [AdSchedule]"
dbRecordset.Open strSQL, dbConnection

' Get the data

Do While Not dbRecordset.EOF
For Each vntData In dbRecordset.Fields
cmbAdBlock.AddItem vntData
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

cmbAdBlock.Clear

' Connect to the database

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb;Persist Security Info=False"
dbConnection.Open (strConnectionString)

' Send the SQL command

strSQL = "SELECT * FROM [AdSchedule]"
dbRecordset.Open strSQL, dbConnection

' Get the data

Do While Not dbRecordset.EOF
For Each vntData In dbRecordset.Fields
cmbAdBlock.AddItem vntData
Next vntData
dbRecordset.MoveNext
Loop

End Sub
