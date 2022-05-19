VERSION 5.00
Begin VB.Form frmRemoveFromDBSearch 
   Caption         =   "Search"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmRemoveFromDBSearch.frx":0000
      Left            =   1560
      List            =   "frmRemoveFromDBSearch.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox cmbSection 
      DataSource      =   "adoRemoveFromDBSearch"
      Height          =   315
      ItemData        =   "frmRemoveFromDBSearch.frx":0058
      Left            =   4680
      List            =   "frmRemoveFromDBSearch.frx":0062
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO!"
      Default         =   -1  'True
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblRB 
      Caption         =   ")"
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
      Left            =   5760
      TabIndex        =   8
      Top             =   240
      Width           =   135
   End
   Begin VB.Label lblLB 
      Caption         =   "("
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
      Left            =   4560
      TabIndex        =   7
      Top             =   240
      Width           =   135
   End
   Begin VB.Label lblIn 
      Caption         =   "in"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lblSearchFor 
      Caption         =   "Search for"
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
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmRemoveFromDBSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

' Clear the data and hide the form

Me.Visible = False
txtSearch.text = ""

End Sub

Private Sub cmdGo_Click()

' Initiate variables

Dim strSQL As String
Dim strMessage As String
Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strData(3) As String
Dim intLooper As Integer
Dim strItem As String
Dim x As ADODB.Field

' Clear previous results

frmSearchResultsRemoveFromDB.lstSearchResults.ListItems.Clear

' Connect to the database

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb"
dbConnection.Open (strConnectionString)

' Only allow database connection if there is data in the two drop down boxes

If cmbSection.text <> "" And cmbType.text <> "" Then
strSQL = "SELECT [itemid], [artist], [track] FROM [" & LCase$(cmbType.text) & "] WHERE [" & LCase$(cmbSection.text) & "] LIKE ""%" & txtSearch.text & "%"""
dbRecordset.Open strSQL, dbConnection

' Pull out of database

Do Until dbRecordset.EOF
intLooper = 0
For Each x In dbRecordset.Fields
intLooper = intLooper + 1
strData(intLooper) = x
Next x
dbRecordset.MoveNext

' Produce results

With frmSearchResultsRemoveFromDB.lstSearchResults.ListItems.Add(, , strData(1))
.SubItems(1) = strData(2)
.SubItems(2) = strData(3)
End With

Loop

' Close the database connection

dbRecordset.Close

' Check for results

If frmSearchResultsRemoveFromDB.lstSearchResults.ListItems.Count > 0 Then
strMessage = frmSearchResultsRemoveFromDB.lstSearchResults.ListItems.Count & " results found."
MsgBox strMessage, vbinfromation, "Search Results"
Me.Visible = False
frmSearchResultsRemoveFromDB.Visible = True
End If

If frmSearchResultsRemoveFromDB.lstSearchResults.ListItems.Count = 0 Then
strMessage = "No results found." & Chr(10) & Chr(10) & "Try another, broader search."
MsgBox strMessage, vbinfromation, "Search Results"
End If

Else
MsgBox "Please choose an option from each drop-down box.", vbInformation, "Lack of Data"
End If

End Sub

Private Sub Form_Load()

    ' Defaults

    cmbType.text = "Songs"
    cmbSection.text = "Artist"

End Sub



