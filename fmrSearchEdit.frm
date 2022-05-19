VERSION 5.00
Begin VB.Form frmSearchEdit 
   Caption         =   "Search (Edit Entries)"
   ClientHeight    =   1245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1245
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "fmrSearchEdit.frx":0000
      Left            =   1440
      List            =   "fmrSearchEdit.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox cmbSection 
      DataSource      =   "adoRemoveFromDBSearch"
      Height          =   315
      ItemData        =   "fmrSearchEdit.frx":0058
      Left            =   4560
      List            =   "fmrSearchEdit.frx":0065
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO!"
      Default         =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
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
      Left            =   5640
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
      Left            =   4440
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
      Left            =   960
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
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmSearchEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

txtSearch.text = ""
frmAdminOptions.Visible = True
Unload Me

End Sub

Private Sub cmdGo_Click()

' Initiate variables

Dim strSQL As String
Dim strMessage As String
Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strData(1 To 12) As String
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
strSQL = "SELECT * FROM [" & LCase$(cmbType.text) & "] WHERE [" & LCase$(cmbSection.text) & "] LIKE ""%" & txtSearch.text & "%"""
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

With frmSearchResultsEdit.lstSearchResults.ListItems.Add(, , strData(1))
For intLooper = 2 To 12
.SubItems(intLooper - 1) = strData(intLooper)
Next intLooper
End With

Loop

' Close the database connection

dbRecordset.Close

' Check for results

If frmSearchResultsEdit.lstSearchResults.ListItems.Count > 0 Then
strMessage = frmSearchResultsEdit.lstSearchResults.ListItems.Count & " results found."
MsgBox strMessage, vbinfromation, "Search Results"
Me.Visible = False
frmSearchResultsEdit.Visible = True
End If

If frmSearchResultsEdit.lstSearchResults.ListItems.Count = 0 Then
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

