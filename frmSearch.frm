VERSION 5.00
Begin VB.Form frmSearch 
   Caption         =   "Search"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO!"
      Default         =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cmbSection 
      DataSource      =   "adoRemoveFromDBSearch"
      Height          =   315
      ItemData        =   "frmSearch.frx":0000
      Left            =   4560
      List            =   "frmSearch.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmSearch.frx":002B
      Left            =   1440
      List            =   "frmSearch.frx":0041
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1575
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
      TabIndex        =   8
      Top             =   120
      Width           =   1335
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
      TabIndex        =   7
      Top             =   600
      Width           =   255
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
      TabIndex        =   6
      Top             =   120
      Width           =   135
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
      TabIndex        =   5
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

' Clear and unload

txtSearch.text = ""
Unload Me

End Sub

Private Sub cmdGo_Click()

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strSearchKey As String
Dim strConnectionString As String
Dim strSQL As String
Dim strMessage As String
Dim intLooper As Integer
Dim vntData As Variant
Dim strData(3) As String

' Check that there is a valid search

If cmbType.text = "" Or cmbSection.text = "" Then
strMessage = "Not a valid search." & Chr(10) & Chr(10) & "Please select an option in each of the dropdown boxes."
Call Error_Handler(strMessage, 1)
Exit Sub
End If

' Clear previous results

frmSearchResults.lstResults.ListItems.Clear

' Connect to the database

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb"
dbConnection.Open strConnectionString

' Create the sql command and open the recordset

strSQL = "SELECT [itemid], [artist], [track] FROM [" & LCase$(cmbType.text) & "] WHERE " & LCase$(cmbSection.text) & " LIKE ""%" & txtSearch.text & "%"" ORDER BY [artist]"
dbRecordset.Open strSQL, dbConnection

' Cycle through the record set

Do While Not dbRecordset.EOF
intLooper = 0
For Each vntData In dbRecordset.Fields
intLooper = intLooper + 1
strData(intLooper) = vntData
Next vntData
With frmSearchResults.lstResults.ListItems.Add(, , strData(1))
.SubItems(1) = strData(2)
.SubItems(2) = strData(3)
End With
dbRecordset.MoveNext
Loop

' Show search results

frmAudioWall.Visible = True
Call doAudioWall
Me.Visible = False

End Sub


Private Sub Form_Load()

    ' Defaults

    cmbType.text = "Songs"
    cmbSection.text = "Artist"

End Sub
