VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmAddAdBreak 
   Caption         =   "Add An Advert Break"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lstItemData 
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3625
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
         Object.Width           =   4810
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Artist"
         Object.Width           =   4810
      EndProperty
   End
   Begin VB.TextBox txtMinutes 
      Height          =   285
      Left            =   5280
      TabIndex        =   5
      Text            =   "00"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtHour 
      Height          =   285
      Left            =   4680
      TabIndex        =   4
      Text            =   "00"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Text            =   "2005"
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      ItemData        =   "frmAddAdBreak.frx":0000
      Left            =   2520
      List            =   "frmAddAdBreak.frx":002B
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cmbDay 
      Height          =   315
      ItemData        =   "frmAddAdBreak.frx":0091
      Left            =   1680
      List            =   "frmAddAdBreak.frx":0108
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblColon 
      Alignment       =   1  'Right Justify
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblAt 
      Alignment       =   1  'Right Justify
      Caption         =   "at"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblDateTime 
      Alignment       =   1  'Right Justify
      Caption         =   "Date && Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddAdBreak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

frmSearchResultsAddAdvert.Visible = True
frmSearchResultsAddAdvert.tmrGetAd.Enabled = True
Me.Visible = False

End Sub

Private Sub cmdCancel_Click()

Me.Visible = False
lstItemData.ListItems.Clear
txtHour.text = "00"
txtMinutes.text = "00"

End Sub

Private Sub cmdClear_Click()

lstItemData.ListItems.Clear

End Sub


Private Sub cmdOK_Click()

Dim intListLength As Integer
Dim strTempData() As String
Dim strFromItem As String
Dim intNumberOfItems As Integer
Dim intLooper As Integer
Dim intListPosition As Integer
Dim strMessage As String
Dim intHours As Integer
Dim strHours As String
Dim strMinutes As String
Dim intMinutes As Integer
Dim intYear As Integer
Dim strYear As String
Dim intDay As Integer
Dim intMonth As Integer
Dim strMonth As String
Dim strToday As String
Dim dblDateDiff As Double
Dim strDateTime As String
Dim dbConnection As New ADODB.Connection
Dim strSQL As String
Dim strConnectionString As String

' Get the number of data items

intNumberOfItems = lstItemData.ListItems.Count

' Redim the arrays

ReDim strItemNumber(lstItemData.ListItems.Count) As String
ReDim strItemText(lstItemData.ListItems.Count) As String

' Check that there is actually an ad break to process

If intNumberOfItems = 0 Then
strMessage = "There are no advertisements in this break." & Chr(10) & Chr(10) & "Please select some advertisements or select cancel."
Call Error_Handler(strMessage, 1)
Exit Sub
End If

' Get the data using a loop

For intLooper = 1 To intNumberOfItems
strItemNumber(intLooper) = lstItemData.ListItems.Item(intLooper).text
strItemText(intLooper) = lstItemData.ListItems.Item(intLooper).SubItems(1) & " - " & lstItemData.ListItems.Item(intLooper).SubItems(2)
Next intLooper

' Get the time and date

' On Error GoTo noDataEntered

intHours = txtHour.text
intMinutes = txtMinutes.text
strMonth = cmbMonth.text
intYear = txtYear.text
intDay = cmbDay.text
strYear = intYear

' Convert minutes and seconds into a valid format

If intMinutes < 10 Then
strMinutes = "0" & intMinutes
End If
If intMinutes >= 10 Then
strMinutes = intMinutes
End If

If intHours < 10 Then
strHours = "0" & intHours
End If
If intHours >= 10 Then
strHours = intHours
End If

' Convert the month

Select Case strMonth
Case "January"
intMonth = 1
Case "February"
intMonth = 2
Case "March"
intMonth = 3
Case "April"
intMonth = 4
Case "May"
intMonth = 5
Case "June"
intMonth = 6
Case "July"
intMonth = 7
Case "August"
intMonth = 8
Case "September"
intMonth = 9
Case "October"
intMonth = 10
Case "November"
intMonth = 11
Case "December"
intMonth = 12
End Select

' Checks that a valid number of days has been entered for the month

If intMonth = 2 Then
If intDay > 29 Then
strMessage = "Too many days for this month selected." & Chr(10) & Chr(10) & "Please select an earlier day and try again."
Call Error_Handler(strMessage, 1)
Exit Sub
End If

' Check for a leap year
' This is done by checking if the year divides by 4
' Another check is made to make sure that if it is a century it will divide by 400

If intDay = 29 Then
If intYear Mod 4 > 0 Then
strMessage = "Too many days for this month selected." & Chr(10) & Chr(10) & "Please select an earlier day and try again."
Call Error_Handler(strMessage, 1)
Exit Sub
End If

strYear = Right$(strYear, 2)

If strYear = "00" Then
If intYear Mod 400 > 0 Then
strMessage = "Too many days for this month selected." & Chr(10) & Chr(10) & "Please select an earlier day and try again."
Call Error_Handler(strMessage, 1)
Exit Sub
End If
End If
End If
End If

If intMonth = 4 Or intMonth = 6 Or intMonth = 9 Or intMonth = 11 Then  ' April, June, September, November
If intDay > 30 Then
strMessage = "Too many days for this month selected." & Chr(10) & Chr(10) & "Please select an earlier day and try again."
Call Error_Handler(strMessage, 1)
Exit Sub
End If
End If

' Compare it to now

strDateTime = intDay & "/" & intMonth & "/" & intYear & " " & strHours & ":" & strMinutes & ":00"
strToday = Now()
dblDateDiff = DateDiff("s", strToday, strDateTime)

' If the date is earlier than now then

If dblDateDiff <= 0 Then
strMessage = "Invalid time entered." & Chr(10) & Chr(10) & "Please enter a later date / time."
Call Error_Handler(strMessage, 1)
Exit Sub
End If

' Otherwise go ahead and enter it into the database

On Error GoTo DuplicateAdBreak

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb"
dbConnection.Open strConnectionString
strSQL = "INSERT INTO AdSchedule VALUES (""" & strDateTime & """)"
dbConnection.Execute strSQL

' Create the new table and add all of the values

strSQL = "CREATE TABLE [" & strDateTime & "] (itemId integer, itemText string)"
dbConnection.Execute strSQL

For intLooper = 1 To lstItemData.ListItems.Count
strSQL = "INSERT INTO [" & strDateTime & "] VALUES (" & strItemNumber(intLooper) & ", """ & strItemText(intLooper) & """)"
dbConnection.Execute strSQL
Next intLooper

' Close the connection and screen

dbConnection.Close
Unload Me
frmAdverts.Visible = True
lstItemData.ListItems.Clear
txtHour.text = "00"
txtMinutes.text = "00"

Exit Sub

noDataEntered:

strMessage = "No Date Entered." & Chr(10) & Chr(10) & "Please enter a valid date."
Call Error_Handler(strMessage, 1)
Exit Sub

DuplicateAdBreak:

strMessage = "The time & date selected is the same as antother advert break." & Chr(10) & Chr(10) & "Please enter a different date / time."
Call Error_Handler(strMessage, 1)

End Sub

Private Sub cmdRemove_Click()

Dim strSelectedData As String
Dim intDataItemNumber As Integer

' Checks that a data has been selected

strSelectedData = lstItemData.SelectedItem.text
If strSelectedData = "" Then
Exit Sub
End If

' Delete the item

lstItemData.ListItems.Remove lstItemData.SelectedItem.Index

End Sub

Private Sub txtHour_Change()

Dim intCheck As Integer
Dim dblCheck As Double

' Check that it is a number

On Error GoTo NotANumber
intCheck = txtHour.text

' Checks that the number is in the limits

If intCheck > 23 Or intCheck < 0 Then
txtHour.text = "00"
intCheck = 0
End If

' Check that an integer has been entered

dblCheck = txtHour.text
If dblCheck <> intCheck Then
txtHour.text = intCheck
End If

Exit Sub

NotANumber:

txtHour.text = ""

End Sub

Private Sub txtMinutes_Change()

Dim intCheck As Integer
Dim dblCheck As Double

' Check that it is a number

On Error GoTo NotANumber
intCheck = txtMinutes.text

' Checks that the number is in the limits

If intCheck > 59 Or intCheck < 0 Then
txtMinutes.text = "00"
intCheck = 0
End If

' Check that an integer has been entered

dblCheck = txtMinutes.text
If dblCheck <> intCheck Then
txtMinutes.text = intCheck
End If

Exit Sub

NotANumber:

txtMinutes.text = ""

End Sub

Private Sub txtYear_Change()

' This subroutine checks that a number has been entered

Dim intCheck As Integer
Dim dblCheck As Double

On Error GoTo NotANumber
intCheck = txtYear.text
Exit Sub

' Check that an integer has been entered

dblCheck = txtHour.text
If dblCheck <> intCheck Then
txtYear.text = intCheck
End If

Exit Sub

NotANumber:

txtYear.text = ""

End Sub
