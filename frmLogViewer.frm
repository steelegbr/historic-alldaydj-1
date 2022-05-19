VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmLogViewer 
   Caption         =   "Log Viewer"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmLogViewer.frx":0000
      Left            =   240
      List            =   "frmLogViewer.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin ComctlLib.ListView lstLogView 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Track ID"
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
         Text            =   "Record Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Composer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Length"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Played"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Time Played"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "frmLogViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbType_Click()

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strConnectionString As String
Dim strSQL As String
Dim vntData As Variant
Dim strData(8) As String
Dim intLooper As Integer

' Clear the previous results

lstLogView.ListItems.Clear

' Connect

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\logs.mdb"
dbConnection.Open strConnectionString

' Open the recordset

strSQL = "SELECT [trackid], [artist], [track], [record_company], [composer], [length_time], [played_time], [when_played] FROM [" & LCase$(cmbType.text) & "] ORDER BY [when_played]"
dbRecordset.Open strSQL, dbConnection

' Retrieve the data

Do While Not dbRecordset.EOF
intLooper = 0
For Each vntData In dbRecordset.Fields
intLooper = intLooper + 1
strData(intLooper) = vntData
Next vntData
With lstLogView.ListItems.Add(, , strData(1))
.SubItems(1) = strData(2)
.SubItems(2) = strData(3)
.SubItems(3) = strData(4)
.SubItems(4) = strData(5)
.SubItems(5) = strData(6)
.SubItems(6) = strData(7)
.SubItems(7) = strData(8)
End With
dbRecordset.MoveNext
Loop

End Sub

Private Sub cmdClear_Click()

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strSQL As String
Dim strConnectionString As String
Dim strMessage As String
Dim ynClear As Variant

' Check there is something to clear

If cmbType.text = "" Then
strMessage = "Cannot clear nothing." & Chr(10) & Chr(10) & "Please select a log to clear."
Call Error_Handler(strMessage, 1)
Exit Sub
End If

' Confirm the user wishes to clear the log

strMessage = "Confirm you wish to clear the " & cmbType.text & " log?"
ynClear = MsgBox(strMessage, vbYesNo + vbQuestion, "Confirm")

' Exit the sub if no is selected

If ynClear = vbNo Then
Exit Sub
End If

' Connect, delete and disconnect

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\logs.mdb"
strSQL = "DELETE FROM [" & cmbType.text & "]"
dbConnection.Open strConnectionString
dbConnection.Execute strSQL
dbConnection.Close

' Refresh the screen

Call cmbType_Click

End Sub

Private Sub cmdClose_Click()

' Hide the window and clear all values

lstLogView.ListItems.Clear
Me.Visible = False

End Sub

Private Sub cmdPrint_Click()


Const Margin = 60
Const COL_MARGIN = 240

Dim ymin As Single
Dim ymax As Single
Dim xmin As Single
Dim xmax As Single
Dim num_cols As Integer
Dim column_header As ColumnHeader
Dim list_item As ListItem
Dim i As Integer
Dim num_subitems As Integer
Dim col_wid() As Single
Dim x As Single
Dim y As Single
Dim line_hgt As Single
Dim ynPrint As Variant
' Check this the file is to be printed

ynPrint = MsgBox("Are you sure you wish to print this report?", vbYesNo + vbQuestion, "Print")
If ynPrint = vbNo Then
Exit Sub
End If

' Print the data
' Code below edited from an example at http://www.vb-helper.com/howto_listview_print.html

    Printer.CurrentX = 1440
    Printer.CurrentY = 1440

    xmin = Printer.CurrentX
    ymin = Printer.CurrentY
    Printer.Orientation = 2

    ' ******************
    ' Get column widths.
    num_cols = lstLogView.ColumnHeaders.Count
    ReDim col_wid(1 To num_cols)

    ' Check the column headers.
    For i = 1 To num_cols
        col_wid(i) = _
            Printer.TextWidth(lstLogView.ColumnHeaders(i).text)
    Next i

    ' Check the items.
    num_subitems = num_cols - 1
    For Each list_item In lstLogView.ListItems
        ' Check the item.
        If col_wid(1) < Printer.TextWidth(list_item.text) _
            Then _
           col_wid(1) = Printer.TextWidth(list_item.text)

        ' Check the subitems.
        For i = 1 To num_subitems
            If col_wid(i + 1) < _
                Printer.TextWidth(list_item.SubItems(i)) _
                Then _
               col_wid(i + 1) = _
                   Printer.TextWidth(list_item.SubItems(i))
        Next i
    Next list_item

    ' Add a column margin.
    For i = 1 To num_cols
        col_wid(i) = col_wid(i) + COL_MARGIN
    Next i

    ' *************************
    ' Print the column headers.
    Printer.CurrentY = ymin + Margin
    Printer.CurrentX = xmin + Margin
    x = xmin + Margin
    For i = 1 To num_cols
        Printer.CurrentX = x
        Printer.Print FittedText( _
            lstLogView.ColumnHeaders(i).text, col_wid(i));
        x = x + col_wid(i)
    Next i
    xmax = x + Margin

    Printer.Print
    line_hgt = Printer.TextHeight("X")
    y = Printer.CurrentY + line_hgt / 2
    Printer.Line (xmin, y)-(xmax, y)
    y = y + line_hgt / 2

    ' Print the rows.
    num_subitems = num_cols - 1
    For Each list_item In lstLogView.ListItems
        x = xmin + Margin

        ' Print the item.
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print FittedText( _
            list_item.text, col_wid(1));
        x = x + col_wid(1)

        ' Print the subitems.
        For i = 1 To num_subitems
            Printer.CurrentX = x
            Printer.Print FittedText( _
                list_item.SubItems(i), col_wid(i + 1));
            x = x + col_wid(i + 1)
        Next i

        y = y + line_hgt * 1.5
    Next list_item
    ymax = y

    ' Draw lines around it all.
    Printer.Line (xmin, ymin)-(xmax, ymax), , B

    x = xmin + Margin / 2
    For i = 1 To num_cols - 1
        x = x + col_wid(i)
        Printer.Line (x, ymin)-(x, ymax)
    Next i
    
    Printer.EndDoc
    
End Sub

' Return as much text as will fit in this width.
Private Function FittedText(ByVal txt As String, ByVal wid _
    As Single) As String
    Do While Printer.TextWidth(txt) > wid
        txt = Left$(txt, Len(txt) - 1)
    Loop
    FittedText = txt
End Function
