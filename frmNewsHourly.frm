VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmNewsHourly 
   Caption         =   "News / Hourly Settings"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSecs 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtMins 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
   Begin VB.CheckBox chkNews 
      Alignment       =   1  'Right Justify
      Caption         =   "Run News?"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdlCommon 
      Left            =   3000
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Browse"
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdThenPlayBrowse 
      Caption         =   "Browse"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtThenPlayLocation 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtLength 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblColon 
      Alignment       =   2  'Center
      Caption         =   ":"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblNewsLength 
      Alignment       =   1  'Right Justify
      Caption         =   "Feed Length:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblNewsFile 
      Alignment       =   1  'Right Justify
      Caption         =   "Then Play:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblLength 
      Alignment       =   1  'Right Justify
      Caption         =   "Length:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblNewsJingle 
      Alignment       =   1  'Right Justify
      Caption         =   "News Jingle:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmNewsHourly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdBrowse_Click()

' Initiate Variables

Dim strFileToLoad As String

' Use common controls for opening file

cdlCommon.InitDir = App.Path
cdlCommon.Filter = "MP3 Files (*.mp3)|*.mp3|" & _
                   "WMA Files (*.wma)|*.wma|" & _
                   "WAV Files (*.wav)|*.wav|" & _
                   "OGG Files (*.ogg)|*.ogg|" & _
                   "AIFF Files (*.aiff)|*.aiff|"
                   
cdlCommon.ShowOpen
strFileToLoad = cdlCommon.FileName

' Load the location into the textbox

txtLocation.text = strFileToLoad

End Sub

Private Sub cmdCancel_Click()

' Clear the screen

txtLocation.text = ""
txtLength.text = ""
txtThenPlayLocation.text = ""

' Hide the screen

Unload Me

End Sub

Private Sub cmdOK_Click()

' Initiate Variables

Dim strJingleLocation As String
Dim intLength As Long
Dim strPlayAfterLocation As String
Dim strFileLocation As String
Dim strMessageText As String
Dim intMins As Integer
Dim intSecs As Integer

' Get values

On Error GoTo WrongDataType

strJingleLocation = txtLocation.text
intLength = txtLength.text
strPlayAfterLocation = txtThenPlayLocation.text
intMins = txtMins.text
intSecs = txtSecs.text

' Check they are valid

If intLength < 0 Then
strMessageText = "Whoa! Time machines have not been invented yet!" & Chr(10) & Chr(10) & "Please enter the time as a posative value."
Call Error_Handler(strMessageText, 1)
Exit Sub
End If

If intLength > 3599 Then
strMessageText = "The time you entered for the jingle length was too long." & Chr(10) & Chr(10) & "Check the value entered and try again."
Call Error_Handler(strMessageText, 1)
Exit Sub
End If

If intMins < 0 Then
    strMessageText = "You cannot have a negative time length." & vbCrLf & vbCrLf & "Please enter a positive whole number of minutes in the minutes box."
    Call Error_Handler(strMessageText, 1)
    Exit Sub
End If

If intSecs < 0 Then
    strMessageText = "You cannot have a negative time length." & vbCrLf & vbCrLf & "Please enter a positive whole number of minutes in the seconds box."
    Call Error_Handler(strMessageText, 1)
    Exit Sub
End If

If intMins > 59 Then
    strMessageText = "The number of minutes you entered was too large." & vbCrLf & vbCrLf & "Please enter a whole number between (and including) 0 and 59."
    Call Error_Handler(strMessageText, 1)
    Exit Sub
End If

If intSecs > 59 Then
    strMessageText = "The number of seconds you entered was too large." & vbCrLf & vbCrLf & "Please enter a whole number between (and including) 0 and 59."
    Call Error_Handler(strMessageText, 1)
    Exit Sub
End If

If intMins = 0 And intSecs = 0 Then
    strMessageText = "You have entered a zero time length." & vbCrLf & vbCrLf & "Please enter a time between (and including) 00:00 and 59:59."
    Call Error_Handler(strMessageText, 1)
    Exit Sub
End If

' Delete the old file

strFileLocation = App.Path & "\news.dat"
Kill strFileLocation

' Create the new

Open strFileLocation For Output As #1
Print #1, strJingleLocation
Print #1, intLength
Print #1, strPlayAfterLocation
Print #1, chkNews.value
Print #1, intMins
Print #1, intSecs
Close #1

Me.Visible = False

Exit Sub

WrongDataType:
strMessageText = "Please enter a number into the length / minutes / seconds box(es)."
Call Error_Handler(strMessageText, 1)
Exit Sub

End Sub

Private Sub cmdThenPlayBrowse_Click()

' Initiate Variables

Dim strFileToLoad As String

' Use common controls for opening file

cdlCommon.InitDir = App.Path
cdlCommon.Filter = "MP3 Files (*.mp3)|*.mp3|" & _
                   "WMA Files (*.wma)|*.wma|" & _
                   "WAV Files (*.wav)|*.wav|"
                   
cdlCommon.ShowOpen
strFileToLoad = cdlCommon.FileName

' Load the location into the textbox

txtThenPlayLocation.text = strFileToLoad

End Sub

Private Sub Form_Load()

' Initiate Variables

Dim strJingleLocation As String
Dim strLength As String
Dim strPlayAfterLocation As String
Dim strFileLocation As String
Dim strNews As String
Dim strDescription As String
Dim lngInput As Long
Dim strMins As String
Dim strSecs As String

' Read the data from the file

strFileLocation = App.Path + "\news.dat"
Open strFileLocation For Input As #1
Line Input #1, strJingleLocation
Line Input #1, strLength
Line Input #1, strPlayAfterLocation
Line Input #1, strNews
Line Input #1, strMins
Line Input #1, strSecs
Close #1

' Write it to the screen

txtLocation.text = strJingleLocation
txtLength.text = strLength
txtThenPlayLocation.text = strPlayAfterLocation
chkNews.value = strNews
txtMins.text = strMins
txtSecs.text = strSecs

End Sub
