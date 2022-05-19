VERSION 5.00
Begin VB.Form frmSoundCard 
   Caption         =   "Soundcard Selection"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstPFL 
      Height          =   840
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   3375
   End
   Begin VB.ListBox lstMain 
      Height          =   840
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblPFL 
      Alignment       =   1  'Right Justify
      Caption         =   "PFL:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblMain 
      Alignment       =   1  'Right Justify
      Caption         =   "Main Output:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmSoundCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

' Clear the screen and hide it

lstMain.Clear
lstPFL.Clear
frmAdminOptions.Visible = True
Unload Me

End Sub

Private Sub cmdOK_Click()

Dim strMain As String
Dim strPFL As String
Dim strFileLocation As String
Dim intMain As Integer
Dim intPFL As Integer

' Get the values

strMain = lstMain.text
strPFL = lstPFL.text

' Check that something was selected

If strMain = "" Or strPFL = "" Then
Call Error_Handler("Please select a soundcard from each box.", 1)
Exit Sub
End If

intMain = lstMain.ListIndex
intPFL = lstPFL.ListIndex

' Otherwise delete the file and enter the data

strFileLocation = App.Path & "\soundcard.dat"
Kill strFileLocation
Open strFileLocation For Output As #1
Print #1, intMain
Print #1, intPFL
Close #1

' Clear the screen and hide it

lstMain.Clear
lstPFL.Clear
frmAdminOptions.Visible = True
Unload Me

End Sub

Private Sub Form_Load()

Dim intCards As Integer
Dim intLooper As Integer

' Get the number of items

intCards = 0
Do While BASS_GetDeviceDescription(intCards)
intCards = intCards + 1
Loop

' Create the array

intCards = intCards - 1
ReDim strCardDescription(intCards) As String

' Enter the data into the array

For intLooper = 0 To intCards
strCardDescription(intLooper) = VBStrFromAnsiPtr(BASS_GetDeviceDescription(intLooper))
Next intLooper

' Display on the screen

For intLooper = 0 To intCards
lstMain.AddItem strCardDescription(intLooper)
lstPFL.AddItem strCardDescription(intLooper)
Next intLooper

End Sub
