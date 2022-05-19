VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNetworking 
   Caption         =   "Networking Options"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin MSComDlg.CommonDialog cdlCommon 
      Left            =   360
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.OptionButton radioLocation 
      Caption         =   "Use Remote Music Database && Settings"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin VB.OptionButton radioLocation 
      Caption         =   "Use Local Music Database && Settings"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmNetworking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()

Dim strFileTitle As String
Dim strFileName As String

' Use common controls for opening file

On Error GoTo Handler

cdlCommon.InitDir = App.Path
cdlCommon.Filter = "Record Collection |record_collection.mdb|"
cdlCommon.CancelError = True
cdlCommon.ShowOpen

strFileName = cdlCommon.FileName
strFileTitle = cdlCommon.FileTitle
strFileName = Left$(strFileName, (Len(strFileName) - Len("\" & strFileTitle)))
txtPath.text = strFileName

' Error handling

Handler:

End Sub

Private Sub cmdCancel_Click()

' Unload the form after clearing all boxes

txtPath.text = ""
Unload Me

End Sub

Private Sub cmdOK_Click()

Dim strSelectedOption As Integer
Dim strPath As String
Dim strLocation As String

' Retrieve the data from the radio buttons

strSelectedOption = radioLocation(0).value

' If they selected their own path then process
' Otherwise use App.Path

If strSelectedOption = 0 Then
strPath = txtPath.text
If strPath = "" Then
Call Error_Handler("Invalid Path. " & Chr(10) & Chr(10) & "Please select a valid path to database.", 1)
Exit Sub
End If
Else
strPath = "App.Path"
End If

' Save the data to the file

strLocation = App.Path & "\networking.dat"
Open strLocation For Output As #1
Print #1, strPath
Close #1

' Unload the form

Unload Me

End Sub

Private Sub Form_Load()

Dim strLocation As String
Dim strPath As String

' Load items from text file

strLocation = App.Path & "\networking.dat"
Open strLocation For Input As #1
Line Input #1, strPath
Close #1

' Process the data

If strPath = "App.Path" Then
    strPath = App.Path
    radioLocation(0).value = True
Else
    radioLocation(1).value = True
End If

txtPath.text = strPath

End Sub
