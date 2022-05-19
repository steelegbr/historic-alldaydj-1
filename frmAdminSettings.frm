VERSION 5.00
Begin VB.Form frmAdminSettings 
   Caption         =   "Admin Settings"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLoadData 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   2280
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtPassword2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtPassword1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtStationName 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Station Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lblUserName 
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmAdminSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()

' Clear values and hide

txtStationName.text = ""
txtUserName.text = ""
txtPassword1.text = ""
txtPassword2.text = ""
Me.Visible = False

End Sub

Private Sub cmdOK_Click()

' Initiate variables

Dim strStationName As String
Dim strUsername As String
Dim strPassword1 As String
Dim strPassword2 As String
Dim strMessage As String
Dim strFilePath As String
Dim strTemp As String

' Get values

strStationName = txtStationName.text
strUsername = txtUserName.text
strPassword1 = txtPassword1.text
strPassword2 = txtPassword2.text

If strPassword1 <> strPassword2 Then
strMessage = "Passwords do NOT match!" & Chr(10) & Chr(10) & "No data was changed"
txtPassword1.text = ""
txtPassword2.text = ""
txtPassword1.SetFocus
Call Error_Handler(strMessage, 1)
Exit Sub
End If

' If the password or username section is blank then use old settings

If strPassword1 = "" Or strUsername = "" Then
strFilePath = App.Path & "\station.dat"
Open strFilePath For Input As #1
Line Input #1, strTemp
Line Input #1, strUsername
Line Input #1, strPassword1
Close #1
End If

' Delete old file and write new one

strFilePath = App.Path & "\station.dat"
Kill strFilePath

Open strFilePath For Output As #1
Print #1, strStationName
Print #1, strUsername
Print #1, strPassword1
Close #1

' Change the station name displayed

main.Caption = "AllDay DJ - " & strStationName

' Clear values and hide

txtStationName.text = ""
txtUserName.text = ""
txtPassword1.text = ""
txtPassword2.text = ""
Me.Visible = False

End Sub

Private Sub tmrLoadData_Timer()

' Initiate variables

Dim strFilePath As String
Dim strStationName As String

' Set file path

strFilePath = App.Path & "\station.dat"

' Load data from file

Open strFilePath For Input As #1
Line Input #1, strStationName
Close #1

' Show it onscreen

txtStationName.text = strStationName

' Stop the timer

tmrLoadData.Enabled = False

End Sub
