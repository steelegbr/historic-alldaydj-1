VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login To Admin Options"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancelLogin 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelLogin_Click()

' Hide the login screen and clear the typed characters

Me.Visible = False
txtPassword.text = ""
txtUserName.text = ""

End Sub


Private Sub cmdLogin_Click()

' Create the variables to be used

Dim strCorrectUsername As String
Dim strUsernameEntered As String
Dim strCorrectPassword As String
Dim strPasswordEntered As String
Dim strStationID As String
Dim strFileLocation As String
Dim strMessage As String

' Open the station.dat file with password

strFileLocation = App.Path + "\station.dat"
Open strFileLocation For Input As #1
Line Input #1, strStationID
Line Input #1, strCorrectUsername
Line Input #1, strCorrectPassword
Close #1

' Get the user entered data

strUsernameEntered = txtUserName.text
strPasswordEntered = txtPassword.text

' Check if login is a success

If strUsernameEntered = strCorrectUsername And strPasswordEntered = strCorrectPassword Then
txtUserName.text = ""
txtPassword.text = ""
txtUserName.SetFocus
Me.Visible = False
frmAdminOptions.Visible = False
frmAdminOptions.Visible = True
main.menuLogout.Visible = True
AdminLoggedIn = True
Else
strMessage = "Invalid username / password!"
txtUserName.text = ""
txtPassword.text = ""
txtUserName.SetFocus
Call Error_Handler(strMessage, 0)
End If

End Sub
