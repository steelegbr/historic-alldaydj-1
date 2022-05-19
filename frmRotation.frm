VERSION 5.00
Begin VB.Form frmRotation 
   Caption         =   "Rotation"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtRotLength 
      Height          =   375
      Left            =   660
      TabIndex        =   6
      Text            =   "0"
      Top             =   1245
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmRotation.frx":0000
      Left            =   120
      List            =   "frmRotation.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblRotation 
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label lblRotationTitle 
      Caption         =   "Rotation:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmRotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

Dim intRotLength As Integer

If cmbType.text <> "" Then
Select Case cmbType.text
Case "Songs"
lblRotation.Caption = lblRotation.Caption + "s;"
intRotLength = txtRotLength.text
intRotLength = intRotLength + 1
If intRotLength < 0 Then
intRotLength = 0
End If
txtRotLength.text = intRotLength
Case "Jingles"
lblRotation.Caption = lblRotation.Caption + "j;"
intRotLength = txtRotLength.text
intRotLength = intRotLength + 1
If intRotLength < 0 Then
intRotLength = 0
End If
txtRotLength.text = intRotLength
Case "Specialized"
lblRotation.Caption = lblRotation.Caption + "x;"
intRotLength = txtRotLength.text
intRotLength = intRotLength + 1
If intRotLength < 0 Then
intRotLength = 0
End If
txtRotLength.text = intRotLength
Case "Seasonal"
lblRotation.Caption = lblRotation.Caption + "o;"
intRotLength = txtRotLength.text
intRotLength = intRotLength + 1
If intRotLength < 0 Then
intRotLength = 0
End If
txtRotLength.text = intRotLength
Case "Adverts"
lblRotation.Caption = lblRotation.Caption + "a;"
intRotLength = txtRotLength.text
intRotLength = intRotLength + 1
If intRotLength < 0 Then
intRotLength = 0
End If
txtRotLength.text = intRotLength
Case "Voice_Tracks"
lblRotation.Caption = lblRotation.Caption + "v;"
intRotLength = txtRotLength.text
intRotLength = intRotLength + 1
If intRotLength < 0 Then
intRotLength = 0
End If
txtRotLength.text = intRotLength
End Select
End If

End Sub

Private Sub cmdRemove_Click()

Dim strCaption As String
Dim intRotLength

' Remove the latest addition from the rotation

If Len(lblRotation.Caption) >= 2 Then
strCaption = lblRotation.Caption
strCaption = Left$(strCaption, (Len(lblRotation.Caption) - 2))
lblRotation.Caption = strCaption
intRotLength = txtRotLength.text
intRotLength = intRotLength - 1
If intRotLength < 0 Then
intRotLength = 0
End If
txtRotLength.text = intRotLength
End If

End Sub

Private Sub cmdSave_Click()

Dim strLocation As String
Dim strMessage As String

If Len(lblRotation.Caption) > 0 Then
strLocation = App.Path & "\rotation.dat"
Kill strLocation    ' Destroy the old file
Open strLocation For Output As #1
Print #1, lblRotation.Caption
Print #1, txtRotLength.text
Close #1
Unload Me
frmAdminOptions.Visible = True
End If

If Len(lblRotation.Caption) = 0 Then
strMessage = "Cannot save a blank rotation"
Call Error_Handler(strMessage, 1)
End If

End Sub

Private Sub Form_Load()

Dim strLocation As String
Dim strRotation As String
Dim strRotLength As String

' Load the file

strLocation = App.Path + "\rotation.dat"
Open strLocation For Input As #1
Line Input #1, strRotation
Line Input #1, strRotLength
Close #1

' Output the data

lblRotation.Caption = strRotation
txtRotLength.text = strRotLength

End Sub




