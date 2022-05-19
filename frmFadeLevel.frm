VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFadeLevel 
   Caption         =   "Fade Level"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLevel 
      Interval        =   25
      Left            =   1320
      Top             =   1320
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin MSComctlLib.Slider sldLevel 
      Height          =   3015
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   5318
      _Version        =   393216
      Orientation     =   1
      Max             =   100
      TickStyle       =   3
   End
   Begin VB.Label lblCurrent 
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "0%"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbl100 
      Caption         =   "100%"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      Caption         =   "Fade Level"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmFadeLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()

    Dim strFileName As String
    Dim intFreeFile As Integer
    Dim intLevel As Integer
    
    intLevel = 100 - sldLevel.value
    
    ' Write the value to the file
    
    If Right(App.Path, 1) = "\" Then
        strFileName = App.Path & "fade.dat"
    Else
        strFileName = App.Path & "\" & "fade.dat"
    End If
    
    intFreeFile = FreeFile
    Open strFileName For Output As #intFreeFile
    
    Write #intFreeFile, intLevel
    
    Close #intFreeFile
    
    ' Set the global level
    
    FadeLevel = intLevel
    
    ' Hide this screen
    
    Unload Me

End Sub

Private Sub Form_Load()

    Dim strFileName As String
    Dim intFreeFile As Integer
    Dim intLevel As Integer
    Dim strTemp As String
    
    ' Write the value to the file
    
    If Right(App.Path, 1) = "\" Then
        strFileName = App.Path & "fade.dat"
    Else
        strFileName = App.Path & "\" & "fade.dat"
    End If
    
    intFreeFile = FreeFile
    Open strFileName For Input As #intFreeFile
    
    Line Input #intFreeFile, strTemp
    intLevel = strTemp
    
    Close #intFreeFile
    
    ' Set the onscreen
    
    sldLevel.value = 100 - intLevel

End Sub

Private Sub tmrLevel_Timer()

    Dim intLevel As Integer
    
    ' Display the current level
    
    intLevel = 100 - sldLevel.value
    lblCurrent.Caption = intLevel & "%"

End Sub
