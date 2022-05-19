VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAGC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Automatic Gain Control (AGC)"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   3720
      Width           =   975
   End
   Begin MSComctlLib.Slider sldTargetVolume 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Max             =   100
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldMinimumVolume 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Max             =   100
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldAdjustmentRate 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Max             =   500
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldDelay 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Max             =   200
      TickStyle       =   3
   End
   Begin VB.Label Label1 
      Caption         =   "Delay (ms)"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblAdjustmentRate 
      Caption         =   "Adjustment Rate [SLOW ---> FAST]"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label lblMinimum 
      Caption         =   "Minimum Volume (%)"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblTargetVolume 
      Caption         =   "Target Volume (%)"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAGC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

    ' Hide this screen and show the admin menu
    
    frmAdminOptions.Visible = True
    Unload Me

End Sub

Private Sub cmdSave_Click()

    Dim strFileName As String
    Dim intFreeFile As Integer
    
    ' Load the settings into the global settings
    
    GlobalAGC.fRate = sldAdjustmentRate.value / 10000
    GlobalAGC.fDelay = sldDelay.value / 1000
    GlobalAGC.fQuiet = sldMinimumVolume.value / 100
    GlobalAGC.fTarget = sldTargetVolume.value / 100
    GlobalAGC.fGain = 1
    
    ' Open the file
    
    strFileName = RPP(App.Path) & "agc.dat"
    intFreeFile = FreeFile
    Open strFileName For Random As intFreeFile
    
    ' Put the data into the file
    
    Put intFreeFile, 1, GlobalAGC
    
    ' Close the file
    
    Close intFreeFile
    
    ' Hide this screen
    
    frmAdminOptions.Visible = True
    Unload Me

End Sub

Private Sub Form_Load()

    ' Load the values from the current global settings
    
    sldAdjustmentRate.value = GlobalAGC.fRate * 10000
    sldDelay.value = GlobalAGC.fDelay * 1000
    sldMinimumVolume.value = GlobalAGC.fQuiet * 100
    sldTargetVolume.value = GlobalAGC.fTarget * 100

End Sub
