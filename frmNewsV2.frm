VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewsV2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "News"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRunNews 
      Alignment       =   1  'Right Justify
      Caption         =   "Run News?"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgMedia 
      Left            =   360
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdFileBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdJingleBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cmbInput 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3000
      Width           =   2895
   End
   Begin VB.ComboBox cmbSoundCard 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtFeedLength 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtLength 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtJingle 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2895
   End
   Begin VB.OptionButton optSource 
      Caption         =   "Sound Card"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.OptionButton optSource 
      Caption         =   "File"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Label lblSeconds 
      Alignment       =   2  'Center
      Caption         =   "s"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   19
      Top             =   840
      Width           =   135
   End
   Begin VB.Label lblSeconds 
      Alignment       =   2  'Center
      Caption         =   "s"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   18
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblInput 
      Caption         =   "Input"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Source:"
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
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblFeedLength 
      Alignment       =   1  'Right Justify
      Caption         =   "Feed Length:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblLength 
      Alignment       =   1  'Right Justify
      Caption         =   "Length:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblJingle 
      Alignment       =   1  'Right Justify
      Caption         =   "Jingle:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmNewsV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbSoundCard_Change()

    Call LoadInputs

End Sub

Private Sub cmbSoundCard_Click()

    Call LoadInputs

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdFileBrowse_Click()

    ' Prepare in case the user cancels
    
    On Error GoTo Exiter
    
    ' Load the file details
    
    dlgMedia.ShowOpen
    txtFile.text = dlgMedia.FileName
    
    ' Do nothing on an error
    
Exiter:

End Sub

Private Sub cmdJingleBrowse_Click()

    ' Prepare in case the user cancels
    
    On Error GoTo Exiter
    
    ' Load the file details
    
    dlgMedia.ShowOpen
    txtJingle.text = dlgMedia.FileName
    
    ' Do nothing on an error
    
Exiter:

End Sub

Private Sub cmdOK_Click()

    Dim strJingle As String
    Dim intJingle As Integer
    Dim intSource As Integer
    Dim boolFile As Boolean
    Dim strFile As String
    Dim intSoundCard As Integer
    Dim intInput As Integer
    Dim boolRunNews As Boolean
    
    Dim strFileName As String
    Dim intFreeFile As Integer
    
    ' Load the values into the variables
    
    strJingle = txtJingle.text
    On Error GoTo invalidJingleLength
    intJingle = txtLength.text
    On Error GoTo invalidFeedLength
    intSource = txtFeedLength.text
    
    boolFile = optSource(0).value
    
    If boolFile = True Then
        strFile = txtFile.text
    Else
        intSoundCard = cmbSoundCard.ListIndex + 1
        intInput = cmbInput.ListIndex
    End If
    
    boolRunNews = chkRunNews.value
    
    ' Validate the numbers
    
    If intJingle <= 0 Then
        Call Error_Handler("You have entered too small a time for the jingle length. Please enter a length between 1 and 3599.", 1)
        Exit Sub
    End If
    
    If intJingle > 3599 Then
        Call Error_Handler("You have entered too large a time for the jingle length. Please enter a length between 1 and 3599.", 1)
        Exit Sub
    End If
    
    If intSource <= 0 Then
        Call Error_Handler("You have entered too small a time for the feed length. Please enter a length between 1 and 3599.", 1)
        Exit Sub
    End If
    
    If intSource > 3599 Then
        Call Error_Handler("You have entered too large a time for the feed length. Please enter a length between 1 and 3599.", 1)
        Exit Sub
    End If
    
    ' Make sure that if soundcard is selected something has been selected
    ' in both dropdown boxes
    
    If boolFile = False Then
        If cmbSoundCard.text = "" Then
            Call Error_Handler("No sound card selected." & vbCrLf & vbCrLf & "Please select a sound card from the drop down box.", 1)
            Exit Sub
        End If
        If cmbInput.text = "" Then
            Call Error_Handler("No input selected." & vbCrLf & vbCrLf & "Please select and input from the drop down box.", 1)
            Exit Sub
        End If
    End If
    
    ' Open the file
    
    strFileName = App.Path & "\news.dat"
    intFreeFile = FreeFile
    Open strFileName For Output As #intFreeFile
    
    ' Write the values into the file
    
    Print #intFreeFile, strJingle
    Print #intFreeFile, intJingle
    Print #intFreeFile, intSource
    Print #intFreeFile, boolRunNews
    Print #intFreeFile, boolFile
    
    If boolFile = True Then
        Print #intFreeFile, strFile
    Else
        Print #intFreeFile, intSoundCard
        Print #intFreeFile, intInput
    End If
    
    ' Close the file
    
    Close intFreeFile
    
    ' Close this screen
    
    Unload Me
    Exit Sub
    
invalidJingleLength:
    
    Call Error_Handler("The jingle length you entered was not a number." & vbCrLf & vbCrLf & "Please enter a value between 1 and 3599.", 1)
    Exit Sub

invalidFeedLength:

    Call Error_Handler("The feed length you entered was not a number." & vbCrLf & vbCrLf & "Please enter a value between 1 and 3599.", 1)
    Exit Sub

End Sub

Private Sub Form_Load()

    Dim intLooper As Integer
    Dim strFileName As String
    Dim intFreeFile As Integer
    Dim strTemp As String
    Dim intTemp As Integer

    ' Set file filter
    
    dlgMedia.Filter = "MP3 Files (*.mp3)|*.mp3|" & _
                        "WMA Files (*.wma)|*.wma|" & _
                        "WAV Files (*.wav)|*.wav|" & _
                        "OGG Files (*.ogg)|*.ogg|" & _
                        "AIFF Files (*.aiff)|*.aiff|"
                        
    ' Set the default file path
    
    'dlgMedia.InitDir = App.Path
    
    ' Throw up an error if cancel is selected
    
    dlgMedia.CancelError = True
    
    ' Load the sound cards
    
    cmbSoundCard.Clear
    
    For intLooper = 0 To main.sgMain.MixerCount - 1
        cmbSoundCard.AddItem main.sgMain.Mixer(intLooper)
    Next intLooper
    
    Call LoadInputs
    
    cmbSoundCard.ListIndex = 0
    cmbInput.ListIndex = 0
    
    ' Load the default values

    strFileName = App.Path & "\news.dat"
    intFreeFile = FreeFile
    Open strFileName For Input As #intFreeFile
    
    On Error GoTo BLANKTEMP
    
    Line Input #intFreeFile, strTemp
    txtJingle.text = strTemp
    Line Input #intFreeFile, strTemp
    txtLength.text = strTemp
    Line Input #intFreeFile, strTemp
    txtFeedLength.text = strTemp
    Line Input #intFreeFile, strTemp
    If strTemp = True Then
        chkRunNews.value = 1
    Else
        chkRunNews.value = 0
    End If
    Line Input #intFreeFile, strTemp
    
    If strTemp = True Then
        optSource(0).value = True
        Line Input #intFreeFile, strTemp
        txtFile.text = strTemp
    Else
        optSource(1).value = True
        Line Input #intFreeFile, strTemp
        intTemp = strTemp
        cmbSoundCard.ListIndex = intTemp - 1
        Line Input #intFreeFile, strTemp
        intTemp = strTemp
        cmbInput.ListIndex = intTemp
    End If
    
    Close #intFreeFile
    Exit Sub
    
BLANKTEMP:
    
    strTemp = ""
    Resume Next

End Sub

Private Sub txtFile_Change()

    ' Redirect to browse button
    
    'cmdFileBrowse.SetFocus

End Sub

Private Sub txtJingle_Click()

    ' Redirect to browse button
    
    cmdJingleBrowse.SetFocus

End Sub

Sub LoadInputs()

    Dim intSoundCard As Integer
    Dim intLooper As Integer
    
    ' Clear the list
    
    cmbInput.Clear
    
    ' Get the sound card number
    
    intSoundCard = cmbSoundCard.ListIndex + 1
    
    ' Set the recording sound card
    
    main.sgMain.CurrentMixer = intSoundCard
    
    ' Show the inputs
    
    For intLooper = 0 To main.sgMain.OutputLineCount - 1
        cmbInput.AddItem main.sgMain.OutputLineName(main.sgMain.OutputLineID(intLooper))
    Next intLooper

End Sub
