VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C5BD6AEB-4923-11D4-8FFE-004F4C0058A2}#1.0#0"; "sguard.ocx"
Begin VB.Form main 
   Caption         =   "AllDay DJ"
   ClientHeight    =   8220
   ClientLeft      =   -2325
   ClientTop       =   270
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   548
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrVolume 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8400
      Top             =   4080
   End
   Begin SoundGuard.sGuard sgMain 
      Height          =   615
      Left            =   4200
      TabIndex        =   59
      Top             =   5520
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      RefreshRate     =   100
   End
   Begin VB.Timer tmrFade 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3960
      Top             =   5640
   End
   Begin VB.CommandButton cmdPlaylistLoad 
      Caption         =   "Load Playlist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   58
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdPlaylistSave 
      Caption         =   "Save Playlist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   57
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   10800
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox txtSongInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   6480
      Width           =   7575
   End
   Begin VB.TextBox txtSong 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox txtArtist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2640
      Width           =   3015
   End
   Begin VB.CommandButton cmdNextItem 
      Caption         =   "&Next Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Next Item"
      Top             =   5760
      Width           =   2535
   End
   Begin ComctlLib.Slider proSongTime 
      Height          =   375
      Left            =   4200
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   3600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   327682
      TickStyle       =   3
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   11040
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   11040
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   11040
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   11040
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   11040
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   11040
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   11040
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   11040
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   11040
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdPlayInstantPlayer 
      Caption         =   "No File Loaded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   8880
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlayInstantPlayer 
      Caption         =   "No File Loaded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   8880
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlayInstantPlayer 
      Caption         =   "No File Loaded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   8880
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlayInstantPlayer 
      Caption         =   "No File Loaded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   8880
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlayInstantPlayer 
      Caption         =   "No File Loaded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   8880
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlayInstantPlayer 
      Caption         =   "No File Loaded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8880
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlayInstantPlayer 
      Caption         =   "No File Loaded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   8880
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlayInstantPlayer 
      Caption         =   "No File Loaded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8880
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlayInstantPlayer 
      Caption         =   "No File Loaded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8880
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   600
      Width           =   2175
   End
   Begin VB.Timer tmrNews2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5400
      Top             =   5640
   End
   Begin VB.Timer tmrNews 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5880
      Top             =   5640
   End
   Begin VB.Timer tmrTimeDate 
      Interval        =   1
      Left            =   4920
      Top             =   5640
   End
   Begin ComctlLib.ListView lstPlaylist 
      Height          =   4815
      Left            =   120
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Artist"
         Object.Width           =   3492
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Track"
         Object.Width           =   3492
      EndProperty
   End
   Begin VB.CheckBox chkBreakAfter 
      Height          =   375
      Left            =   2880
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6120
      Width           =   255
   End
   Begin MSComDlg.CommonDialog cdlCommon 
      Left            =   6360
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstPreviousSongs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   6960
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5040
      Width           =   4815
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&STOP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2640
      Picture         =   "main.frx":044A
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Stop"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&PLAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1560
      Picture         =   "main.frx":08C0
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Play"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAutomate 
      Caption         =   "&AUTOMATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      MaskColor       =   &H00808080&
      Picture         =   "main.frx":0B0F
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Start Automatic Rotation"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdClearPlaylist 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      ToolTipText     =   "Clear Playlist"
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton cmdRemovePlaylist 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Remove From Playlist"
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton cmdAddPlaylist 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Add To Playlist"
      Top             =   5400
      Width           =   855
   End
   Begin VB.Timer tmrGetData 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4440
      Top             =   5640
   End
   Begin VB.Label lblDateTime 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4200
      TabIndex        =   32
      Top             =   7560
      Width           =   6375
   End
   Begin VB.Label lblBreakAfter 
      BackStyle       =   0  'Transparent
      Caption         =   "&Break After?"
      Height          =   615
      Left            =   1110
      TabIndex        =   29
      Top             =   6105
      Width           =   1815
   End
   Begin VB.Label lblPreviousSongs 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Songs"
      Height          =   375
      Left            =   4560
      TabIndex        =   27
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label lblPlayed_Time 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6960
      TabIndex        =   26
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblRemaining_Time 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6960
      TabIndex        =   25
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblPlaylist_Time 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7080
      TabIndex        =   24
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblNews_Time 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7080
      TabIndex        =   23
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblTimePlayed 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Time Played"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label lblTimeRemaining 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Time Remaining"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lblSongTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Song"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblArtistTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Artist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   19
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblCurrentSong 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current Song"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   18
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblPlaylistTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Time Left On Playlist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label lblNews 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Time Until News"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblF9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F9"
      Height          =   375
      Left            =   8520
      TabIndex        =   15
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label lblF8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F8"
      Height          =   375
      Left            =   8520
      TabIndex        =   14
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lblF7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F7"
      Height          =   375
      Left            =   8520
      TabIndex        =   13
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblF6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F6"
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblF5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F5"
      Height          =   375
      Left            =   8520
      TabIndex        =   11
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblInstantPlayers 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Instant Players"
      Height          =   375
      Left            =   9120
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblF4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F4"
      Height          =   375
      Left            =   8520
      TabIndex        =   9
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblF3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F2"
      Height          =   375
      Left            =   8520
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblF1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F1"
      Height          =   375
      Index           =   0
      Left            =   8520
      TabIndex        =   6
      Top             =   600
      Width           =   375
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu menuAdmin 
      Caption         =   "Admin Options"
   End
   Begin VB.Menu menuLogout 
      Caption         =   "&Logout"
      Visible         =   0   'False
   End
   Begin VB.Menu menuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu menuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' AllDay DJ v1.0.20
' Marc Steele (020774584)
' September 2005 - Jan 2007
' Radio automation system

' Screen resizing

Public intCurrentX As Integer
Public intCurrentY As Integer

' Fade in

Public boolFade As Boolean

' Playlist Movement

Dim boolClick As Boolean
Dim intPLItem As Integer
Dim AllowMainKeyboard As Boolean

' Selected item on playlist

Dim intGlobalSelected

' Tooltip variable

Dim mlToolTip As New clsToolTip



Private Sub cmdAddPlaylist_Click()

' Show the search page

frmSearch.Visible = False
frmSearch.Visible = True

End Sub

Public Sub cmdAutomate_Click()

Dim boolPlaying As Boolean
Dim boolSomethingOnPlaylist As Boolean

' Check if the main player is playing

If lblPlayed_Time.Caption <> "00:00" Then
boolPlaying = True
Else
boolPlaying = False
End If

' If it isn't playing then check if there is something on the playlist

If lstPlaylist.ListItems.Count > 1 Then
boolSomethingOnPlaylist = True
Else
boolSomethingOnPlaylist = False
End If

' If there is something on the playlist but not playing then play

If boolSomethingOnPlaylist And Not boolPlaying Then
Call cmdPlay_Click
End If

' Update the rotation data

If GlobalRotPos = -1 Then
GlobalRotPos = 0
End If

' Set the number of times the function has been called to zero
' This integer is used to prevent system lock-ups

RotTimes = 0


' The rest of the automation is handled by the Do_Automate subroutine.
' This subroutine is called where when then next song is selected
' It will be accessed now if there is nothing playing and there is nothing on the playlist

If Not boolSomethingOnPlaylist Or Not boolPlaying Then
Call Do_Automate
Call cmdPlay_Click
Call Do_Automate
End If

End Sub

Private Sub cmdClearPlaylist_Click()

' Clear the playlist

lstPlaylist.ListItems.Clear
frmPlayers.lstPlaylist.ListItems.Clear

' Update timers

Call Update_Timer

' Call the automation system if appropriate

If GlobalRotPos > -1 Then
Call Do_Automate
End If

End Sub

Private Sub cmdNextItem_Click()

Dim intTemp As Integer
Dim lngChannel As Long

' The sub adds the old item to the database, stops the current player and selects the new item
' It will only do this if there is an item playing, otherwise nothing will happen

If boolStarted = True And lstPlaylist.ListItems.Count > 0 Then
boolStarted = False
intTemp = GlobalPlayer
Call Do_Add_2hr
lngChannel = MainPlayers(GlobalPlayer)
Call BASS_ChannelSlideAttributes(lngChannel, -1, 0, -1, 2000)
Call cmdPlay_Click
If GlobalRotPos <> -1 Then
Call Do_Automate
End If
Call Pre_Load

Do While BASS_ChannelIsSliding(lngChannel) > 0
    DoEvents
Loop

Call BASS_StreamFree(lngChannel)
End If

End Sub

Private Sub cmdPlay_Click()

Dim intPlayerNumber As Integer
Dim intSelectedItem As Integer
Dim vntData(1 To 13) As Variant
Dim intLooper As Integer
Dim lngIntro As Long
Dim dblIntro As Integer
Dim lngChannel As Long
Dim strFile As String
Dim lngMax As Long
Dim lngTemp As Long
Dim dblRatio As Double

' Play start
' Used as a marker when there is no file

Play_Start:

' Check there is something to play

If lstPlaylist.ListItems.Count = 0 Then
Exit Sub
End If

' Quit if there is already something playing

If boolStarted = True Then
Exit Sub
End If

' Get the player number and choose the next one

intPlayerNumber = GlobalPlayer
intPlayerNumber = intPlayerNumber + 1
If intPlayerNumber = 4 Then
intPlayerNumber = 0
End If

GlobalPlayer = intPlayerNumber
intPlayer = intPlayerNumber

' Move the next item to the playing list

vntData(1) = frmPlayers.lstPlaylist.ListItems.item(1).text
For intLooper = 2 To 13
vntData(intLooper) = frmPlayers.lstPlaylist.ListItems.item(1).SubItems(intLooper - 1)
Next intLooper
frmPlayers.lstPlaylist.ListItems.Remove 1
frmPlayers.lstCurrent.ListItems.Clear
With frmPlayers.lstCurrent.ListItems.Add(, , vntData(1))
For intLooper = 2 To 13
.SubItems(intLooper - 1) = vntData(intLooper)
Next intLooper
End With

' Remove from the visible playlist

lstPlaylist.ListItems.Remove 1

' Load the data into the appropriate places

On Error Resume Next

proSongTime.min = 0
proSongTime.max = 0
proSongTime.max = Int(vntData(6))
proSongTime.min = Int(vntData(5))
txtArtist.text = vntData(2)
txtSong.text = vntData(3)
txtSongInfo.text = vntData(11)
strFile = vntData(7)

' Check that the file exists
' If it doesn't then load the next item

If FileExists(strFile) = False Then
frmPlayers.lstCurrent.ListItems.Clear
If GlobalRotPos > -1 Then
Call Do_Automate
End If
GoTo Play_Start
End If

' Start the player
' Play, get the intro position and jump
' Use seperate decoding for WMA

Call BASS_SetDevice(GetSoundCard(MAIN_SOUND_CARD))
If UCase$(Right$(strFile, 3)) <> "WMA" Then
lngChannel = BASS_StreamCreateFile(BASSFALSE, strFile, 0, 0, BASS_STREAM_AUTOFREE)
lngIntro = BASS_ChannelSeconds2Bytes(lngChannel, vntData(5))
Call BASS_ChannelSetPosition(lngChannel, lngIntro)
Call BASS_FX_DSP_Set(lngChannel, BASS_FX_DSPFX_DAMP, 0)
Call BASS_FX_DSP_SetParameters(lngChannel, BASS_FX_DSPFX_DAMP, GlobalAGC)
Call BASS_FX_DSP_Reset(lngChannel, BASS_FX_DSPFX_DAMP)
Call BASS_ChannelPlay(lngChannel, BASSFALSE)
Else
lngChannel = BASS_WMA_StreamCreateFile(BASSFALSE, strFile, 0, 0, BASS_STREAM_AUTOFREE)
lngIntro = BASS_ChannelSeconds2Bytes(lngChannel, vntData(5))
Call BASS_ChannelSetPosition(lngChannel, lngIntro)
Call BASS_FX_DSP_Set(lngChannel, BASS_FX_DSPFX_DAMP, 0)
Call BASS_FX_DSP_SetParameters(lngChannel, BASS_FX_DSPFX_DAMP, GlobalAGC)
Call BASS_FX_DSP_Reset(lngChannel, BASS_FX_DSPFX_DAMP)
Call BASS_ChannelPlay(lngChannel, BASSFALSE)
End If

Call Set_Compressor(lngChannel)

MainPlayers(intPlayer) = lngChannel
GlobalPlayer = intPlayer
IsFading(GlobalPlayer) = False

' Start the song timers

boolStarted = True
tmrGetData.Enabled = True
GlobalDate = Now()

' Live and loud

Call BASS_ChannelSetAttributes(lngChannel, -1, 100, -101)

End Sub

Private Sub cmdPlaylistLoad_Click()

Dim intLooper As Integer
Dim intItems As Integer
Dim strTemp As String
Dim strFileName As String

' Get the file

On Error GoTo NoFile

cdlCommon.InitDir = App.Path & "\Playlists"
cdlCommon.Filter = "AllDay DJ Playlist (*.addpl)|*.addpl|"
cdlCommon.CancelError = True
cdlCommon.ShowOpen
strFileName = cdlCommon.FileName

' Check there is something selected

If strFileName = "" Then
Exit Sub
End If

' Open the file

Open strFileName For Input As #1
Line Input #1, strTemp

' Get the number of items and create arrays

intItems = strTemp
ReDim strID(1 To intItems) As String
ReDim strType(1 To intItems) As String

' Retrieve the data

For intLooper = 1 To intItems
Line Input #1, strID(intLooper)
Line Input #1, strType(intLooper)
Next intLooper

' Close the file

Close #1

' Add to the current playlist

For intLooper = 1 To intItems
Call doAddToPlaylist(strID(intLooper), strType(intLooper))
Next intLooper

Exit Sub

' Error handling

NoFile:

End Sub

Private Sub cmdPlaylistSave_Click()

Dim intItems As Integer
Dim intLooper As Integer
Dim strFileName As String

' Check there are items to save

intItems = lstPlaylist.ListItems.Count
If intItems = 0 Then
Call Error_Handler("Cannot save an empty playlist!", 1)
Exit Sub
End If

' Get the addpl file

On Error GoTo NoFile

cdlCommon.InitDir = App.Path & "\Playlists"
cdlCommon.Filter = "AllDay DJ Playlist (*.addpl)|*.addpl|"
cdlCommon.CancelError = True
cdlCommon.ShowSave
strFileName = cdlCommon.FileName

' Exit if nothing was selected

If strFileName = "" Then
Exit Sub
End If

' Check if the file already exists

If FileExists(strFileName) Then
Kill strFileName
End If

' Retrieve the items from the playlist

ReDim intID(1 To intItems) As Integer
ReDim strType(1 To intItems) As String

For intLooper = 1 To intItems
intID(intLooper) = frmPlayers.lstPlaylist.ListItems.item(intLooper).text
strType(intLooper) = frmPlayers.lstPlaylist.ListItems.item(intLooper).SubItems(12)
Next intLooper

' Output to the file

Open strFileName For Output As #1
Print #1, intItems

For intLooper = 1 To intItems
Print #1, intID(intLooper)
Print #1, strType(intLooper)
Next intLooper

Close #1

' Error handling

NoFile:

End Sub

Private Sub cmdRemovePlaylist_Click()

Dim strMessage As String
Dim intSelected As Integer

' Remove the track from the playlist only if a file is selected

On Error GoTo NoItemSelected

If lstPlaylist.SelectedItem = "" Then
Exit Sub
End If

' The the item number

intSelected = lstPlaylist.SelectedItem.Index
lstPlaylist.ListItems.Remove intSelected
frmPlayers.lstPlaylist.ListItems.Remove intSelected

' Update timers

Call Update_Timer

' Call the automation system if appropriate

If GlobalRotPos > -1 Then
Call Do_Automate
End If

' Preload the next item

If lstPlaylist.ListItems.Count > 0 Then
Call Pre_Load
End If

Exit Sub

' Error handling

NoItemSelected:

strMessage = "Cannot remove nothing!" & Chr(10) & Chr(10) & "Please select an item then select remove again."
Call Error_Handler(strMessage, 1)
End Sub

Private Sub cmdSave_Click()

Dim dbConnection As New ADODB.Connection
Dim strText As String
Dim strConnectionString As String

' Get the data from the form

strText = txtSongInfo.text
If Len(strText) > 65536 Then
Call Error_Handler("Too much data." & Chr(10) & Chr(10) & "Please shorten the information.", 1)
Exit Sub
End If

' Exit the subroutine if there is no song playing

If lblPlayed_Time.Caption = "00:00" Then
Exit Sub
End If

' Update the database

On Error GoTo We_Have_Apostrophies

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb;Persist Security Info=False"
dbConnection.Open strConnectionString
strSQL = "UPDATE " & frmPlayers.lstCurrent.ListItems.item(1).SubItems(12) & " SET [extrainfo] = """ & strText & """ WHERE [itemid] = " & frmPlayers.lstCurrent.ListItems.item(1).text
dbConnection.Execute strSQL
dbConnection.Close

Exit Sub

We_Have_Apostrophies:

Call Error_Handler("The information you tried to enter has an apostrophie ("") in it." & Chr(10) & Chr(10) & "Please remove the apostrophies and try saving again.", 1)

End Sub


Private Sub Form_Load()

' Inititate Varaibles

Dim ctrlAll As Control
Dim strStationID As String
Dim strFileLocation As String
Dim strConnetionString As String
Dim tmpAppPath As String
Dim dbConnection As New ADODB.Connection
Dim strConnectionString As String
Dim lngStream As Long
Dim lngLineInChannel As Long
Dim lngPlugin As Long
Dim strPlugin As String

' The admin has not logger in

AdminLoggedIn = False
AllowMainKeyboard = True

' We are not fading the tracks out

For intLooper = 0 To 3
    IsFading(intLooper) = False
Next intLooper

' Randomize

Randomize

' Set the boolean value

boolStarted = False

' Open the station.txt with the station name

strFileLocation = App.Path + "\station.dat"
Open strFileLocation For Input As #1
Line Input #1, strStationID
Close #1

' Set the station name

main.Caption = main.Caption + " - " + strStationID

' Prepare for all eventualities

On Error Resume Next

' Set default font

For Each ctrlAll In main.Controls
Let ctrlAll.Font.Name = "Comic Sans MS"
Next ctrlAll

' Set screen size

intCurrentX = Me.Width
intCurrentY = Me.Height
Me.Width = Screen.Width
Me.Height = Screen.Height

' Change to digital readout for timer labels

lblNews_Time.Font.Name = "Digital Readout"
lblPlayed_Time.Font.Name = "Digital Readout"
lblPlaylist_Time.Font.Name = "Digital Readout"
lblRemaining_Time.Font.Name = "Digital Readout"

' Clear the 2 hour log

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\logs.mdb;Persist Security Info=False"
dbConnection.Open strConnectionString
dbConnection.Execute "DELETE FROM 2hr"
dbconection.Close

' Initialise bass.dll using default sound card

If Not FileExists(RPP(App.Path) & "bass.dll") Then
Call Error_Handler("BASS.DLL does not exist.", 0)
Unload Me
End If
If BASS_Init(GetSoundCard(MAIN_SOUND_CARD), 44100, 0, Me.hwnd, 0) = BASSFALSE Then
Call Error_Handler("Cannot initialise sound system." & Chr(10) & Chr(10) & "Will now exit.", 0)
Unload Me
End If

If GetSoundCard(PFL_SOUND_CARD) <> GetSoundCard(MAIN_SOUND_CARD) Then
If BASS_Init(GetSoundCard(PFL_SOUND_CARD), 44100, 0, Me.hwnd, 0) = BASSFALSE Then
Call Error_Handler("Cannot initialise sound system." & Chr(10) & Chr(10) & "Will now exit.", 0)
Unload Me
End If
End If

' Set volume

Call BASS_SetVolume(100)

' Playlist movement

boolClick = False

' Load the AGC settings

Call GetAGCSettings

' Initialize the recording function

intLooper = 1
Do While BASS_GetDeviceDescription(intLooper) <> 0
    Call BASS_RecordInit(intLooper)
    intLooper = intLooper + 1
Loop

Set GlobalVolume = New clsVolume

' Automation

GlobalRotPos = -1

' Current player

GlobalPlayer = 0

' Main players

For intLooper = 0 To 3
    MainPlayers(intLooper) = 0
Next intLooper

' Load the fader level

Call loadFade

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

boolClick = False

End Sub

Private Sub Form_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Call LVDragDropSingle(lstPlaylist, x, y)
End Sub

Private Sub Form_Resize()

Dim intNewX As Integer
Dim intNewY As Integer
Dim dblScaleFactorX As Double
Dim dblScaleFactorY As Double
Dim ctrlAll As Control

' Update variables

intNewX = Me.Width
intNewY = Me.Height

If intNewY > 1000 Then
dblScaleFactorX = intNewX / intCurrentX
dblScaleFactorY = intNewY / intCurrentY
intCurrentX = Me.Width
intCurrentY = Me.Height

On Error Resume Next

' Resize the controls

For Each ctrlAll In Me.Controls
ctrlAll.Width = ctrlAll.Width * dblScaleFactorX
ctrlAll.Left = ctrlAll.Left * dblScaleFactorX
ctrlAll.Font.Size = ctrlAll.Font.Size * dblScaleFactorY
ctrlAll.Height = ctrlAll.Height * dblScaleFactorY
ctrlAll.Top = ctrlAll.Top * dblScaleFactorY
Next ctrlAll
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Stop all players

Call cmdStop_Click
Call BASS_RecordFree
Call BASS_Stop

' Unload all other forms

For Each frmFormsToUnload In Forms
Unload frmFormsToUnload
Next frmFormsToUnload

End Sub

Private Sub lstPlaylist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
lstPlaylist.SelectedItem = lstPlaylist.HitTest(x, y)
End Sub

Private Sub lstPlaylist_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim intItemNumber As Integer
    Dim strText As String
    
    On Error GoTo Exiter
    
    ' Get the item number
    
    intItemNumber = lstPlaylist.HitTest(x, y).Index
    
    ' Generate the string
    
    strText = "Artist: " & frmPlayers.lstPlaylist.ListItems.item(intItemNumber).SubItems(1) & vbCrLf
    strText = strText & "Title: " & frmPlayers.lstPlaylist.ListItems.item(intItemNumber).SubItems(2) & vbCrLf
    strText = strText & "Length: " & SecondsToMins(frmPlayers.lstPlaylist.ListItems.item(intItemNumber).SubItems(3)) & vbCrLf
    strText = strText & "Intro: " & Int(frmPlayers.lstPlaylist.ListItems.item(intItemNumber).SubItems(11)) & "s"
    
    ' Set the popup
    
    mlToolTip.Create Me
    mlToolTip.MaxTipWidth = 240
    mlToolTip.DelayTime(ttDelayShow) = 20000
        
    mlToolTip.AddTool lstPlaylist
    mlToolTip.ToolText(lstPlaylist) = strText
    
    Exit Sub
    
    ' Error processing
    
Exiter:
    
    lstPlaylist.ToolTipText = ""

End Sub

Private Sub lstPlaylist_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
lstPlaylist.MousePointer = ccDefault
Call LVDragDropSingle(lstPlaylist, x, y)
End Sub

Private Sub menuAbout_Click()

' Initiate Variables

Dim strAboutMessage As String

' Display about message

Let strAboutMessage = "AllDay DJ v" & App.Major & "." & App.Minor & "." & App.Revision & " (Full Version)" + Chr(10) + Chr(10) + "Copyright © 2005 - 2006 Marc Steele." + Chr(10) + Chr(10) + "Powered by BASS (www.un4seen.com)"
MsgBox strAboutMessage, vbOKOnly Or vbInformation, "About AllDay DJ"

End Sub

Private Sub menuAdmin_Click()

' Show the appropriate screen

If AdminLoggedIn = False Then
    frmLogin.Visible = True
Else
    frmAdminOptions.Visible = True
End If

End Sub

Private Sub menuExit_Click()

' Close all forms

For Each frmFormsToUnload In Forms
Unload frmFormsToUnload
Next frmFormsToUnload

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

' Uses function keys to start instant players

Select Case KeyCode
Case vbKeyF1:
Call cmdPlayInstantPlayer_Click(0)
Case vbKeyF2:
Call cmdPlayInstantPlayer_Click(1)
Case vbKeyF3:
Call cmdPlayInstantPlayer_Click(2)
Case vbKeyF4:
Call cmdPlayInstantPlayer_Click(3)
Case vbKeyF5:
Call cmdPlayInstantPlayer_Click(4)
Case vbKeyF6:
Call cmdPlayInstantPlayer_Click(5)
Case vbKeyF7:
Call cmdPlayInstantPlayer_Click(6)
Case vbKeyF8:
Call cmdPlayInstantPlayer_Click(7)
Case vbKeyF9:
Call cmdPlayInstantPlayer_Click(8)

' Instant player keys on numpad

Case vbKeyNumpad1:
Call cmdPlayInstantPlayer_Click(0)
Case vbKeyNumpad2:
Call cmdPlayInstantPlayer_Click(1)
Case vbKeyNumpad3:
Call cmdPlayInstantPlayer_Click(2)
Case vbKeyNumpad4:
Call cmdPlayInstantPlayer_Click(3)
Case vbKeyNumpad5:
Call cmdPlayInstantPlayer_Click(4)
Case vbKeyNumpad6:
Call cmdPlayInstantPlayer_Click(5)
Case vbKeyNumpad7:
Call cmdPlayInstantPlayer_Click(6)
Case vbKeyNumpad8:
Call cmdPlayInstantPlayer_Click(7)
Case vbKeyNumpad9:
Call cmdPlayInstantPlayer_Click(8)

' Main player keys

Case vbKeyAdd:
    Call cmdPlay_Click
Case vbKeyMultiply:
    If chkBreakAfter.value = 0 Then
        chkBreakAfter.value = 1
    Else
        chkBreakAfter.value = 0
    End If
Case vbKeyDecimal:
    Call cmdNextItem_Click
Case vbKeyNumpad0:
    Call cmdAutomate_Click
Case vbKeySubtract:
    Call cmdStop_Click


End Select
End Sub

Private Sub cmdStop_Click()

Dim intLooper As Integer
Dim lngChannel As Long

' Stop the data updater

tmrGetData.Enabled = False

' Add to the 2 hour log

If lblPlayed_Time.Caption <> "00:00" Then
Call Do_Add_2hr
End If

' Stop the players

On Error Resume Next

Call BASS_SetDevice(GetSoundCard(MAIN_SOUND_CARD))
For intLooper = 0 To 3
lngChannel = MainPlayers(intLooper)
Call BASS_ChannelSlideAttributes(lngChannel, -1, 0, -1, 2000)
Debug.Print BASS_ChannelIsSliding(lngChannel)
Do While BASS_ChannelIsSliding(lngChannel) > 0
    ' Do nothing
Loop
Call BASS_StreamFree(lngChannel)
Next intLooper

For intLooper = 0 To 8
lngChannel = InstantPlayers(intLooper)
Call BASS_ChannelStop(lngChannel)
Next intLooper

' Clear all timers

lblPlayed_Time.Caption = "00:00"
lblRemaining_Time.Caption = "00:00"
txtArtist.text = ""
txtSong.text = ""
proSongTime.value = 0
boolStarted = False
txtSongInfo.text = ""

' Stop automation

GlobalRotPos = -1

' Preload if appropriate and show intro length

If lstPlaylist.ListItems.Count > 0 Then
Call Pre_Load
dblIntro = frmPlayers.lstPlaylist.ListItems.item(1).SubItems(11)
dblIntro = Int(dblIntro - dblCurrentPosition)
lblRemaining_Time.Caption = dblIntro
lblRemaining_Time.BackColor = vbGreen
lblRemaining_Time.ForeColor = vbBlack
lblTimeRemaining.Caption = "Intro Countdown"
Else
lblRemaining_Time.BackColor = vbBlack
lblRemaining_Time.ForeColor = vbRed
lblTimeRemaining.Caption = "Time Remaining"
lblRemaining_Time.Caption = "00:00"
End If

End Sub


Private Sub cmdPlayInstantPlayer_Click(Index As Integer)

Dim lngStream As Long
Dim intLooper As Integer
Dim boolSliding As Boolean

' Play the appropriate string

On Error GoTo Sub_Ender

Call BASS_SetDevice(GetSoundCard(MAIN_SOUND_CARD))
lngStream = InstantPlayers(Index)
Call BASS_ChannelSetPosition(lngStream, 0)
Call BASS_FX_DSP_Set(lngStream, BASS_FX_DSPFX_DAMP, 100)
Call BASS_FX_DSP_SetParameters(lngStream, BASS_FX_DSPFX_DAMP, GlobalAGC)
Call BASS_ChannelSetAttributes(lngStream, -1, 100, -101)
Call BASS_ChannelPlay(lngStream, 0)

    ' Now reduce volume on the main players

    For intLooper = 0 To 3
        Call BASS_ChannelSlideAttributes(MainPlayers(intLooper), -1, 50, -101, 500)
    Next intLooper
    
    ' Do not move on until faded down
    
    boolSliding = True
    
    While boolSliding
        DoEvents
        boolSliding = False
        For intLooper = 0 To 3
            If BASS_ChannelIsSliding(MainPlayers(intLooper)) = BASSTRUE Then
                boolSliding = True
            End If
        Next intLooper
    Wend
    
    ' Wait for the jingle to finish before fade up
    
    tmrVolume.tag = lngStream
    tmrVolume.Enabled = True

Sub_Ender:
End Sub

Sub cmdLoad_Click(intPlayer As Integer)

' Initiate Variables

Dim strFileToLoad As String
Dim strFileTitle As String
Dim intFileTitleLength As Integer
Dim lngStream As Long

' Use common controls for opening file

On Error GoTo LeaveSub

'cdlCommon.InitDir = App.Path
cdlCommon.Filter = "MP3 Files (*.mp3)|*.mp3|" & _
                   "WMA Files (*.wma)|*.wma|" & _
                   "WAV Files (*.wav)|*.wav|" & _
                   "OGG Files (*.ogg)|*.ogg|" & _
                   "AIFF Files (*.aiff)|*.aiff|"
cdlCommon.CancelError = True
cdlCommon.ShowOpen
strFileToLoad = cdlCommon.FileName
strFileTitle = cdlCommon.FileTitle

' Load the file into the instant player

If UCase$(Right$(strFileToLoad, 3)) = "WMA" Then
lngStream = BASS_WMA_StreamCreateFile(BASSFALSE, strFileToLoad, 0, 0, 0)
frmInstantPlayers.txtInstantPlayer(intPlayer).text = lngStream
Else
lngStream = BASS_StreamCreateFile(BASSFALSE, strFileToLoad, 0, 0, 0)
InstantPlayers(intPlayer) = lngStream
End If

' Load the name onto the player button

intFileTitleLength = Len(strFileTitle)
If intFileTitleLength > 0 Then
intFileTitleLength = intFileTitleLength - 4
strFileTitle = Left$(strFileTitle, intFileTitleLength)
cmdPlayInstantPlayer(intPlayer).Caption = strFileTitle
End If

' Set the font size

If intFileTitleLength > 20 Then
    cmdPlayInstantPlayer(intPlayer).FontSize = cmdLoad(intPlayer).FontSize / 2
Else
    cmdPlayInstantPlayer(intPlayer).FontSize = cmdLoad(intPlayer).FontSize
End If

' Error handler for cancel

LeaveSub:

End Sub



Private Sub menuHelp_Click()

Dim dblTemp As Double

  On Error GoTo Handler

  ' Load PDF file
  
  dblTemp = Shell("AllDayDJ.pdf", , App.Path)
  Exit Sub

  ' Error handler
  
Handler:
  
  MsgBox "Error Number:" & Err.Number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & vbCrLf & "To view the help file, select help from the Start Menu.", vbOKOnly + vbInformation, "Error"

End Sub



Private Sub mnuOB_Click()

frmOB.Visible = True

End Sub

Private Sub menuLogout_Click()

    ' Log the admin out
    
    AdminLoggedIn = False
    menuLogout.Visible = False

End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub tmrFade_Timer()

Dim intCurrent As Integer
Dim intVT As Integer
Dim intLevel As Integer
Dim intLooper As Integer

' Get values

intLevel = tmrFade.tag
intCurrent = GlobalPlayer

Select Case intCurrent
Case 0
intVT = 3
Case 1
intVT = 0
Case 2
intVT = 1
Case 3
intVT = 2
End Select

' Change the volume

If BASS_ChannelIsActive(MainPlayers(intVT)) = BASS_ACTIVE_STOPPED Then
    For intLooper = 0 To 3
        If Not IsFading(intLooper) Then
            Call BASS_ChannelSlideAttributes(MainPlayers(intLooper), -1, 100, -1, 1000)
        End If
    Next intLooper
    tmrFade.Enabled = False
Else
    For intLooper = 0 To 3
        If intLooper <> intVT And IsFading(intLooper) = False Then
            Call BASS_ChannelSetAttributes(MainPlayers(intLooper), -1, FadeLevel, -1)
        End If
    Next intLooper
End If

End Sub

Private Sub tmrGetData_Timer()

Dim dblCurrentPosition As Double
Dim dblLength As Double
Dim dblTimeLeft As Double
Dim dblIntro As Double
Dim strTime As String
Dim intMinutes As Integer
Dim intSeconds As Integer
Dim strMinutes As String
Dim strSeconds As String
Dim intHours As Integer
Dim strHours As String
Dim intNoOfOtherTracks As Integer
Dim intPlayer As Integer
Dim lngStream As Long
Dim vntData As Variant
Dim CurrentTime As Variant

CurrentTime = Now()
boolFade = False

' Use redim to get the data about the other tracks in the playlist

intNoOfOtherTracks = frmPlayers.lstPlaylist.ListItems.Count
ReDim dblOtherTrackTimes(intNoOfOtherTracks) As Double

' Get the current position, time left  and the length of the song

vntData = MainPlayers(GlobalPlayer)
lngStream = vntData
dblCurrentPosition = BASS_ChannelGetPosition(lngStream)
dblCurrentPosition = BASS_ChannelBytes2Seconds(lngStream, dblCurrentPosition)
dblLength = proSongTime.max

If dblLength <= dblCurrentPosition Then
dblCurrentPosition = dblLength
End If
If dblCurrentPosition < 0 Then
dblCurrentPosition = 0
End If

dblTimeLeft = dblLength - dblCurrentPosition

' Process the data to get the current time passed

proSongTime.value = dblCurrentPosition
intMinutes = Int(dblCurrentPosition / 60)
intSeconds = Int(dblCurrentPosition - (intMinutes * 60))

If intMinutes < 10 Then
strMinutes = "0" & intMinutes
End If

If intMinutes >= 10 Then
strMinutes = intMinutes
End If

If intSeconds < 10 Then
strSeconds = "0" & intSeconds
End If

If intSeconds >= 10 Then
strSeconds = intSeconds
End If

lblPlayed_Time.Caption = strMinutes & ":" & strSeconds

' Do the same again to get the time left

dblTimeLeft = dblLength - dblCurrentPosition
intMinutes = Int(dblTimeLeft / 60)
intSeconds = Int(dblTimeLeft - (intMinutes * 60))

If intMinutes < 10 Then
strMinutes = "0" & intMinutes
End If

If intMinutes >= 10 Then
strMinutes = intMinutes
End If

If intSeconds < 10 Then
strSeconds = "0" & intSeconds
End If

If intSeconds >= 10 Then
strSeconds = intSeconds
End If

lblRemaining_Time.Caption = strMinutes & ":" & strSeconds
lblRemaining_Time.BackColor = vbBlack
lblRemaining_Time.ForeColor = vbRed
lblTimeRemaining.Caption = "Time Remaining"

' Voice tracking

On Error Resume Next

If lstPlaylist.ListItems.Count > 0 Then
dblIntro = Round(frmPlayers.lstPlaylist.ListItems.item(1).SubItems(11), 0) - Round(frmPlayers.lstPlaylist.ListItems.item(1).SubItems(4), 0)
If LCase$(frmPlayers.lstCurrent.ListItems.item(1).SubItems(12)) = "voice_tracks" And dblTimeLeft <= dblIntro Then
dblCurrentPosition = dblLength

' Inititate the fade timer

tmrFade.tag = 80
boolFade = True

End If
End If

' Get the remaining playlist time

For intLooper = 1 To intNoOfOtherTracks
dblTimeLeft = dblTimeLeft + frmPlayers.lstPlaylist.ListItems.item(intLooper).SubItems(3)
Next intLooper

' Change the times to integers

intHours = Int(dblTimeLeft / 3600)
intMinutes = Int((dblTimeLeft / 60) - (intHours * 60))
intSeconds = Int(dblTimeLeft Mod 3600)

If intSeconds >= 60 Then
Do While intSeconds >= 60
intSeconds = intSeconds - 60
Loop
End If

' Change times to strings and display

If intMinutes < 10 Then
strMinutes = "0" & intMinutes
End If

If intMinutes >= 10 Then
strMinutes = intMinutes
End If

If intSeconds < 10 Then
strSeconds = "0" & intSeconds
End If

If intSeconds >= 10 Then
strSeconds = intSeconds
End If

If intHours < 10 Then
strHours = "0" & intHours
End If

If intHours >= 10 Then
strHours = intHours
End If

lblPlaylist_Time.Caption = strHours & ":" & strMinutes & ":" & strSeconds

' Get the intro

dblIntro = frmPlayers.lstCurrent.ListItems.item(1).SubItems(11)
If dblCurrentPosition < dblIntro Then
dblIntro = Int(dblIntro - dblCurrentPosition)
lblRemaining_Time.Caption = dblIntro
lblRemaining_Time.BackColor = vbGreen
lblRemaining_Time.ForeColor = vbBlack
lblTimeRemaining.Caption = "Intro Countdown"
End If

' Check if remaining time is zero and if so, loads the next song
' Calling the preload for the next item

If dblCurrentPosition = dblLength Or BASS_ChannelIsActive(lngStream) = 0 Then
Call Do_Add_2hr
tmrGetData.Enabled = False
boolStarted = False

' Stop current stream if appropriate
' Otherwise a 2 second fade

If frmPlayers.lstCurrent.ListItems.item(1).SubItems(7) = False Then
lngChannel = MainPlayers(GlobalPlayer)
Call BASS_StreamFree(lngChannel)
Else
If LCase$(frmPlayers.lstCurrent.ListItems.item(1).SubItems(12)) <> "voice_tracks" Then
lngChannel = MainPlayers(GlobalPlayer)
IsFading(GlobalPlayer) = True
Call BASS_ChannelSlideAttributes(lngChannel, -1, 0, -1, 2000)
End If
End If

' Show the countdown timer

If frmPlayers.lstPlaylist.ListItems.Count > 0 Then
dblIntro = frmPlayers.lstPlaylist.ListItems.item(1).SubItems(11)
dblIntro = Int(dblIntro - dblCurrentPosition)
lblRemaining_Time.Caption = dblIntro
lblRemaining_Time.BackColor = vbGreen
lblRemaining_Time.ForeColor = vbBlack
lblTimeRemaining.Caption = "Intro Countdown"
End If

' Break After

If chkBreakAfter.value = 0 Then
Call cmdPlay_Click
Else
chkBreakAfter.value = 0
End If

' Automation

If GlobalRotPos > -1 Then
Call Do_Automate
End If
Call Pre_Load
End If

' Check for fade out

    If BASS_ChannelIsSliding(lngChannel) = BASSTRUE Then
        While (BASS_ChannelIsSliding(lngChannel) = BASSTRUE)
            DoEvents
        Wend
        Call BASS_StreamFree(lngChannel)
    End If

End Sub

Sub Do_Add_2hr()

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strConnectionString As String
Dim strLength As String
Dim dblLength As Double
Dim intMinutes As Integer
Dim intSeconds As Integer
Dim strSeconds As String
Dim strMinutes As String
Dim intHours As Integer
Dim strHours As String
Dim dblTimeDiff As Double
Dim strTimeDiff As String
Dim vntData As Variant
Dim subItems_8 As Variant
Dim subItems_9 As Variant

' Get the song length and convert to string

dblLength = proSongTime.max
intHours = Int(dblLength / 3600)
intMinutes = Int((dblLength / 60) - (intHours * 60))
intSeconds = Int(dblLength - (intMinutes * 60) - (intHours * 3600))

If intMinutes < 10 Then
strMinutes = "0" & intMinutes
End If

If intMinutes >= 10 Then
strMinutes = intMinutes
End If

If intSeconds < 10 Then
strSeconds = "0" & intSeconds
End If

If intSeconds >= 10 Then
strSeconds = intSeconds
End If

If intHours < 10 Then
strHours = "0" & intHours
End If

If intHours >= 10 Then
strHours = intHours
End If

strLength = strHours & ":" & strMinutes & ":" & strSeconds

' Get the length played and convert to time

dblTimeDiff = DateDiff("s", GlobalDate, Now())
dblTimeDiff = Round(dblTimeDiff)

' Assume clocks moved back if time is negative

Do While dblTimeDiff < 0
dblTimeDiff = dblTimeDiff + 3600
Loop

' As we cannot play a song for logner than it actually is, we will make the times the same

If dblTimeDiff > dblLength Then
dblTimeDiff = dblLength
End If

intHours = Int(dblTimeDiff / 3600)
intMinutes = Int((dblTimeDiff / 60) - (intHours * 60))
intSeconds = Int(dblTimeDiff - (intMinutes * 60) - (intHours * 3600))

If intMinutes < 10 Then
strMinutes = "0" & intMinutes
End If

If intMinutes >= 10 Then
strMinutes = intMinutes
End If

If intSeconds < 10 Then
strSeconds = "0" & intSeconds
End If

If intSeconds >= 10 Then
strSeconds = intSeconds
End If

If intHours < 10 Then
strHours = "0" & intHours
End If

If intHours >= 10 Then
strHours = intHours
End If

strTimeDiff = strHours & ":" & strMinutes & ":" & strSeconds

' Add to log

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\logs.mdb"
dbConnection.Open strConnectionString

' If not an advert then insert into the main log and 2 hour log

subItems_8 = frmPlayers.lstCurrent.ListItems.item(1).SubItems(8)
subItems_9 = frmPlayers.lstCurrent.ListItems.item(1).SubItems(9)

If subItems_8 = "" Then
subItems_8 = "N/A"
End If

If subItems_9 = "" Then
subItems_9 = "N/A"
End If

If frmPlayers.lstCurrent.ListItems.item(1).SubItems(12) <> "adverts" And frmPlayers.lstCurrent.ListItems.item(1).SubItems(12) <> "voice_tracks" Then
strSQL = "INSERT INTO [2hr] VALUES (""" & frmPlayers.lstCurrent.ListItems.item(1).SubItems(1) & """, """ & frmPlayers.lstCurrent.ListItems.item(1).SubItems(2) & """, " & frmPlayers.lstCurrent.ListItems.item(1).text & ", " & dblLength & ", " & dblTimeDiff & ", """ & strLength & """, """ & strTimeDiff & """, """ & GlobalDate & """, """ & subItems_8 & """, """ & subItems_9 & """)"
dbConnection.Execute strSQL
strSQL = "INSERT INTO [main] VALUES (""" & frmPlayers.lstCurrent.ListItems.item(1).SubItems(1) & """, """ & frmPlayers.lstCurrent.ListItems.item(1).SubItems(2) & """, " & frmPlayers.lstCurrent.ListItems.item(1).text & ", " & dblLength & ", " & dblTimeDiff & ", """ & strLength & """, """ & strTimeDiff & """, """ & GlobalDate & """, """ & subItems_8 & """, """ & subItems_9 & """)"
dbConnection.Execute strSQL
End If

' If an advert then insert into the advert log

If frmPlayers.lstCurrent.ListItems.item(1).SubItems(12) = "adverts" Then
strSQL = "INSERT INTO [adverts] VALUES (""" & frmPlayers.lstCurrent.ListItems.item(1).SubItems(1) & """, """ & frmPlayers.lstCurrent.ListItems.item(1).SubItems(2) & """, " & frmPlayers.lstCurrent.ListItems.item(1).text & ", " & dblLength & ", " & dblTimeDiff & ", """ & strLength & """, """ & strTimeDiff & """, """ & GlobalDate & """, """ & subItems_8 & """, """ & subItems_9 & """)"
dbConnection.Execute strSQL
End If

' Remove old entries from the 2 hour log

strSQL = "SELECT [when_played] FROM [2hr]"
dbRecordset.Open strSQL, dbConnection
Do While Not dbRecordset.EOF
For Each vntData In dbRecordset.Fields
If DateDiff("s", Now(), vntData) > 7200 Then
strSQL = "DELETE FROM [2hr] WHERE [when_played] = " & vntData
dbConnection.Execute strSQL
End If
Next vntData
dbRecordset.MoveNext
Loop

' Load the latest data to the just played section

lstPreviousSongs.AddItem (frmPlayers.lstCurrent.ListItems.item(1).SubItems(1) & " - " & frmPlayers.lstCurrent.ListItems.item(1).SubItems(2))
txtSong.text = ""
txtArtist.text = ""
lblPlayed_Time.Caption = "00:00"
lblRemaining_Time.Caption = "00:00"

' Restrict the just played section to 3 items

If lstPreviousSongs.ListCount > 3 Then
lstPreviousSongs.RemoveItem 0
End If

' Close the connection

dbRecordset.Close
dbConnection.Close

' Inititate the fade

If boolFade Then
tmrFade.Enabled = True
End If

End Sub

Sub check_ad_break(ByVal strNow As String)

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim vntData As Variant
Dim strConnectionString As String
Dim boolAdvert As Boolean
Dim strSQL As String
Dim intNoOfAds As Integer
Dim intLooper As Integer
Dim intLooper2 As Integer
Dim intOtherItems As Integer

' Set initial values

boolAdvert = False
intNoOfAds = 0
intLooper = 0
intLooper2 = 0

' Connect

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb;Persist Security Info=False"
dbConnection.Open strConnectionString

' Run sql command and check for adverts

strSQL = "SELECT * FROM [AdSchedule]"
dbRecordset.Open strSQL, dbConnection

Do While Not dbRecordset.EOF
For Each vntData In dbRecordset.Fields
If DateDiff("s", strNow, vntData) = 0 Then
boolPlaylist = True
strNow = vntData
End If
Next vntData
dbRecordset.MoveNext
Loop

' Break if there is no ad break

If boolAdvert = False Then
Exit Sub
End If

' Get how many adverts are in the break

dbRecordset.Close
strSQL = "SELECT * FROM [" & strNow & "]"
dbRecordset.Open strSQL, dbConnection

Do While Not dbRecordset.EOF
For Each vntData In dbRecordset.Fields
' Do nothing
Next vntData
intNoOfAds = intNoOfAds + 1
dbRecordset.MoveNext
Loop

dbRecordset.Close

' If there are no ads in the break then delete and ignore

If intNoOfAds = 0 Then
strSQL = "DELETE [" & strNow & "]"
dbConnection.Execute strSQL
strSQL = "DELETE FROM [AdSchedule] WHERE [AdBlock] = """ & strNow & """"
dbConnection.Execute strSQL
dbConnection.Close
Exit Sub
End If

' Create the array

ReDim intFileNo(intNoOfAds) As Integer

' Get the data into the array

strSQL = "SELECT [itemid] FROM [" & strNow & "]"
dbRecordset.Open strSQL, dbConnection

Do While Not dbRecordset.EOF
intLooper = intLooper + 1
For Each vntData In dbRecordset.Fields
intFileNo(intLooper) = vntData
Next vntData
dbRecordset.MoveNext
Loop

dbRecordset.Close
intLooper = 0

' Get the remaining data from the adverts table using the previous data

For intLooper = 1 To intNoOfAds
Call doAddToPlaylist(intFileNo(intLooper), "adverts")
Next intLooper

' Delete the loaded ad break

strSQL = "DROP TABLE [" & strNow & "]"
dbConnection.Execute strSQL
strSQL = "DELETE FROM [AdSchedule] WHERE [AdBlock] = """ & strNow & """"
dbConnection.Execute strSQL

' Close the connection

dbConnection.Close

' Move the ad break to the top of the list
' First get how many other items there are

intOtherItems = frmPlayers.lstPlaylist.ListItems.Count - intNoOfAds

' Add these items to an array then delete them from the list

ReDim vntOtherItems(intNoOfAds, 12)

For intLooper = 1 To intOtherItems
vntOtherItems(intLooper, 1) = frmPlayers.lstPlaylist.ListItems.item(intLooper).text
For intLooper2 = 2 To 12
vntOtherItems(intLooper, intLooper2) = frmPlayers.lstPlaylist.ListItems.item(intLooper).SubItems(intLooper2 - 1)
Next intLooper2
Next intLooper

For intLooper = 1 To intOtherItems
frmPlayers.lstPlaylist.ListItems.Remove 1
Next intLooper

' Add them to the end of the playlist

For intLooper = 1 To intOtherItems
With frmPlayers.lstPlaylist.ListItems.Add(, , vntOtherItems(intLooper, 1))
For intLooper2 = 2 To 12
.SubItems(intLooper2 - 1) = vntOtherItems(intLooper, intLooper2)
Next intLooper2
End With
Next intLooper

' Update the names

lstPlaylist.ListItems.Clear
For intLooper = 1 To frmPlayers.lstPlaylist.ListItems.Count
With lstPlaylist.ListItems.Add(, , frmPlayers.lstPlaylist.ListItems.item(intLooper).SubItems(1))
.SubItems(1) = frmPlayers.lstPlaylist.ListItems.item(intLooper).SubItems(2)
End With
Next intLooper

End Sub

Sub check_playlist(ByVal strNow As String)

Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim vntData As Variant
Dim strConnectionString As String
Dim boolPlaylist As Boolean
Dim strSQL As String
Dim intNoOfPlaylists As Integer
Dim intLooper As Integer
Dim intLooper2 As Integer
Dim intItems As Integer
Dim strTemp As String
Dim strID As String
Dim strType As String
Dim intTotalItems As String

' Set initial values

boolPlaylist = False
intNoOfPlaylists = 0
intLooper = 0
intLooper2 = 0

' Connect

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb;Persist Security Info=False"
dbConnection.Open strConnectionString

' Run sql command and check for playlists

strSQL = "SELECT * FROM [PlaylistSchedule]"
dbRecordset.Open strSQL, dbConnection

Do While Not dbRecordset.EOF
For Each vntData In dbRecordset.Fields
If DateDiff("s", strNow, vntData) = 0 Then
boolPlaylist = True
strNow = vntData
End If
Next vntData
dbRecordset.MoveNext
Loop

' Break if there is no playlist scheduled

If boolPlaylist = False Then
Exit Sub
End If

' Get how many playlists are required

dbRecordset.Close
strSQL = "SELECT * FROM [PL" & strNow & "]"
dbRecordset.Open strSQL, dbConnection

Do While Not dbRecordset.EOF
For Each vntData In dbRecordset.Fields
' Do nothing
Next vntData
intNoOfPlaylists = intNoOfPlaylists + 1
dbRecordset.MoveNext
Loop

dbRecordset.Close

' Create the array

ReDim strPlaylist(intNoOfPlaylists) As String

' Get the data into the array

strSQL = "SELECT [playlistblock] FROM [PL" & strNow & "]"
dbRecordset.Open strSQL, dbConnection

Do While Not dbRecordset.EOF
intLooper = intLooper + 1
For Each vntData In dbRecordset.Fields
strPlaylist(intLooper) = vntData
Next vntData
dbRecordset.MoveNext
Loop

dbRecordset.Close
intLooper = 0
intTotalItems = 0

' Get the remaining data from the playlist files using the previous data

For intLooper = 1 To intNoOfPlaylists
Open strPlaylist(intLooper) For Input As #1
Line Input #1, strTemp
intItems = strTemp
intTotalItems = intTotalItems + intItems

For intLooper2 = 1 To intItems
Line Input #1, strID
Line Input #1, strType
Call doAddToPlaylist(strID, strType)
Next intLooper2

Close #1
Next intLooper

' Delete the loaded playlists

strSQL = "DROP TABLE [PL" & strNow & "]"
dbConnection.Execute strSQL
strSQL = "DELETE FROM [PlaylistSchedule] WHERE [PlaylistBlock] = """ & strNow & """"
dbConnection.Execute strSQL

' Close the connection

dbConnection.Close

' Move the playlist to the top of the list
' First get how many other items there are

intOtherItems = frmPlayers.lstPlaylist.ListItems.Count - intTotalItems

' Add these items to an array then delete them from the list

ReDim vntOtherItems(intOtherItems, 12)

For intLooper = 1 To intOtherItems
vntOtherItems(intLooper, 1) = frmPlayers.lstPlaylist.ListItems.item(intLooper).text
For intLooper2 = 2 To 12
vntOtherItems(intLooper, intLooper2) = frmPlayers.lstPlaylist.ListItems.item(intLooper).SubItems(intLooper2 - 1)
Next intLooper2
Next intLooper

For intLooper = 1 To intOtherItems
frmPlayers.lstPlaylist.ListItems.Remove 1
Next intLooper

' Add them to the end of the playlist

For intLooper = 1 To intOtherItems
With frmPlayers.lstPlaylist.ListItems.Add(, , vntOtherItems(intLooper, 1))
For intLooper2 = 2 To 12
.SubItems(intLooper2 - 1) = vntOtherItems(intLooper, intLooper2)
Next intLooper2
End With
Next intLooper

' Update the names

lstPlaylist.ListItems.Clear
For intLooper = 1 To frmPlayers.lstPlaylist.ListItems.Count
With lstPlaylist.ListItems.Add(, , frmPlayers.lstPlaylist.ListItems.item(intLooper).SubItems(1))
.SubItems(1) = frmPlayers.lstPlaylist.ListItems.item(intLooper).SubItems(2)
End With
Next intLooper

End Sub

Private Sub tmrNews_Timer()

Dim strNews As String

' Keep going until the news jingle finishes

strNews = Right$(Format(time(), "HH:mm:ss"), 5)

' If it is finished then play the news track and hold for required time

If strNews = "00:00" Then
tmrNews.Enabled = False
lngChannel = MainPlayers(1)
Call BASS_ChannelSetAttributes(lngChannel, -1, 100, -101)
Call BASS_ChannelPlay(lngChannel, 1)
Call BASS_FX_DSP_Set(lngChannel, BASS_FX_DSPFX_DAMP, 100)
Call BASS_FX_DSP_SetParameters(lngChannel, BASS_FX_DSPFX_DAMP, GlobalAGC)
tmrNews2.Enabled = True
Call ChangeVol(0)
End If

End Sub

Private Sub tmrNews2_Timer()

Dim strNews As String
Dim strFileName As String
Dim strTemp As String
Dim strTime As String
Dim strMins As String
Dim strSecs As String
Dim intMins As Integer
Dim intSecs As Integer

' Open the file to obtain the news length

strFileName = App.Path & "\news.dat"
Open strFileName For Input As #1

Line Input #1, strTemp
Line Input #1, strTemp
Line Input #1, strSecs

Close #1

' Load the values into the integers

intSecs = strSecs
intMins = Int(intSecs / 60)
intSecs = intSecs - (intMins * 60)

' Generate the time string

If intMins >= 10 Then
    strTime = intMins & ":"
Else
    strTime = "0" & intMins & ":"
End If

If intSecs >= 10 Then
    strTime = strTime & intSecs
Else
    strTime = strTime & "0" & intSecs
End If

' Keep going until the news finishes

strNews = Right$(Format(time(), "HH:mm:ss"), 5)

' If the news period has finished then restart the players
' Do not autostart the players if break after is on

If DateDiff("s", strTime, strNews) = 0 Then
    tmrNews2.Enabled = False
    cmdAutomate.Enabled = True
    cmdPlay.Enabled = True
    cmdStop.Enabled = True
    lngChannel = MainPlayers(0)
    Call BASS_StreamFree(lngChannel)
    lngChannel = MainPlayers(1)
    Call ChangeVol(1)
    
    Call BASS_StreamFree(lngChannel)
    boolStarted = False
    If chkBreakAfter.value = 0 Then
        Call cmdPlay_Click
    End If
    If GlobalRotPos > -1 Then
        Call Do_Automate
    End If
End If

End Sub

Private Sub tmrTimeDate_Timer()

Dim strTime As String
Dim strDay As String
Dim strDate As String
Dim intMonth As Integer
Dim strMonth As String
Dim CurrentTime As Variant

CurrentTime = Now()

' Update current time

strTime = Format(time(), "HH:mm:ss")
strDate = Day(Date)
strDay = Weekday(Date, vbMonday)
intMonth = Month(Date)

' Create date string

Select Case strDate
Case "1", "21", "31"
strDate = strDate & "st"
Case "2", "22"
strDate = strDate & "nd"
Case "3", "23"
strDate = strDate & "rd"
Case Else
strDate = strDate & "th"
End Select

' Create month string

Select Case intMonth
Case 1
strMonth = "January"
Case 2
strMonth = "February"
Case 3
strMonth = "March"
Case 4
strMonth = "April"
Case 5
strMonth = "May"
Case 6
strMonth = "June"
Case 7
strMonth = "July"
Case 8
strMonth = "August"
Case 9
strMonth = "September"
Case 10
strMonth = "October"
Case 11
strMonth = "November"
Case 12
strMonth = "December"
End Select

' Create day string

Select Case strDay
Case 1
strDay = "Monday"
Case 2
strDay = "Tuesday"
Case 3
strDay = "Wednesday"
Case 4
strDay = "Thursday"
Case 5
strDay = "Friday"
Case 6
strDay = "Saturday"
Case 7
strDay = "Sunday"
End Select

lblDateTime.Caption = strDay & ", " & strDate & " " & strMonth & " " & Year(Date) & " - " & strTime

' Work out time to news

Call Time_To_News

' Get adverts

If Right$(Format(time(), "HH:mm:ss"), 2) = "00" Then
Call check_ad_break(CurrentTime)
Call check_playlist(CurrentTime)
End If

End Sub

Sub Time_To_News()

Dim strSeconds As String
Dim strMinutes As String
Dim strNow As String
Dim intSecondsToHour As Integer
Dim intSeconds As Integer
Dim intMinutes As Integer
Dim strFileLocation As String
Dim strSecondsToHour As String
Dim intTemp As Integer
Dim strTemp As String
Dim strRunNews As String
Dim lngChannel As Long
Dim intInput As Integer
Dim intLooper As Integer
Dim strJingle As String
Dim intJingle As Integer
Dim intSource As Integer
Dim boolFile As Boolean
Dim strFile As String
Dim intSoundCard As Integer
Dim boolRunNews As Boolean

On Error GoTo BLANKTEMP

' Open the news.dat with the news data

strFileLocation = App.Path + "\news.dat"
Open strFileLocation For Input As #1
Line Input #1, strJingle
Line Input #1, strTemp
intJingle = strTemp
Line Input #1, strTemp
intSource = strTemp
Line Input #1, strTemp
boolRunNews = strTemp
Line Input #1, strTemp
boolFile = strTemp

If boolFile = True Then
    Line Input #1, strFile
Else
    Line Input #1, strTemp
    intSoundCard = strTemp
    Line Input #1, strTemp
    intInput = strTemp
End If
Close #1

' Cancel the news timers if the news has been disabled

If boolRunNews = False Then
lblNews_Time.Visible = False
lblNews.Visible = False
Exit Sub
End If

' Otherwise show the timers

lblNews_Time.Visible = True
lblNews.Visible = True

' Get the time in seconds to the hour

strNow = Format(time(), "HH:mm:ss")
strTemp = Right$(strNow, 2)
intSeconds = strTemp
strNow = Left$(strNow, 5)
strTemp = Right$(strNow, 2)
intMinutes = strTemp
intSeconds = intSeconds + (intMinutes * 60)
intSecondsToHour = intSeconds

' Do the check
' If it is time, lock all of the players and play the news jingles

intSeconds = 3600 - intJingle - intSecondsToHour
If intSeconds = 0 And tmrNews.Enabled = False Then

' Log previous track

If lblPlayed_Time.Caption <> "00:00" Then
Call Do_Add_2hr
End If

' Stop All Players

On Error Resume Next

Call BASS_SetDevice(GetSoundCard(MAIN_SOUND_CARD))
For intLooper = 0 To 3
lngChannel = MainPlayers(intLooper)
'Call BASS_StreamFree(lngChannel)
Call BASS_ChannelSlideAttributes(lngChannel, -1, 0, -101, 2000)
Next intLooper

For intLooper = 0 To 8
lngChannel = frmInstantPlayers.txtInstantPlayer(intLooper).text
Call BASS_ChannelStop(lngChannel)
Next intLooper

' Disable buttons

cmdAutomate.Enabled = False
cmdPlay.Enabled = False
cmdStop.Enabled = False
tmrNews.Enabled = True
tmrGetData.Enabled = False

' Setup display

lblPlayed_Time.Caption = "00:00"
lblRemaining_Time.Caption = "00:00"
lblNews_Time.Caption = "00:00"
txtArtist.text = "News"
txtSong.text = "News"
proSongTime.value = 0
lblRemaining_Time.Caption = dblIntro
lblRemaining_Time.BackColor = vbBlack
lblRemaining_Time.ForeColor = vbRed
lblTimeRemaining.Caption = "Time Remaining"
lblRemaining_Time.Caption = "00:00"

' Play news files

lngChannel = BASS_StreamCreateFile(BASSFALSE, strJingle, 0, 0, 0)
MainPlayers(0) = lngChannel
Call BASS_ChannelSetAttributes(lngChannel, -1, 100, -101)
Call BASS_FX_DSP_Set(lngChannel, BASS_FX_DSPFX_DAMP, 100)
Call BASS_FX_DSP_SetParameters(lngChannel, BASS_FX_DSPFX_DAMP, GlobalAGC)
Call BASS_ChannelPlay(lngChannel, BASSFALSE)

If boolFile = True Then
    lngChannel = BASS_StreamCreateFile(BASSFALSE, strFile, 0, 0, 0)
    MainPlayers(1) = lngChannel
    boolUseFeed = False
Else
    gblFeed = Trim$(UCase$(sgMain.OutputLineName(sgMain.OutputLineID(intInput))))
    MainPlayers(1) = 0
    boolUseFeed = True
End If
End If

' Now get the current time to news

If intSeconds < 0 Then
intSeconds = 0
End If

intMinutes = Int(intSeconds / 60)
intSeconds = Int(intSeconds Mod 60)

If intMinutes < 10 Then
strMinutes = "0" & intMinutes
End If

If intMinutes >= 10 Then
strMinutes = intMinutes
End If

If intSeconds < 10 Then
strSeconds = "0" & intSeconds
End If

If intSeconds >= 10 Then
strSeconds = intSeconds
End If

lblNews_Time.Caption = strMinutes & ":" & strSeconds
Exit Sub

BLANKTEMP:

strTemp = ""
Resume Next

End Sub

Sub Update_Timer()

Dim dblTotal As Double
Dim intItems As Integer
Dim intLooper As Integer
Dim intSeconds As Integer
Dim intMinutes As Integer
Dim intHours As Integer
Dim strSeconds As String
Dim strMinutes As String
Dim strHours As String
Dim strListTime As String
Dim intListTime As Integer

' Get the number of items

intItems = frmPlayers.lstPlaylist.ListItems.Count

' Cycle through the items adding the totals

dblTotal = 0
For intLooper = 1 To intItems
strListTime = frmPlayers.lstPlaylist.ListItems.item(intLooper).SubItems(3)
intListTime = strListTime
dblTotal = dblTotal + intListTime
Next intLooper

' Add the reamaining time on the players

If lblPlayed_Time.Caption <> "00:00" Then
dblTotal = dblTotal + (proSongTime.max - proSongTime.value)
End If

' Change the times to integers

intHours = Int(dblTotal / 3600)
intMinutes = Int((dblTotal / 60) - (intHours * 60))
intSeconds = Int(dblTotal Mod 3600)

If intSeconds > 60 Then
Do While intSeconds > 60
intSeconds = intSeconds - 60
Loop
End If

' Change times to strings and display


If intMinutes < 10 Then
strMinutes = "0" & intMinutes
End If

If intMinutes >= 10 Then
strMinutes = intMinutes
End If

If intSeconds < 10 Then
strSeconds = "0" & intSeconds
End If

If intSeconds >= 10 Then
strSeconds = intSeconds
End If

If intHours < 10 Then
strHours = "0" & intHours
End If

If intHours >= 10 Then
strHours = intHours
End If

lblPlaylist_Time.Caption = strHours & ":" & strMinutes & ":" & strSeconds

End Sub

Sub Do_Automate()

Dim intRotPos As Integer
Dim boolTimeRestriction As Boolean
Dim strFileLocation As String
Dim strRotation As String
Dim strRotLength As String
Dim intRotLength As Integer
Dim intLooper As Integer
Dim dbConnection As New ADODB.Connection
Dim dbRecordset As New ADODB.Recordset
Dim strConnectionString As String
Dim intItemsAvailable As Integer
Dim dblTimeLeft As Double
Dim intSelected As Integer
Dim vntData(3) As Variant
Dim intLooper2 As Integer
Dim intOccurences As Integer
Dim strTemp As String
Dim boolNewsVisible As Boolean

' Starting position for automation iteration
' Used to prevent memory overload by opening too many copies of the same subroutine

Start_Of_Automation:

' DoEvents prevents locking out other events if a perpertual loop does occur

DoEvents

' Quit now if there is already something on the playlist
' Updated to allow 2 items on playlist for effective voice tracking

If lstPlaylist.ListItems.Count > 2 Then
Exit Sub
End If

' Get the data from the text file

strFileLocation = App.Path & "\rotation.dat"
Open strFileLocation For Input As #1
Line Input #1, strRotation
Line Input #1, strRotLength
Close #1

' Convert the data

intRotLength = strRotLength

' Check if the news is running

boolNewsVisible = lblNews.Visible

' Update the rotation times variable

intRotTimes = RotTimes
intRotTimes = intRotTimes + 1
RotTimes = intRotTimes

' Quit as appropriate
' If the only reason repetition occured is because the news is too close then ignore the time to news

If intRotTimes > intRotLength Then
RotTimes = 0
If boolNewsVisible = False Then
Exit Sub
Else
boolNewsVisible = False
End If
End If

' Create the array

ReDim strRotationArray(intRotLength - 1) As String
ReDim strRotationArrayFull(intRotLength) As String

' Get the rotation position

intRotPos = GlobalRotPos
intRotPos = intRotPos + 1
If intRotPos > intRotLength Then
intRotPos = 1
End If
GlobalRotPos = intRotPos

' Split the rotation file

strRotationArray() = Split(strRotation, ";")

' Convert each letter to an actual item

For intLooper = 1 To intRotLength
Select Case strRotationArray(intLooper - 1)
Case "s"
strRotationArrayFull(intLooper) = "songs"
Case "j"
strRotationArrayFull(intLooper) = "jingles"
Case "x"
strRotationArrayFull(intLooper) = "specialized"
Case "o"
strRotationArrayFull(intLooper) = "seasonal"
Case "a"
strRotationArrayFull(intLooper) = "adverts"
Case "v"
strRotationArrayFull(intLooper) = "voice_tracks"
End Select
Next intLooper

' Now connect to the database

strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\record_collection.mdb;Persist Security Info=False"
dbConnection.Open strConnectionString

' Get the time to the news

' Open the file

strFileLocation = App.Path + "\news.dat"
Open strFileLocation For Input As #1
Line Input #1, strTemp
Line Input #1, strSecondsToHour
Close #1

' Get the actual seconds to the hour

strTemp = Format(time(), "HH:mm:ss")
dblTimeLeft = Right$(strTemp, 2)
strTemp = Left$(strTemp, 5)
dblTimeLeft = dblTimeLeft + (Right$(strTemp, 2) * 60)

' Take away the length of the news jingle

dblTimeLeft = dblTimeLeft - strSecondsToHour

' Now we have the current time to news we can adjust for other items

If frmPlayers.lstPlaylist.ListItems.Count > 0 Then
For intLooper = 1 To frmPlayers.lstPlaylist.ListItems.Count
dblTimeLeft = dbltimerleft - frmPlayers.lstPlaylist.ListItems.item(intLooper).SubItems(3)
Next intLooper
End If

' Also account for the item currently playing

If boolStarted = True Then
dblTimeLeft = dblTimeLeft - (proSongTime.max - BASS_ChannelBytes2Seconds(GlobalPlayer, BASS_ChannelGetPosition(GlobalPlayer)))
End If

' Create the sql command
' A restricted search will be placed if we are waiting for the news, else an unrestricted search will occur

If Not boolNewsVisible Or dblTimeLeft <= 0 Then
strSQL = "SELECT [itemid] FROM " & strRotationArrayFull(intRotPos)
End If

If boolNewsVisible Then
strSQL = "SELECT [itemid] FROM " & strRotationArrayFull(intRotPos) & " WHERE length <= " & (dblTimeLeft + 1)
End If

' Get the items from the database

intItemsAvailable = 0
dbRecordset.Open strSQL, dbConnection
Do While Not dbRecordset.EOF
intItemsAvailable = intItemsAvailable + 1
dbRecordset.MoveNext
Loop

dbRecordset.Close

' If there is nothing then recall this sub and exit

If intItemsAvailable = 0 Then
dbConnection.Close
Erase strRotationArray
Erase strRotationArrayFull
GoTo Start_Of_Automation
End If

' If there is items, create the arrays

ReDim intID(intItemsAvailable) As Integer
ReDim strArtist(intItemsAvailable) As String
ReDim strTrack(intItemsAvailable) As String

' Select the lucky item

intSelected = Int(Rnd * intItemsAvailable) + 1

' Get the data

If Not boolNewsVisible Or dblTimeLeft <= 0 Then
strSQL = "SELECT [itemid], [artist], [track] FROM " & strRotationArrayFull(intRotPos)
End If
If boolNewsVisible Then
strSQL = "SELECT [itemid], [artist], [track] FROM " & strRotationArrayFull(intRotPos) & " WHERE length <= " & (dblTimeLeft + 1)
End If

intLooper = 0
dbRecordset.Open strSQL, dbConnection
Do While Not dbRecordset.EOF
intLooper = intLooper + 1
intLooper2 = 0
For Each x In dbRecordset.Fields
intLooper2 = intLooper2 + 1
vntData(intLooper2) = x
Next x
intID(intLooper) = vntData(1)
strArtist(intLooper) = vntData(2)
strTrack(intLooper) = vntData(3)
dbRecordset.MoveNext
Loop

' If there is nothing returned as a result then loop round using
' the next part of the rotation

If intLooper = 0 Then
    GoTo Start_Of_Automation
End If

' Check the item is not in the two hour log

intLooper = 0

Do_2hr_check:

dbRecordset.Close
dbConnection.Close
strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath() & "\logs.mdb;Persist Security Info=False"
dbConnection.Open strConnectionString

strSQL = "SELECT * FROM 2hr WHERE [artist] = """ & strArtist(intSelected) & """"
dbRecordset.Open strSQL, dbConnection

intLooper = intLooper + 1

intOccurences = 0
Do While Not dbRecordset.EOF
intoccureneces = intOccurences + 1
dbRecordset.MoveNext
Loop

' Select another item if played more than 3 times in 2 hours (Digital Millenium Copyright Act)

If intOccurences >= 3 Then
intSelected = Int(Rnd * intItemsAvailable) + 1
GoTo Do_2hr_check
End If

' Exit if all values have been tried

If intLooper = intItemsAvailable And intOccurences >= 3 Then
Erase strRotationArray
Erase strRotationArrayFull
Erase intID
Erase strArtist
Erase strTrack
GoTo Start_Of_Automation
End If

' Reset the crash prevention variable

RotTimes = 0

' Add the selected item to the playlist

Call doAddToPlaylist(intID(intSelected), strRotationArrayFull(intRotPos))

' Close the connection

dbRecordset.Close
dbConnection.Close

End Sub

Private Sub tmrVolume_Timer()

    Dim intLooper As Integer
    
    ' Fade up or down as appropriate
    
    If BASS_ChannelIsActive(tmrVolume.tag) = BASSTRUE Then
        For intLooper = 0 To 3
            Call BASS_ChannelSetAttributes(MainPlayers(intLooper), -1, FadeLevel, -101)
        Next intLooper
    Else
        For intLooper = 0 To 3
            Call BASS_ChannelSlideAttributes(MainPlayers(intLooper), -1, 100, -101, 500)
        Next intLooper
        tmrVolume.Enabled = False
    End If

End Sub

Private Sub txtArtist_GotFocus()
lstPlaylist.SetFocus
End Sub

Private Sub txtSong_GotFocus()
lstPlaylist.SetFocus
End Sub

Public Sub LVDragDropSingle(ByRef lvList As ListView, ByVal x As Single, ByVal y As Single)

Dim VisibleSelectedItem(0 To 1) As Variant
Dim HiddenSelectedItem(0 To 12) As Variant
Dim intSelected As Integer
Dim intNewPos As Integer
Dim lstItem As ListItem

On Error GoTo Exiter

' Get the selected item and it's new position

intSelected = lstPlaylist.SelectedItem.Index
lstPlaylist.SelectedItem = lstPlaylist.HitTest(x, y)
intNewPos = lstPlaylist.SelectedItem.Index

' Leave if there is no change to be made

If intSelected = intNewPos Then
    Exit Sub
End If

' Load the variable from the selected item

VisibleSelectedItem(0) = lstPlaylist.ListItems.item(intSelected).text
VisibleSelectedItem(1) = lstPlaylist.ListItems.item(intSelected).SubItems(1)

HiddenSelectedItem(0) = frmPlayers.lstPlaylist.SelectedItem.text

For intLooper = 1 To 12
HiddenSelectedItem(intLooper) = frmPlayers.lstPlaylist.ListItems.item(intSelected).SubItems(intLooper)
Next intLooper

' Remove the item from the playlist

lstPlaylist.ListItems.Remove (intSelected)
frmPlayers.lstPlaylist.ListItems.Remove (intSelected)

' Place the item in the new position

With lstPlaylist.ListItems.Add(intNewPos, , VisibleSelectedItem(0))
.SubItems(1) = VisibleSelectedItem(1)
End With

With frmPlayers.lstPlaylist.ListItems.Add(intNewPos, , HiddenSelectedItem(0))
For intLooper = 1 To 12
.SubItems(intLooper) = HiddenSelectedItem(intLooper)
Next intLooper
End With

' Preload the item

Call Pre_Load

Exiter:

End Sub

Private Sub txtSongInfo_GotFocus()

    AllowMainKeyboard = False

End Sub

Private Sub txtSongInfo_LostFocus()

    AllowMainKeyboard = True

End Sub
