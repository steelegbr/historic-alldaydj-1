' BASS_WADSP 2.2.0.5 Visual Basic API, (c) 2005-2006 TEN53 GbR, Bernd Niedergesaess.
' Requires BASS - available @ www.un4seen.com
' See the BASS_WADSP.txt file for detailed documentation

Attribute VB_Name = "BASS_WADSP"

' Winamp SDK message parameter values (for lParam)
Global Const BASS_WADSP_IPC_GETOUTPUTTIME    = 105
Global Const BASS_WADSP_IPC_ISPLAYING        = 104
Global Const BASS_WADSP_IPC_GETVERSION       = 0
Global Const BASS_WADSP_IPC_STARTPLAY        = 102
Global Const BASS_WADSP_IPC_GETINFO          = 126
Global Const BASS_WADSP_IPC_GETLISTLENGTH    = 124
Global Const BASS_WADSP_IPC_GETLISTPOS       = 125
Global Const BASS_WADSP_IPC_GETPLAYLISTFILE  = 211
Global Const BASS_WADSP_IPC_GETPLAYLISTTITLE = 212
Global Const BASS_WADSP_IPC                  = 1024


Declare Function BASS_WADSP_Init Lib "bass_wadsp.dll" (ByVal hwndMain As Long) As Long
Declare Function BASS_WADSP_Free Lib "bass_wadsp.dll" () As Long
Declare Sub BASS_WADSP_FreeDSP Lib "bass_wadsp.dll" (ByVal plugin As Long)
Declare Function BASS_WADSP_GetFakeWinampWnd Lib "bass_wadsp.dll" (ByVal plugin As Long) As Long
Declare Sub BASS_WADSP_SetSongTitle Lib "bass_wadsp.dll" (ByVal plugin As Long, ByVal thetitle As Long)    ' thetitle is actually a string
Declare Sub BASS_WADSP_SetFileName Lib "bass_wadsp.dll" (ByVal plugin As Long, ByVal thefile As Long)      ' thefile is actually a string

Declare Function BASS_WADSP_Load Lib "bass_wadsp.dll" (ByVal dspfile As String, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal proc As Long) As Long
Declare Sub BASS_WADSP_Config Lib "bass_wadsp.dll" (ByVal plugin As Long)
Declare Sub BASS_WADSP_Start Lib "bass_wadsp.dll" (ByVal plugin As Long, ByVal module As Single, ByVal hchan As Long)
Declare Sub BASS_WADSP_Stop Lib "bass_wadsp.dll" (ByVal plugin As Long)
Declare Sub BASS_WADSP_SetChannel Lib "bass_wadsp.dll" (ByVal plugin As Long, ByVal hchan As Long)
Declare Function BASS_WADSP_GetModule Lib "bass_wadsp.dll" (ByVal plugin As Long) As Single
Declare Function BASS_WADSP_ChannelSetDSP Lib "bass_wadsp.dll" (ByVal plugin As Long, ByVal hchan As Long, ByVal priority As Single) As Long
Declare Function BASS_WADSP_ChannelRemoveDSP Lib "bass_wadsp.dll" (ByVal plugin As Long) As Long

Declare Function BASS_WADSP_ModifySamplesSTREAM Lib "bass_wadsp.dll" (ByVal plugin As Long, ByRef buffer As Any, ByVal length As Long) As Long
Declare Function BASS_WADSP_ModifySamplesDSP Lib "bass_wadsp.dll" (ByVal plugin As Long, ByRef buffer As Any, ByVal length As Long) As Long

Declare Function BASS_WADSP_GetName Lib "bass_wadsp.dll" (ByVal plugin As Long) As Long    ' returns actually a string
Declare Function BASS_WADSP_GetModuleCount Lib "bass_wadsp.dll" (ByVal plugin As Long) As Single
Declare Function BASS_WADSP_GetModuleName Lib "bass_wadsp.dll" (ByVal plugin As Long, ByVal module As Single) As Long    ' returns actually a string

Declare Sub BASS_WADSP_PluginInfoFree Lib "bass_wadsp.dll" ()
Declare Function BASS_WADSP_PluginInfoLoad Lib "bass_wadsp.dll" (ByVal dspfile As String) As Long
Declare Function BASS_WADSP_PluginInfoGetName Lib "bass_wadsp.dll" () As Long    ' returns actually a string
Declare Function BASS_WADSP_PluginInfoGetModuleCount Lib "bass_wadsp.dll" () As Single
Declare Function BASS_WADSP_PluginInfoGetModuleName Lib "bass_wadsp.dll" (ByVal module As Single) As Long    ' returns actually a string


Sub WINAMPWINPROC(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long)
    
    'CALLBACK FUNCTION !!! (but do not use this one directly - this is simple the declaration!)

    ' User defined Window Message Process Handler
    ' hwnd   : The Window handle we are dealing with
    ' msg    : The window message send
    ' wParam : The wParam message parameter see the Winamp SDK for further details
    ' lParam : The lParam message parameter see the Winamp SDK for further details
    
End Sub
