VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' +++++++++++++ Class module for Volume Control with Visual Basic ++++++++++++++
' ++++++++ Klassenmodul für die Lautstärkeeinstellung mit Visual Basic +++++++++
' ++++ Components: Main Volume, Microphone, WaveIn, LineIn, CD-Audio, MIDI +++++
'  Die Komponenten: Gesamtlautstärke, Mikrofon, WaveIn, LineIn, CD-Audio, MIDI
' + Ursprünglicher Autor: Unbekannt. Modifiziert und erweitert von J. Thümmler +
'*******************************************************************************
'         03/29/02: Added the Mute functionallity to the class module
'         29.03.02: Mute-Funktionen für die Volume-Arten ergänzt
'*******************************************************************************
Option Explicit
'*******************************************************************************
Private Const MMSYSERR_NOERROR& = 0
Private Const MAXPNAMELEN& = 32
Private Const MIXER_LONG_NAME_CHARS& = 64
Private Const MIXER_SHORT_NAME_CHARS& = 16
Private Const MIXER_GETLINEINFOF_LINEID& = &H2
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE& = &H3&
Private Const MIXER_GETCONTROLDETAILSF_VALUE& = &H0&
Private Const MIXER_GETCONTROLDETAILSF_LISTTEXT& = &H1&
Private Const MIXER_GETLINECONTROLSF_ONEBYID& = &H1
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE& = &H2&
Private Const MIXER_OBJECTF_WAVEOUT& = &H10000000
Private Const MIXER_SETCONTROLDETAILSF_VALUE& = &H0&
'--------------------------------------------------------------------------------
Private Const MIXERCONTROL_CT_CLASS_FADER& = &H50000000
Private Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Private Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED& = &H30000
Private Const MIXERCONTROL_CONTROLTYPE_FADER& = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME& = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Private Const MIXERCONTROL_CONTROLTYPE_BASS& = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Private Const MIXERCONTROL_CONTROLTYPE_TREBLE& = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Private Const MIXERCONTROL_CONTROLTYPE_EQUALIZER& = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
'--------------------------------------------------------------------------------
Private Const MIXERLINE_COMPONENTTYPE_SRC_FIRST& = &H1000&
Private Const MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 0)
Private Const MIXERLINE_COMPONENTTYPE_SRC_DIGITAL& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)
Private Const MIXERLINE_COMPONENTTYPE_SRC_LINE& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Private Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Private Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Private Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Private Const MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)
Private Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
Private Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Private Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)
Private Const MIXERLINE_COMPONENTTYPE_SRC_ANALOG& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)
Private Const MIXERLINE_COMPONENTTYPE_SRC_LAST& = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)
'--------------------------------------------------------------------------------
Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST& = &H0&
Private Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 0)
Private Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 1)
Private Const MIXERLINE_COMPONENTTYPE_DST_LINE& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 2)
Private Const MIXERLINE_COMPONENTTYPE_DST_MONITOR& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 3)
Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Private Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 5)
Private Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 6)
Private Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Private Const MIXERLINE_COMPONENTTYPE_DST_VOICEIN& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Private Const MIXERLINE_COMPONENTTYPE_DST_LAST& = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
'********************************************************************************
Private Type MIXERCAPS
  wMid As Integer                   '#  manufacturer id
  wPid As Integer                   '#  product id
  vDriverVersion As Long            '#  version of the driver
  szPname As String * MAXPNAMELEN   '#  product name
  fdwSupport As Long                '#  misc. support bits
  cDestinations As Long             '#  count of destinations
End Type
'--------------------------------------------------------------------------------
Private Type MIXERCONTROL
  cbStruct As Long                  '#  size in Byte of MIXERCONTROL
  dwControlID As Long               '#  unique control id for mixer device
  dwControlType As Long             '#  MIXERCONTROL_CONTROLTYPE_xxx
  fdwControl As Long                '#  MIXERCONTROL_CONTROLF_xxx
  cMultipleItems As Long            '#  if MIXERCONTROL_CONTROLF_MULTIPLE set
  szShortName As String * MIXER_SHORT_NAME_CHARS  '# short name of control
  szName As String * MIXER_LONG_NAME_CHARS        '# long name of control
  lMinimum As Long                  '#  Minimum value
  lMaximum As Long                  '#  Maximum value
  reserved(10) As Long              '#  reserved structure space
End Type
'--------------------------------------------------------------------------------
Private Type MIXERCONTROLDETAILS
  cbStruct As Long                  '#  size in Byte of MIXERCONTROLDETAILS
  dwControlID As Long               '#  control id to get/set details on
  cChannels As Long                 '#  number of channels in paDetails array
  item As Long                      '#  hwndOwner or cMultipleItems
  cbDetails As Long                 '#  size of _one_ details_XX struct
  paDetails As Long                 '#  pointer to array of details_XX structs
End Type
'--------------------------------------------------------------------------------
Private Type MIXERCONTROLDETAILS_UNSIGNED
  dwValue As Long                   '#  value of the control
End Type
'--------------------------------------------------------------------------------
Private Type MIXERLINE
  cbStruct As Long                  '#  size of MIXERLINE structure
  dwDestination As Long             '#  zero based destination index
  dwSource As Long                  '#  zero based source index (if source)
  dwLineID As Long                  '#  unique line id for mixer device
  fdwLine As Long                   '#  state/information about line
  dwUser As Long                    '#  driver specific information
  dwComponentType As Long           '#  component Private Type line connects to
  cChannels As Long                 '#  number of channels line supports
  cConnections As Long              '#  number of connections (possible)
  cControls As Long                 '#  number of controls at this line
  szShortName As String * MIXER_SHORT_NAME_CHARS
  szName As String * MIXER_LONG_NAME_CHARS
  dwType As Long
  dwDeviceID As Long
  wMid  As Integer
  wPid As Integer
  vDriverVersion As Long
  szPname As String * MAXPNAMELEN
End Type
'--------------------------------------------------------------------------------
Private Type MIXERLINECONTROLS
  cbStruct As Long                  '#  size in Byte of MIXERLINECONTROLS
  dwLineID As Long                  '#  line id (from MIXERLINE.dwLineID)
                                    '#  MIXER_GETLINECONTROLSF_ONEBYID or
  dwControl As Long                 '#  MIXER_GETLINECONTROLSF_ONEBYTYPE
  cControls As Long                 '#  count of controls pamxctrl points to
  cbmxctrl As Long                  '#  size in Byte of _one_ MIXERCONTROL
  pamxctrl As Long                  '#  pointer to first MIXERCONTROL array
End Type
'********************************************************************************
Private Declare Function mixerClose& Lib "winmm.dll" (ByVal hmx&)
Private Declare Function mixerGetControlDetails& Lib "winmm.dll" Alias "mixerGetControlDetailsA" _
               (ByVal hmxobj&, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails&)
Private Declare Function mixerGetDevCaps& Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId&, _
                ByVal pmxcaps As MIXERCAPS, ByVal cbmxcaps&)
Private Declare Function mixerGetID& Lib "winmm.dll" (ByVal hmxobj&, pumxID&, ByVal fdwId&)
Private Declare Function mixerGetLineControls& Lib "winmm.dll" Alias "mixerGetLineControlsA" _
               (ByVal hmxobj&, pmxlc As MIXERLINECONTROLS, ByVal fdwControls&)
Private Declare Function mixerGetLineInfo& Lib "winmm.dll" Alias "mixerGetLineInfoA" _
               (ByVal hmxobj&, pmxl As MIXERLINE, ByVal fdwInfo&)
Private Declare Function mixerGetNumDevs& Lib "winmm.dll" ()
Private Declare Function mixerMessage& Lib "winmm.dll" (ByVal hmx&, ByVal uMsg&, _
                ByVal dwParam1&, ByVal dwParam2&)
Private Declare Function mixerOpen& Lib "winmm.dll" (phmx&, ByVal uMxId&, _
                ByVal dwCallback&, ByVal dwInstance&, ByVal fdwOpen&)
Private Declare Function mixerSetControlDetails& Lib "winmm.dll" (ByVal hmxobj&, _
                pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails&)
Private Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" _
               (Struct As Any, ByVal ptr&, ByVal cb&)
Private Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" _
               (ByVal ptr&, Struct As Any, ByVal cb&)
Private Declare Function GlobalAlloc& Lib "kernel32" (ByVal wFlags&, ByVal dwBytes&)
Private Declare Function GlobalLock& Lib "kernel32" (ByVal hmem&)
Private Declare Function GlobalFree& Lib "kernel32" (ByVal hmem&)
'********************************************************************************
Dim hMixer&                           '# mixer handle
Dim volCtrl As MIXERCONTROL           '# waveout volume control (main volume)
Dim micCtrl As MIXERCONTROL           '# microphone volume control
Dim wavCtrl As MIXERCONTROL           '# wave-in volume control
Dim linCtrl As MIXERCONTROL           '# line-in volume control
Dim cd_Ctrl As MIXERCONTROL           '# audio-cd volume control
Dim midCtrl As MIXERCONTROL           '# MIDI volume control
Dim auxCtrl As MIXERCONTROL           '# aux volume control

Dim volMute As MIXERCONTROL           '# waveout mute control (main volume)
Dim micMute As MIXERCONTROL           '# microphone mute control
Dim wavMute As MIXERCONTROL           '# wave-in mute control
Dim linMute As MIXERCONTROL           '# line-in mute control
Dim cdrMute As MIXERCONTROL           '# audio-cd mute control
Dim midMute As MIXERCONTROL           '# MIDI mute control
Dim auxMute As MIXERCONTROL           '# aux mute control

Dim MinVolVol&, MaxVolVol&            '# Min/Max values for VolControl
Dim MinMicVol&, MaxMicVol&            '# Min/Max values for MicroControl
Dim MinWavVol&, MaxWavVol&            '# Min/Max values for WaveControl
Dim MinLinVol&, MaxLinVol&            '# Min/Max values for LineControl
Dim MinCD_Vol&, MaxCD_Vol&            '# Min/Max values for CD Control
Dim MinMidVol&, MaxMidVol&            '# Min/Max values for MIDI
Dim MinAuxVol&, MaxAuxVol&            '# Min/Max values for Aux

Dim mxl As MIXERLINE
Dim mxc As MIXERCONTROL
Dim mxcd As MIXERCONTROLDETAILS
Dim mxlc As MIXERLINECONTROLS
Dim Struct As MIXERCONTROLDETAILS_UNSIGNED

Dim bOK As Boolean                    '# boolean return code
Dim hmem&                             '# Memory handle
Dim rc&                               '# return code
Dim CtrlState&  '# 1=Main; 2=Micro; 4=Wave; 8=Line; 16=CD; 32=MIDI
Dim MuteState&  '# 1=Main; 2=Micro; 4=Wave; 8=Line; 16=CD; 32=MIDI
'********************************************************************************
'********************************************************************************
'# Hier folgt die "Schnittstelle" des Klassenmoduls nach außen
'# Here the class module "interface" follows
'********************************************************************************
'# Limits of the values
'--------------------------------------------------------------------------------
Public Property Get MinWaveVolume&()          '# Minimalwert der Gesamtlautstärke
  MinWaveVolume = MinVolVol                   '# Minium value of main volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MaxWaveVolume&()          '# Maximalwert der Gesamtlautstärke
  MaxWaveVolume = MaxVolVol                   '# Maximum value of main volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MinMicVolume&()           '# Minimalwert der Mikrolautstärke
  MinMicVolume = MinMicVol                    '# Minium value of micro volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MaxMicVolume&()           '# Maximalwert der Mikrolautstärke
  MaxMicVolume = MaxMicVol                    '# Maximum value of micro volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MinWavInVolume&()         '# Minimalwert der WaveInlautstärke
  MinWavInVolume = MinWavVol                  '# Minium value of wave in volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MaxWavInVolume&()         '# Maximalwert der WaveInlautstärke
  MaxWavInVolume = MaxWavVol                  '# Maximum value of wave in volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MinLineInVolume&()        '# Minimalwert der LineInlautstärke
  MinLineInVolume = MinLinVol                 '# Minium value of line in volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MaxLineInVolume&()        '# Maximalwert der LineInlautstärke
  MaxLineInVolume = MaxLinVol                 '# Maximum value of line in volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MinCDVolume&()            '# Minimalwert der CD-Lautstärke
  MinCDVolume = MinCD_Vol                     '# Minium value of CD volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MaxCDVolume&()            '# Maximalwert der CD-Lautstärke
  MaxCDVolume = MaxCD_Vol                     '# Maximum value of CD volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MinMidVolume&()           '# Minimalwert der MIDI-Lautstärke
  MinMidVolume = MinMidVol                    '# Minium value of MIDI volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MaxMidVolume&()           '# Maximalwert der MIDI-Lautstärke
  MaxMidVolume = MaxMidVol                    '# Maximum value of MIDI volume
End Property
'--------------------------------------------------------------------------------
'# Current values
'--------------------------------------------------------------------------------
Public Property Get WaveVolume&()             '# Aktuelle Gesamtlautstärke
  WaveVolume = GetValue(volCtrl)              '# Current value of main volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MicroVolume&()            '# Aktuelle Mikrofonlautstärke
  MicroVolume = GetValue(micCtrl)             '# Current value of micro volume
End Property
'--------------------------------------------------------------------------------
Public Property Get WaveInVolume&()           '# Aktuelle WaveIn-Lautstärke
  WaveInVolume = GetValue(wavCtrl)            '# Current value of wave in volume
End Property
'--------------------------------------------------------------------------------
Public Property Get LineInVolume&()           '# Aktuelle LineIn-Lautstärke
  LineInVolume = GetValue(linCtrl)            '# Current value of line in volume
End Property
'--------------------------------------------------------------------------------
Public Property Get CD_Volume&()              '# Aktuelle CD-Lautstärke
  CD_Volume = GetValue(cd_Ctrl)               '# Current value of CD volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MIDIVolume&()             '# Aktuelle MIDI-Lautstärke
  MIDIVolume = GetValue(midCtrl)              '# Current value of MIDI volume
End Property
'--------------------------------------------------------------------------------
'# Mute Values
'--------------------------------------------------------------------------------
Public Property Get WaveMute&()               '# Mute-Status der Gesamtlautstärke
  WaveMute = GetValue(volMute)                '# Mute State of main volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MicroMute&()              '# Mute-Status der Mikrofonlautstärke
  MicroMute = GetValue(micMute)               '# Mute State of micro volume
End Property
'--------------------------------------------------------------------------------
Public Property Get WaveInMute&()             '# Mute-Status der WaveIn-Lautstärke
  WaveInMute = GetValue(wavMute)              '# Mute State of wave in volume
End Property
'--------------------------------------------------------------------------------
Public Property Get LineInMute&()             '# Mute-Status der LineIn-Lautstärke
  LineInMute = GetValue(linMute)              '# Mute State of line in volume
End Property
'--------------------------------------------------------------------------------
Public Property Get CD_Mute&()                '# Mute-Status der CD-Lautstärke
  CD_Mute = GetValue(cdrMute)                 '# Mute State of CD volume
End Property
'--------------------------------------------------------------------------------
Public Property Get MIDIMute&()               '# Mute-Status der MIDI-Lautstärke
  MIDIMute = GetValue(midMute)                '# Mute State of MIDI volume
End Property
'********************************************************************************
'# Volume Settings
'--------------------------------------------------------------------------------
Public Property Let WaveVolume(ByVal NewVolume&)   '# Einstellung Gesamtlautstärke
  If NewVolume < MinVolVol Then NewVolume = MinVolVol       '# Set main volume
  If NewVolume > MaxVolVol Then NewVolume = MaxVolVol
  Call SetValue(volCtrl, NewVolume)
End Property
'--------------------------------------------------------------------------------
Public Property Let MicroVolume(ByVal NewVolume&)  '# Einstellung Mikro-Lautstärke
  If NewVolume < MinMicVol Then NewVolume = MinMicVol       '# Set micro volume
  If NewVolume > MaxMicVol Then NewVolume = MaxMicVol
  Call SetValue(micCtrl, NewVolume)
End Property
'--------------------------------------------------------------------------------
Public Property Let WaveInVolume(ByVal NewVolume&) '# Einstellung Wave-Lautstärke
  If NewVolume < MinWavVol Then NewVolume = MinWavVol       '# Set wave in volume
  If NewVolume > MaxWavVol Then NewVolume = MaxWavVol
  Call SetValue(wavCtrl, NewVolume)
End Property
'--------------------------------------------------------------------------------
Public Property Let LineInVolume(ByVal NewVolume&) '# Einstellung Line-Lautstärke
  If NewVolume < MinLinVol Then NewVolume = MinLinVol       '# Set line in volume
  If NewVolume > MaxLinVol Then NewVolume = MaxLinVol
  Call SetValue(linCtrl, NewVolume)
End Property
'--------------------------------------------------------------------------------
Public Property Let CD_Volume(ByVal NewVolume&)    '# Einstellung CD-Lautstärke
  If NewVolume < MinCD_Vol Then NewVolume = MinCD_Vol       '# Set CD volume
  If NewVolume > MaxCD_Vol Then NewVolume = MaxCD_Vol
  Call SetValue(cd_Ctrl, NewVolume)
End Property
'--------------------------------------------------------------------------------
Public Property Let MIDIVolume(ByVal NewVolume&)   '# Einstellung MIDI Lautstärke
  If NewVolume < MinMidVol Then NewVolume = MinMidVol       '# Set MIDI volume
  If NewVolume > MaxMidVol Then NewVolume = MaxMidVol
  Call SetValue(midCtrl, NewVolume)
End Property
'--------------------------------------------------------------------------------
'# Mute Settings
'--------------------------------------------------------------------------------
Public Property Let WaveMute(ByVal NewValue&)      '# Mute-Status der Gesamtlautstärke setzen
  Call SetValue(volMute, NewValue)                 '# Set Mute State of main volume
End Property
'--------------------------------------------------------------------------------
Public Property Let MicroMute(ByVal NewValue&)     '# Mute-Status der Mikrofonlautstärke setzen
  Call SetValue(micMute, NewValue)                 '# Set Mute State of micro volume
End Property
'--------------------------------------------------------------------------------
Public Property Let WaveInMute(ByVal NewValue&)    '# Mute-Status der WaveInlautstärke setzen
  Call SetValue(wavMute, NewValue)                 '# Set Mute State of wave in volume
End Property
'--------------------------------------------------------------------------------
Public Property Let LineInMute(ByVal NewValue&)    '# Mute-Status der LineInlautstärke setzen
  Call SetValue(linMute, NewValue)                 '# Set Mute State of line in Volume
End Property
'--------------------------------------------------------------------------------
Public Property Let CD_Mute(ByVal NewValue&)       '# Mute-Status der CD-Lautstärke setzen
  Call SetValue(cdrMute, NewValue)                 '# Set Mute State of CD volume
End Property
'--------------------------------------------------------------------------------
Public Property Let MIDIMute(ByVal NewValue&)      '# Mute-Status der MIDI-Lautstärke setzen
  Call SetValue(midMute, NewValue)                 '# Set Mute State of MIDI volume
End Property
'********************************************************************************
' Functions to determine whether a control exists
'--------------------------------------------------------------------------------
Friend Function IsCtrl(ByVal Index&) As Boolean    '# True, wenn Control vorhanden
  IsCtrl = CtrlState And Index                     '# True, when control exists
End Function
'--------------------------------------------------------------------------------
Friend Function IsMute(ByVal Index&) As Boolean    '# True, wenn Control vorhanden
  IsMute = MuteState And Index                     '# True, when control exists
End Function
'********************************************************************************
'********************************************************************************
'# Hier werden die Grundparameter abgefragt bzw. gesetzt
'# Here the the fundamental parameters are determined rsp. set
Private Sub Class_Initialize()
  '---------------------------------------------------------------------
  '# Open the mixer with deviceID 0.
  rc = mixerOpen(hMixer, 0, 0, 0, 0)
  If rc <> MMSYSERR_NOERROR Then
    MsgBox "No mixer found!"
    Exit Sub
  End If
  '---------------------------------------------------------------------
  '# Get the waveout volume control  (Gesamtlautstärke/main volume)
  bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, volCtrl, 1)
  If bOK Then
    '# If the Function successfully gets the volume control, the maximum
    '# and minimum values are specified by lMaximum and lMinimum
    CtrlState = CtrlState Or 1
    With volCtrl
      MinVolVol = .lMinimum
      MaxVolVol = .lMaximum
    End With
    '# Get the waveout mute control
    bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, volMute, 2)
    If bOK Then MuteState = MuteState Or 1
  End If
  '---------------------------------------------------------------------
  '# Get the microphone volume control
  bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, micCtrl, 1)
  If bOK Then
    CtrlState = CtrlState Or 2
    With micCtrl
      MinMicVol = .lMinimum
      MaxMicVol = .lMaximum
    End With
    '# Get the microphone mute control
    bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, micMute, 2)
    If bOK Then MuteState = MuteState Or 2
  End If
  '---------------------------------------------------------------------
  '# Get the wave-in volume control
  bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, wavCtrl, 1)
  If bOK Then
    CtrlState = CtrlState Or 4
    With wavCtrl
      MinWavVol = .lMinimum
      MaxWavVol = .lMaximum
    End With
    '# Get the wave-in mute control
    bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, wavMute, 2)
    If bOK Then MuteState = MuteState Or 4
  End If
  '---------------------------------------------------------------------
  '# Get the line-in volume control
  bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_SRC_LINE, linCtrl, 1)
  If bOK Then
    CtrlState = CtrlState Or 8
    With linCtrl
      MinLinVol = .lMinimum
      MaxLinVol = .lMaximum
    End With
    '# Get the line-in mute control
    bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_SRC_LINE, linMute, 2)
    If bOK Then MuteState = MuteState Or 8
  Else  '# Some manufacturer use this type of device for line-in
    bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY, linCtrl, 1)
    If bOK Then
      CtrlState = CtrlState Or 8
      With linCtrl
        MinLinVol = .lMinimum
        MaxLinVol = .lMaximum
      End With
      '# Get the line-in mute control
      bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY, linMute, 2)
      If bOK Then MuteState = MuteState Or 8
    End If
  End If
  '---------------------------------------------------------------------
  '# Get the CD audio volume control
  bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC, cd_Ctrl, 1)
  If bOK Then
    CtrlState = CtrlState Or 16
    With cd_Ctrl
      MinCD_Vol = .lMinimum
      MaxCD_Vol = .lMaximum
    End With
    '# Get the CD audio mute control
    bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC, cdrMute, 2)
    If bOK Then MuteState = MuteState Or 16
  End If
  '---------------------------------------------------------------------
  '# Get the MIDI volume control
  bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER, midCtrl, 1)
  If bOK Then
    CtrlState = CtrlState Or 32
    With midCtrl
      MinMidVol = .lMinimum
      MaxMidVol = .lMaximum
    End With
    '# Get the MIDI mute control
    bOK = GetMixerControl(hMixer, MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER, midMute, 2)
    If bOK Then MuteState = MuteState Or 32
  End If
  'Debug.Print CtrlState
  'Debug.Print MuteState
  '---------------------------------------------------------------------
  '# In the same way may be opened other devices;
  '# for possible types refere to the constants
  '# MIXERLINE_COMPONENTTYPE_SRC_XXX and
  '# MIXERLINE_COMPONENTTYPE_DST_XXX
  '---------------------------------------------------------------------
End Sub

'********************************************************************************
'# This Function attempts to obtain a mixer control. Returns True if successful.

Private Function GetMixerControl(ByVal hMixer&, ByVal componentType&, _
                                 mxc As MIXERCONTROL, ByVal cType&) As Boolean
Dim ctrlType&, infoType&
  
  Select Case cType
  Case 1: ctrlType = MIXERCONTROL_CONTROLTYPE_VOLUME '# search for volume controls
          infoType = MIXER_GETLINEINFOF_COMPONENTTYPE
  Case 2: ctrlType = MIXERCONTROL_CONTROLTYPE_MUTE   '# search for mute controls
          infoType = MIXER_GETLINEINFOF_LINEID
  '...  To be extended
  Case Else: GetMixerControl = False: Exit Function
  End Select
    
  
  mxl.cbStruct = Len(mxl)
  mxl.dwComponentType = componentType
  
  '# Obtain a line corresponding to the component Private Type
  rc = mixerGetLineInfo(hMixer, mxl, infoType)
  
  If (MMSYSERR_NOERROR = rc) Then
    With mxlc
      .cbStruct = Len(mxlc)
      .dwLineID = mxl.dwLineID
      .dwControl = ctrlType
      .cControls = 1
      .cbmxctrl = Len(mxc)
      
      '# Allocate a buffer for the control
      hmem = GlobalAlloc(&H40, Len(mxc))
      .pamxctrl = GlobalLock(hmem)
      mxc.cbStruct = Len(mxc)
    End With
    '# Get the control
    rc = mixerGetLineControls(hMixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
    
    If (MMSYSERR_NOERROR = rc) Then
      GetMixerControl = True
      '# Copy the control into the destination structure
      CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
    Else
      GetMixerControl = False
    End If
    
    Call GlobalFree(hmem)
    Exit Function
  End If
  
  GetMixerControl = False
End Function

'********************************************************************************
'# This Function sets the value for a mixer control. Returns True if successful

Private Function SetValue(mxctl As MIXERCONTROL, ByVal volume&) As Boolean
                         
    mxc = mxctl
    Call PrepareStruct
    
    '# Copy the data into the control value buffer
    Struct.dwValue = volume
    CopyPtrFromStruct mxcd.paDetails, Struct, Len(Struct)
    
    '# Set the control value
    rc = mixerSetControlDetails(hMixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hmem)
    
    If (rc = MMSYSERR_NOERROR) Then
      SetValue = True
    Else
      SetValue = False
    End If
End Function

'********************************************************************************
'# Diese Funktion gibt den aktuell eingestellten Wert
'# für das an die Funktion übergebene MixerControl zurück
'# This function returns the current value of the passed mixercontrol

Private Function GetValue&(mxctl As MIXERCONTROL)
    
  mxc = mxctl
  Call PrepareStruct
  
  '# Get the control value
  rc = mixerGetControlDetails(hMixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
  
  '# Copy the data from control value buffer
  CopyStructFromPtr Struct, mxcd.paDetails, Len(Struct)
  Call GlobalFree(hmem)
  
  If (rc = MMSYSERR_NOERROR) Then
    GetValue = Struct.dwValue '# Aktuell eingestellter Wert / Current value
  Else
    GetValue = 0
  End If
End Function

'********************************************************************************
'# Initialisieren der MIXERCONTROLDETAILS-Struktur
'# Initialize the MIXERCONTROLDETAILS structure
Private Sub PrepareStruct()

Struct.dwValue = 0
With mxcd
  .item = 0
  .dwControlID = mxc.dwControlID
  .cbStruct = Len(mxcd)
  .cbDetails = Len(Struct)
  '# Allocate a buffer for the control value buffer
  hmem = GlobalAlloc(&H40, Len(Struct))
  .paDetails = GlobalLock(hmem)
  .cChannels = 1
End With
End Sub

'********************************************************************************
Private Sub Class_Terminate()
'# Close the mixer to free the memory
Call mixerClose(hMixer)
End Sub

'********************************************************************************

