Attribute VB_Name = "BASS_FX"
'=============================================================================
' BASS_FX 2.2 - Copyright (c) 2002-2005 (: JOBnik! :) [Arthur Aminov, ISRAEL]
'                                                     [http://www.jobnik.org]
'
'     bugs/suggestions/questions -> e-mail: bass_fx@jobnik.org
'    ----------------------------------------------------------
'
' NOTE: This module will only work with BASS_FX version 2.2
'       Check http://www.un4seen.com for any later versions of BASS_FX.BAS
'
' * Requires BASS & BASS.BAS 2.2 - available @ www.un4seen.com
'=============================================================================

' Error codes returned by BASS_ErrorGetCode
Public Const BASS_FX_ERROR_NODECODE = 100    ' Not a decoding channel
Public Const BASS_FX_ERROR_STEREO = 101      ' Only for stereo
Public Const BASS_FX_ERROR_BPMINUSE = 102    ' BPM detection is in use
Public Const BASS_FX_ERROR_BPMX2 = 103       ' BPM has already multed by 2

' Tempo / Reverse / BPM flag
Public Const BASS_FX_FREESOURCE = &H10000    ' Free the source handle as well?

'=============================================================================================
'   D S P (Digital Signal Processing)
'=============================================================================================

' DSP effects
Public Enum DSPFX
    BASS_FX_DSPFX_SWAP                       ' Swap channels: left<=>right : for STEREO Only!
    BASS_FX_DSPFX_ROTATE                     ' A channels volume ping-pong : for STEREO Only!
    BASS_FX_DSPFX_ECHO                       ' Echo
    BASS_FX_DSPFX_FLANGER                    ' Flanger
    BASS_FX_DSPFX_VOLUME                     ' Volume Amplifier : L/R for STEREO, L for MONO
    BASS_FX_DSPFX_PEAKEQ                     ' Peaking Equalizer
    BASS_FX_DSPFX_REVERB                     ' Reverb
    BASS_FX_DSPFX_LPF                        ' Low Pass Filter
    BASS_FX_DSPFX_S2M                        ' Stereo 2 Mono : for STEREO Only!
    BASS_FX_DSPFX_DAMP                       ' Dynamic Amplification
    BASS_FX_DSPFX_AUTOWAH                    ' Auto WAH
    BASS_FX_DSPFX_ECHO2                      ' Echo 2
    BASS_FX_DSPFX_PHASER                     ' Phaser
    BASS_FX_DSPFX_ECHO3                      ' Echo 3
    BASS_FX_DSPFX_CHORUS                     ' Chorus
    BASS_FX_DSPFX_APF                        ' All Pass Filter
    BASS_FX_DSPFX_COMPRESSOR                 ' Compressor
    BASS_FX_DSPFX_DISTORTION                 ' Distortion
End Enum

' Flanger
Public Type BASS_FX_DSPFLANGER
    fWetDry As Single                        ' [0.........1]
    fSpeed As Single                         ' [0.........n]
End Type

' Echo
Public Type BASS_FX_DSPECHO
    fLevel As Single                         ' [0.........1]
    lDelay As Long                           ' [1200..30000]
End Type

' Reverb
Public Type BASS_FX_DSPREVERB
    fLevel As Single                         ' [0.........1]
    lDelay As Long                           ' [1200..10000]
End Type

' Volume
Public Type BASS_FX_DSPVOLUME                ' L/R for STEREO, L for MONO for both channels
    fLeft As Single                          ' [0....1....n]
    fRight As Single                         ' [0....1....n]
End Type

' Peaking Equalizer
Public Type BASS_FX_DSPPEAKEQ
    lBand As Long                            ' [0..n] more bands => more memory & cpu usage
    lFreq As Long                            ' current samplerate of a stream/music
    fBandwidth As Single                     ' in octaves [0..4] - Q is not in use
    fQ As Single                             ' the EE kinda definition [0..1] - Bandwidth is not in use
    fCenter As Single                        ' eg. 1000Hz
    fGain As Single                          ' in dB. eg. [-15..0..+15]
End Type

' Low Pass Filter
Public Type BASS_FX_DSPLPF
    lFreq As Long                            ' current samplerate of a stream/music
    fResonance As Single                     ' [2..........10]
    fCutOffFreq As Single                    ' [200Hz..5000Hz]
End Type

' Dynamic Amplification
Public Type BASS_FX_DSPDAMP
    fTarget As Single                        ' target level                             [0.......1]
    fQuiet As Single                         ' quiet  level                             [0.......1]
    fRate As Single                          ' amp adjustment rate                      [e.g: 0.02]
    fGain As Single                          ' amplification level                      [0.......1]
    fDelay As Single                         ' delay in seconds before increasing level [0..Length]
End Type

' Auto WAH
Public Type BASS_FX_DSPAUTOWAH
    fDryMix As Single                        ' dry (unaffected) signal mix              [-2......2]
    fWetMix As Single                        ' wet (affected) signal mix                [-2......2]
    fFeedback As Single                      ' feedback                                 [-1......1]
    fRate As Single                          ' rate of sweep in cycles per second       [0<....<10]
    fRange As Single                         ' sweep range in octaves                   [0<....<10]
    fFreq As Single                          ' base frequency of sweep Hz               [0<...1000]
End Type

' Echo 2
Public Type BASS_FX_DSPECHO2
    fDryMix As Single                        ' dry (unaffected) signal mix              [-2......2]
    fWetMix As Single                        ' wet (affected) signal mix                [-2......2]
    fFeedback As Single                      ' feedback                                 [-1......1]
    fDelay As Single                         ' delay sec                                [0<......6]
End Type

' Phaser
Public Type BASS_FX_DSPPHASER
    fDryMix As Single                        ' dry (unaffected) signal mix              [-2......2]
    fWetMix As Single                        ' wet (affected) signal mix                [-2......2]
    fFeedback As Single                      ' feedback                                 [-1......1]
    fRate As Single                          ' rate of sweep in cycles per second       [0<....<10]
    fRange As Single                         ' sweep range in octaves                   [0<....<10]
    fFreq As Single                          ' base frequency of sweep                  [0<...1000]
End Type

' Echo 3
Public Type BASS_FX_DSPECHO3
    fDryMix As Single                        ' dry (unaffected) signal mix              [-2......2]
    fWetMix As Single                        ' wet (affected) signal mix                [-2......2]
    fDelay As Single                         ' delay sec                                [0<......6]
End Type

' Chorus
Public Type BASS_FX_DSPCHORUS
    fDryMix As Single                        ' dry (unaffected) signal mix              [-2......2]
    fWetMix As Single                        ' wet (affected) signal mix                [-2......2]
    fFeedback As Single                      ' feedback                                 [-1......1]
    fMinSweep As Single                      ' minimal delay ms                         [0<..<6000]
    fMaxSweep As Single                      ' maximum delay ms                         [0<..<6000]
    fRate As Single                          ' rate ms/s                                [0<...1000]
End Type

' All Pass Filter
Public Type BASS_FX_DSPAPF
    fGain As Single                          ' reverberation time                       [-1=<..<=1]
    fDelay As Single                         ' delay sec                                [0<....<=6]
End Type

' Compressor
Public Type BASS_FX_DSPCOMPRESSOR
    fThreshold As Single                     ' compressor threshold                     [0<=...<=1]
    fAttacktime As Single                    ' attack time ms                           [0<.<=1000]
    fReleasetime As Single                   ' release time ms                          [0<.<=5000]
End Type

' Distortion
Public Type BASS_FX_DSPDISTORTION
    fDrive As Single                         ' distortion drive                         [0<=...<=5]
    fDryMix As Single                        ' dry (unaffected) signal mix              [-5<=..<=5]
    fWetMix As Single                        ' wet (affected) signal mix                [-5<=..<=5]
    fFeedback As Single                      ' feedback                                 [-1<=..<=1]
    fVolume As Single                        ' distortion volume                        [0=<...<=2]
End Type

Public Declare Function BASS_FX_DSP_Set Lib "bass_fx.dll" (ByVal handle As Long, ByVal dsp_fx As DSPFX, ByVal priority As Long) As Long
'   Set any chosen DSP effect to any handle.
'   handle   : stream/music/wma/cd/any other supported add-on format
'   dsp_fx   : FX you wish to use
'   priority : The priority of the new DSP, which determines it's position in the DSP chain
'              DSPs with higher priority are called before those with lower.
'   RETURN   : TRUE if created (0=error!)

Public Declare Function BASS_FX_DSP_Remove Lib "bass_fx.dll" (ByVal handle As Long, ByVal dsp_fx As DSPFX) As Long
'   Remove chosen DSP effect.
'   handle : stream/music/wma/cd/any other supported add-on format
'   dsp_fx : FX you wish to remove
'   RETURN : TRUE if removed (0=error!)

Public Declare Function BASS_FX_DSP_SetParameters Lib "bass_fx.dll" (ByVal handle As Long, ByVal dsp_fx As DSPFX, ByRef par As Any) As Long
'   Set the parameters of a DSP effect.
'   handle : stream/music/wma/cd/any other supported add-on format
'   dsp_fx : FX you wish to set parameters to
'   par    : Pointer to the parameter structure
'   RETURN : TRUE if ok (0=error!)

Public Declare Function BASS_FX_DSP_GetParameters Lib "bass_fx.dll" (ByVal handle As Long, ByVal dsp_fx As DSPFX, ByRef par As Any) As Long
'   Retrieve the parameters of a DSP effect.
'   handle : stream/music/wma/cd/any other supported add-on format
'   dsp_fx : FX you wish to get parameters from
'   par    : Pointer to the parameter structure
'   RETURN : TRUE if ok (0=error!)

Public Declare Function BASS_FX_DSP_Reset Lib "bass_fx.dll" (ByVal handle As Long, ByVal dsp_fx As DSPFX) As Long
'   Call this function before changing position to avoid *clicks*
'   handle : stream/music/wma/cd/any other supported add-on format
'   dsp_fx : FX you wish to reset parameters of
'   RETURN : TRUE if ok (0=error!)

'=============================================================================================
'   TEMPO / PITCH SCALING / SAMPLERATE
'=============================================================================================

' tempo flags
Public Const BASS_FX_TEMPO_QUICKALGO = &H400000   ' Using quicker tempo change algorithm gain speed, lose quality
Public Const BASS_FX_TEMPO_NO_AAFILTER = &H800000 ' Not using FIR low-pass (anti-alias) filter gain speed, lose quality

' NOTE: Enable Tempo supported flags in BASS_FX_TempoCreate and the others to source handle.

Public Declare Function BASS_FX_TempoCreate Lib "bass_fx.dll" (ByVal chan As Long, ByVal flags As Long) As Long
'   Creates a resampling stream from a decoding channel.
'   chan     : a handle returned by:
'                   BASS_StreamCreateFile     : flags = BASS_STREAM_DECODE ...
'                   BASS_MusicLoad            : flags = BASS_MUSIC_DECODE ...
'                   BASS_WMA_StreamCreateFile : flags = BASS_STREAM_DECODE ...
'                   BASS_CD_StreamCreate      : flags = BASS_STREAM_DECODE ...
'                   * Any other add-on handle using a decoding channel.
'   flags    : BASS_SAMPLE_SOFTWARE/LOOP/3D/FX or BASS_STREAM_DECODE/AUTOFREE or
'              BASS_SPEAKER_xxx or BASS_FX_TEMPO_xxx or BASS_FX_FREESOURCE
'   RETURN   : the tempo stream handle (0=error!)

Public Declare Function BASS_FX_TempoGetSource Lib "bass_fx.dll" (ByVal chan As Long) As Long
'   Get the source channel handle.
'   chan   : tempo stream handle
'   RETURN : the source channel handle (0=error!)

Public Declare Function BASS_FX_TempoSet Lib "bass_fx.dll" (ByVal chan As Long, ByVal tempo As Single, ByVal samplerate As Single, ByVal pitch As Single) As Long
'   Set new values to tempo/rate/pitch to change its speed.
'   chan       : tempo stream (or source channel) handle
'   tempo      : in Percents  [-95%..0..+5000%]                 (-100 = leave current)
'   samplerate : in Hz, but calculates by the same % as tempo   (   0 = leave current)
'   pitch      : in Semitones [-60....0....+60]                 (-100 = leave current)
'   RETURN     : TRUE if ok (0=error!)

Public Declare Function BASS_FX_TempoGet Lib "bass_fx.dll" (ByVal chan As Long, ByRef tempo As Single, ByRef samplerate As Single, ByRef pitch As Single) As Long
'   Get tempo/rate/pitch values.
'   chan       : tempo stream (or source channel) handle
'   tempo      : current tempo          (0 if not in use)
'   samplerate : current samplerate     (0 if not in use)
'   pitch      : current pitch          (0 if not in use)
'   RETURN     : TRUE if ok (0=error!)

Public Declare Function BASS_FX_TempoGetRateRatio Lib "bass_fx.dll" (ByVal chan As Long) As Single
'   Get the ratio of the resulting rate and source rate (the resampling ratio).
'   chan   : tempo stream (or source channel) handle
'   RETURN : the ratio(0 = Error!)

'=============================================================================================
'   R E V E R S E
'=============================================================================================

' NOTE: 1. MODs won't load without BASS_MUSIC_PRESCAN flag.
'       2. Enable Reverse supported flags in BASS_FX_ReverseCreate and the others to source handle.

Public Declare Function BASS_FX_ReverseCreate Lib "bass_fx.dll" (ByVal chan As Long, ByVal dec_block As Single, ByVal flags As Long) As Long
'   Creates a Reversed stream from a decoding channel.
'   chan      : a handle returned by:
'                   BASS_StreamCreateFile     : flags = BASS_STREAM_DECODE ...
'                   BASS_MusicLoad            : flags = BASS_MUSIC_DECODE Or BASS_MUSIC_PRESCAN ...
'                   BASS_WMA_StreamCreateFile : flags = BASS_STREAM_DECODE ...
'                   BASS_CD_StreamCreate      : flags = BASS_STREAM_DECODE ...
'                   * Other stream add-on formats if created as decoded channel.
'               * For better MP3/2/1 Reverse playback use: BASS_STREAM_PRESCAN flag.
'   dec_block : decode in # of seconds blocks...
'               larger blocks = less seeking overhead but larger spikes.
'   flags     : BASS_SAMPLE_SOFTWARE/LOOP/3D/FX or BASS_STREAM_DECODE/AUTOFREE or
'               BASS_SPEAKER_xxx or BASS_FX_FREESOURCE
'   RETURN    : the reverse stream handle (0=error!)

Public Declare Function BASS_FX_ReverseGetSource Lib "bass_fx.dll" (ByVal chan As Long) As Long
'   Get the source channel handle.
'   chan   : reverse stream handle
'   RETURN : the source channel handle (0=error!)

'=============================================================================================
'   B P M (Beats Per Minute)
'=============================================================================================

' bpm flags
Public Const BASS_FX_BPM_BKGRND = 1   'If in use, then you can do other stuff while detection's in process
Public Const BASS_FX_BPM_MULT2 = 2    'If in use, then will auto multiply bpm by 2 (if BPM < MinBPM*2)

'-----------
' Option - 1 - Get BPM from a decoded channel
'--------------------------------------------
Public Declare Function BASS_FX_BPM_DecodeGet Lib "bass_fx.dll" (ByVal chan As Long, ByVal startSec As Single, ByVal endSec As Single, ByVal minMaxBPM As Long, ByVal flags As Long, ByVal proc As Long) As Single
'   Get the original BPM of a decoding channel.
'   chan      : a handle returned by:
'                   BASS_StreamCreateFile     : flags = BASS_STREAM_DECODE ...
'                   BASS_MusicLoad            : flags = BASS_MUSIC_DECODE Or BASS_MUSIC_PRESCAN ...
'                   BASS_WMA_StreamCreateFile : flags = BASS_STREAM_DECODE ...
'                   BASS_CD_StreamCreate      : flags = BASS_STREAM_DECODE ...
'                   * Any other add-on handle using a decoding channel.
'   startSec  : start detecting position in seconds
'   endSec    : end detecting position in seconds
'   minMaxBPM : set min & max bpm, e.g: MAKELONG(LOWORD.HIWORD), LO=Min, HI=Max. 0 = defaults 45/230
'   flags     : BASS_FX_BPM_xxx or BASS_FX_FREESOURCE
'   *proc     : user defined function to receive the process in percents, use NULL if not in use
'   RETURN    : the original BPM value (-1=error!)

'-----------
' Option - 2 - Auto get BPM by period of time in seconds
'-------------------------------------------------------
Public Declare Function BASS_FX_BPM_CallbackSet Lib "bass_fx.dll" (ByVal handle As Long, ByVal proc As Long, ByVal period As Single, ByVal minMaxBPM As Long, ByVal flags As Long, ByVal user As Long) As Long
'   Enable getting BPM value by period of time in seconds.
'   handle    : stream/music/wma/cd/any other supported add-on format
'   *proc     : user defined function to receive the bpm value
'   period    : detection period in seconds
'   minMaxBPM : set min & max bpm, e.g: MAKELONG(LOWORD.HIWORD), LO=Min, HI=Max. 0 = defaults 45/230
'   flags     : only BASS_FX_BPM_MULT2 flag is used
'   user      : user instance data to pass to the callback function.
'   RETURN    : TRUE if ok (0=error!)

Public Declare Function BASS_FX_BPM_CallbackReset Lib "bass_fx.dll" (ByVal handle As Long) As Long
'   Reset the buffers. Call this function after changing position.
'   handle : stream/music/wma/cd/any other supported add-on format
'   RETURN : TRUE if ok (0=error!)

'---------------------------------------
'  Functions to use with Both options.
'---------------------------------------

' translation options
Public Const BASS_FX_BPM_X2 = 0         ' Multiply the original BPM value by 2 (may be called only once & will change the original BPM as well!).
Public Const BASS_FX_BPM_2FREQ = 1      ' BPM value to Frequency.
Public Const BASS_FX_BPM_FREQ2 = 2      ' Frequency to BPM value.
Public Const BASS_FX_BPM_2PERCENT = 3   ' Percents to BPM value.
Public Const BASS_FX_BPM_PERCENT2 = 4   ' BPM value to Percents.

Public Declare Function BASS_FX_BPM_Translate Lib "bass_fx.dll" (ByVal handle As Long, ByVal val2tran As Single, ByVal trans As Long) As Single
'   Translate the given BPM to FREQ/PERCENT and vice versa or multiply BPM by 2.
'   handle   : stream/music/wma/cd/any other supported add-on format
'   val2tran : specify a value to translate to a given option (no matter if used X2).
'   trans    : any of the above translation option
'   RETURN   : new calculated value. (-1=error!)
'
'   NOTE     : This function will not detect the BPM, it will just translate the detected
'              original BPM value of a given handle.

Public Declare Sub BASS_FX_BPM_Free Lib "bass_fx.dll" (ByVal handle As Long)
'   Free all resources used by a given handle.
'
'   NOTE: If BASS_FX_FREESOURCE is used, then will free the decoding channel as well.
'         You can't set/get this flag with BASS_ChannelSetFlags/BASS_ChannelGetInfo.
'
'   handle : stream/music/wma/cd/any other supported add-on format

Public Sub BPMPROC(ByVal handle As Long, ByVal bpm As Single, ByVal user As Long)
'   CALLBACK FUNCTION! - Uses with Option 2
'
'   Get the BPM after period of time in seconds.
'   handle : handle that the BASS_FX_BPM_CallbackSet has applied to
'   bpm    : the new original bpm value
'   user   : the user instance data given when BASS_FX_BPM_CallbackSet was called.
End Sub

Public Sub BPMPROCESSPROC(ByVal chan As Long, ByVal percent As Single)
'   CALLBACK FUNCTION! - Uses with Option 1
'
'   Get the detection process in percents of a channel.
'   chan    : channel that the BASS_FX_BPM_DecodeGet has applied to
'   percent : the process in percents [0%..100%]
End Sub
