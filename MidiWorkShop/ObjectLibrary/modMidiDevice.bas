Attribute VB_Name = "modMidiDevice"
Option Explicit
Option Base 0
Option Compare Binary

Public Const MAXPNAMELEN = 32  '  max product name length (including NULL)

Public Type MIDIINCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
End Type

Public Type MIDIOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        wTechnology As Integer
        wVoices As Integer
        wNotes As Integer
        wChannelMask As Integer
        dwSupport As Long
End Type

Declare Function GetTickCount Lib "kernel32" () As Long

'  flags for wTechnology field of MIDIOUTCAPS structure
Public Const MOD_MIDIPORT = 1   '  output port
Public Const MOD_SYNTH = 2      '  generic internal synth
Public Const MOD_SQSYNTH = 3    '  square wave internal synth
Public Const MOD_FMSYNTH = 4    '  FM internal synth
Public Const MOD_MAPPER = 5     '  MIDI mapper

'  flags for dwSupport field of MIDIOUTCAPS
Public Const MIDICAPS_VOLUME = &H1          '  supports volume control
Public Const MIDICAPS_LRVOLUME = &H2        '  separate left-right volume control
Public Const MIDICAPS_CACHE = &H4

' GET & SET VOLUME
Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long

' GET NUMBER OF DEVICES INSTALLED
Declare Function midiInGetNumDevs Lib "winmm.dll" () As Long
Declare Function midiOutGetNumDevs Lib "winmm" () As Integer

' GET DEVICE CAPS
Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIINCAPS, ByVal uSize As Long) As Long
Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long

' OPEN / CLOSE A DEVICE

' midiInOpen

' lphMidiIn:            Pointer to an HMIDIIN handle. This location is filled with a handle identifying the opened MIDI input device. The handle is used to identify the device in calls to other MIDI input functions.
' uDeviceID:            Identifier of the MIDI input device to be opened.
' dwCallback:           Pointer to a callback function, a thread identifier, or a handle of a window called with information about incoming MIDI messages. For more information on the callback function, see MidiInProc.
' dwCallbackInstance:   User instance data passed to the callback function. This parameter is not used with window callback functions or threads.
' dwFlags:              Callback flag for opening the device and, optionally, a status flag that helps regulate rapid data transfers. It can be the following values.

Declare Function midiInOpen Lib "winmm.dll" (lphMidiIn As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiInClose Lib "winmm.dll" (ByVal hMidiIn As Long) As Long

' midiOutOpen

' lphmo:                Pointer to an HMIDIOUT handle. This location is filled with a handle identifying the opened MIDI output device. The handle is used to identify the device in calls to other MIDI output functions.
' uDeviceID:            Identifier of the MIDI output device that is to be opened.
' dwCallback:           Pointer to a callback function, an event handle, a thread identifier, or a handle of a window or thread called during MIDI playback to process messages related to the progress of the playback. If no callback is desired, specify NULL for this parameter. For more information on the callback function, see MidiOutProc.
' dwCallbackInstance:   User instance data passed to the callback. This parameter is not used with window callbacks or threads.
' dwFlags:              Callback flag for opening the device. It can be the following values.

' Value Meaning

' CALLBACK_EVENT:       The dwCallback parameter is an event handle. This callback mechanism is for output only.
' CALLBACK_FUNCTION:    The dwCallback parameter is a callback function address.
' CALLBACK_NULL:        There is no callback mechanism. This value is the default setting.
' CALLBACK_THREAD:      The dwCallback parameter is a thread identifier.
' CALLBACK_WINDOW:      The dwCallback parameter is a window handle.

' Return Values

' MMSYSERR_NOERROR:     if successful or an error otherwise. Possible error values include the following.
' MIDIERR_NODEVICE:     No MIDI port was found. This error occurs only when the mapper is opened.
' MMSYSERR_ALLOCATED:   The specified resource is already allocated.
' MMSYSERR_BADDEVICEID: The specified device identifier is out of range.
' MMSYSERR_INVALPARAM:  The specified pointer or structure is invalid.
' MMSYSERR_NOMEM:       The system is unable to allocate or lock memory.

' REMARKS:
' To determine the number of MIDI output devices present in the system, use the midiOutGetNumDevs function. The device identifier specified by wDeviceID varies from zero to one less than the number of devices present. MIDI_MAPPER can also be used as the device identifier.
' If a window or thread is chosen to receive callback information, the following messages are sent to the window procedure or thread to indicate the progress of MIDI output: MM_MOM_OPEN, MM_MOM_CLOSE, and MM_MOM_DONE.
' If a function is chosen to receive callback information, the following messages are sent to the function to indicate the progress of MIDI output: MOM_OPEN, MOM_CLOSE, and MOM_DONE.

Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiOutClose Lib "winmm.dll" (ByVal HMIDIOUT As Long) As Long

' SEND MESSAGE

' REMARKS:
' This function is used to send any MIDI message except for system-exclusive or stream messages.
' This function might not return until the message has been sent to the output device.
' You can send short messages while streams are playing on the same device
' (although you cannot use a running status in this case).
Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal HMIDIOUT As Long, ByVal dwMsg As Long) As Long




' GET DEVICE ID

Declare Function midiOutGetID Lib "winmm.dll" (ByVal HMIDIOUT As Long, lpuDeviceID As Long) As Long
Declare Function midiInGetID Lib "winmm.dll" (ByVal hMidiIn As Long, lpuDeviceID As Long) As Long


'Declare Function midiOutPrepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
'Declare Function midiOutUnprepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
'Declare Function midiOutLongMsg Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
'Declare Function midiOutMessage Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

'Declare Function midiOutReset Lib "winmm.dll" (ByVal HMIDIOUT As Long) As Long
'Declare Function midiOutCachePatches Lib "winmm.dll" (ByVal HMIDIOUT As Long, ByVal uBank As Long, lpPatchArray As Long, ByVal uFlags As Long) As Long
'Declare Function midiOutCacheDrumPatches Lib "winmm.dll" (ByVal HMIDIOUT As Long, ByVal uPatch As Long, lpKeyArray As Long, ByVal uFlags As Long) As Long


Declare Function midiInGetErrorText Lib "winmm.dll" Alias "midiInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long

'Declare Function midiInPrepareHeader Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
''Declare Function midiInUnprepareHeader Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
'Declare Function midiInAddBuffer Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiInStart Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInStop Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInReset Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInMessage Lib "winmm.dll" (ByVal hMidiIn As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'  flags for wFlags parameter of timeSetEvent() function
Public Const TIME_ONESHOT = 0  '  program timer for single event
Public Const TIME_PERIODIC = 1  '  program for continuous periodic event

' Driver callback support

'  flags used with waveOutOpen(), waveInOpen(), midiInOpen(), and
'  midiOutOpen() to specify the type of the dwCallback parameter.
Public Const CALLBACK_TYPEMASK = &H70000      '  callback type mask
Public Const CALLBACK_NULL = &H0        '  no callback
Public Const CALLBACK_WINDOW = &H10000      '  dwCallback is a HWND
Public Const CALLBACK_TASK = &H20000      '  dwCallback is a HTASK
Public Const CALLBACK_FUNCTION = &H30000      '  dwCallback is a FARPROC


Public Type TIMECAPS
    wPeriodMin As Long
    wPeriodMax As Long
End Type

' Applications should not call any system-defined functions from inside a callback function, except for

' PostMessage
' timeGetSystemTime
' timeGetTime
' timeSetEvent
' timeKillEvent
' midiOutShortMsg
' midiOutLongMsg
' OutputDebugString


'Remarks

'Applications should not call any system-defined functions from inside a callback function, except for

'EnterCriticalSection
'LeaveCriticalSection
'midiOutLongMsg
'midiOutShortMsg
'OutputDebugString
'PostMessage
'PostThreadMessage
'SetEvent
'timeGetSystemTime
'timeGetTime
'timeKillEvent
'timeSetEvent


Public Function GetMidiOutTechnology(ByVal lngTechnology As Long) As String
    Select Case lngTechnology
        Case MOD_MIDIPORT
            GetMidiOutTechnology = "MIDI OUTPUT PORT"
        Case MOD_FMSYNTH
            GetMidiOutTechnology = "FM INTERNAL SYNTHYSIZER"
        Case MOD_MAPPER
            GetMidiOutTechnology = "MIDI MAPPER"
        Case MOD_SQSYNTH
            GetMidiOutTechnology = "SQUARE WAVE INTERNAL SYNTHYSIZER"
        Case MOD_SYNTH
            GetMidiOutTechnology = "GENERIC INTERNAL SYNTHYSIZER"
        Case Else
            GetMidiOutTechnology = "UNKNOWN"
    End Select
End Function

Public Function GetMidiOutSupport(ByVal lngSupport As Long) As String
    Dim strReturn       As String
    If lngSupport And MIDICAPS_VOLUME Then
        strReturn = "VOLUME CONTROL"
    End If
    If lngSupport And MIDICAPS_LRVOLUME Then
        If strReturn <> "" Then
            strReturn = strReturn & ", "
        End If
        strReturn = strReturn & "LEFT-RIGHT VOLUME"
    End If
    If lngSupport And MIDICAPS_CACHE Then
        If strReturn <> "" Then
            strReturn = strReturn & ", "
        End If
        strReturn = strReturn & "CACHING"
    End If
    If strReturn = "" Then
        strReturn = "NONE"
    End If
    GetMidiOutSupport = strReturn
End Function


Public Function PrepareShortMessage(ByRef typEvent As MidiEvent, ByVal lngTranspose As Long) As Long
    Dim lngResult As Long
    With typEvent
        Select Case .MidiEventType
            Case MIDI_MIDI_EVENT_TYPE_NOTE_OFF, MIDI_MIDI_EVENT_TYPE_NOTE_ON
                If typEvent.Channel = 9 Then
                    lngTranspose = 0
                End If
                lngResult = .arrData(1) * &H10000 + (.arrData(0) + lngTranspose) * CLng(&H100) + .MidiEventType + .Channel
            Case MIDI_MIDI_EVENT_TYPE_POLYPHONIC_KEY_PRESSURE, MIDI_MIDI_EVENT_TYPE_CONTROL_CHANGE, MIDI_MIDI_EVENT_TYPE_PITCH_WHEEL_CHANGE
                lngResult = .arrData(1) * &H10000 + .arrData(0) * CLng(&H100) + .MidiEventType + .Channel
            Case MIDI_MIDI_EVENT_TYPE_PROGRAM_CHANGE, MIDI_MIDI_EVENT_TYPE_CHANNEL_PRESSURE
                lngResult = .arrData(0) * CLng(&H100) + .MidiEventType + .Channel
        End Select
    End With
    PrepareShortMessage = lngResult
End Function





' PROCEDURE FOR CAPTURING MIDI IN CALLBACKS

Function MidiInProc(ByVal hMidiIn As Long, ByVal wMsg As Long, ByVal dwInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long)

End Function

' PROCEDURE FOR CAPTURING MIDI OUT CALLBACKS

Function MidiOutProc(ByVal HMIDIOUT As Long, ByVal wMsg As Integer, ByVal dwInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long)

End Function




