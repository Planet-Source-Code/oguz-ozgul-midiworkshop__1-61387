VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MidiFile"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0
Option Compare Binary

Private MIDI_FILE_HEADER_DESCRIPTOR                         As String
Private MIDI_FILE_SUPPORTED_HEADER_LENGTH                   As String

Private Const MIDI_FILE_CHUNK_DESCRIPTOR_LENGTH             As Long = 4
Private Const MIDI_FILE_CHUNK_LENGTH_BYTE_COUNT             As Long = 4
Private Const MIDI_FILE_HEADER_FILE_TYPE_BYTE_COUNT         As Long = 2
Private Const MIDI_FILE_HEADER_NUMBER_OF_TRACKS_BYTE_COUNT  As Long = 2
Private Const MIDI_FILE_HEADER_DIVISION_BYTE_COUNT          As Long = 2

Private Const CLASS_NAME                                    As String = "MidiWorkShop.MidiFile"

Public Type MidiEvent
    PreviousEvent                   As Long
    Channel                         As Long
    DeltaTicks                      As Long
    EventType                       As MidiGeneralEventTypes
    MetaEventType                   As MidiMetaEventTypes
    MidiEventType                   As MidiMidiEventTypes
    EventDataLength                 As Long
    arrData()                       As Byte
    NextEvent                       As Long
End Type

Public Enum MidiFileTypes
    MIDI_FILE_TYPE_0 = 0
    MIDI_FILE_TYPE_1 = 1
    MIDI_FILE_TYPE_2 = 2
End Enum

Public Enum MidiDeltaTimeTypes
    MIDI_DELTA_TIME_TYPE_TICKS_PER_QUARTER_NOTE = 0
    MIDI_DELTA_TIME_TYPE_SMPTE = 1
End Enum

Public Enum MidiGeneralEventTypes
    MIDI_EVENT_TYPE_MIDI = 0
    MIDI_EVENT_TYPE_META = 1
    MIDI_EVENT_TYPE_SYSTEM_EXCLUSIVE = 2
End Enum

Public Enum MidiMetaEventTypes
    MIDI_META_EVENT_TYPE_DETERMINER = &HFF
    MIDI_META_EVENT_TYPE_SEQUENCE_NUMBER = &H0
    MIDI_META_EVENT_TYPE_TEXT_EVENT = &H1
    MIDI_META_EVENT_TYPE_COPYRIGHT_NOTICE = &H2
    MIDI_META_EVENT_TYPE_TRACK_NAME = &H3
    MIDI_META_EVENT_TYPE_INSTRUMENT_NAME = &H4
    MIDI_META_EVENT_TYPE_LYRIC = &H5
    MIDI_META_EVENT_TYPE_MARKER = &H6
    MIDI_META_EVENT_TYPE_CUE_POINT = &H7
    MIDI_META_EVENT_TYPE_CHANNEL_PREFIX = &H20
    MIDI_META_EVENT_TYPE_END_OF_TRACK = &H2F
    MIDI_META_EVENT_TYPE_TEMPO_CHANGE = &H51
    MIDI_META_EVENT_TYPE_SMPTE_OFFSET = &H54
    MIDI_META_EVENT_TYPE_TIME_SIGNATURE = &H58
    MIDI_META_EVENT_TYPE_KEY_SIGNATURE = &H59
    MIDI_META_EVENT_TYPE_SEQUENCER_SPECIFIC = &H7F
End Enum

Public Enum MidiMidiEventTypes
    MIDI_MIDI_EVENT_TYPE_NOTE_OFF = &H80
    MIDI_MIDI_EVENT_TYPE_NOTE_ON = &H90
    MIDI_MIDI_EVENT_TYPE_POLYPHONIC_KEY_PRESSURE = &HA0
    MIDI_MIDI_EVENT_TYPE_CONTROL_CHANGE = &HB0
    MIDI_MIDI_EVENT_TYPE_PROGRAM_CHANGE = &HC0
    MIDI_MIDI_EVENT_TYPE_CHANNEL_PRESSURE = &HD0
    MIDI_MIDI_EVENT_TYPE_PITCH_WHEEL_CHANGE = &HE0
End Enum

Public Enum MidiMidiEventDataLengths
    MIDI_MIDI_EVENT_DATA_LENGTH_NOTE_OFF = 2
    MIDI_MIDI_EVENT_DATA_LENGTH_NOTE_ON = 2
    MIDI_MIDI_EVENT_DATA_LENGTH_POLYPHONIC_KEY_PRESSURE = 2
    MIDI_MIDI_EVENT_DATA_LENGTH_CONTROL_CHANGE = 2
    MIDI_MIDI_EVENT_DATA_LENGTH_PROGRAM_CHANGE = 1
    MIDI_MIDI_EVENT_DATA_LENGTH_CHANNEL_PRESSURE = 1
    MIDI_MIDI_EVENT_DATA_LENGTH_PITCH_WHEEL = 2
End Enum

Public Enum MidiSMPTETimeCodes
    MIDI_SMPTE_TIME_CODE_24 = 0
    MIDI_SMPTE_TIME_CODE_25 = 1
    MIDI_SMPTE_TIME_CODE_30 = 2
    MIDI_SMPTE_TIME_CODE_30DROP = 3
End Enum

'----------------------------------------------------------------
' H E A D E R   I N F O R M A T I O N
'----------------------------------------------------------------
Private lngMidiFileType             As MidiWorkShop.MidiFileTypes
Private lngDeltaTimeFormat          As MidiWorkShop.MidiDeltaTimeTypes
Private lngNumberOfTracks           As Long
Private lngTicksPerQuarterNote      As Long
Private lngNegativeSMPTEFormat      As Long
Private lngTicksPerFrame            As Long
Private lngFileSize                 As Long
'----------------------------------------------------------------


Private arrTracks()                 As New MidiWorkShop.Track
Private arrEvents()                 As MidiEvent
Private blnStopped                  As Long
Private arrChannels(15)             As Boolean

Private dblDeltaTickMultiplier      As Double
Private lngTranspose                As Long

Public Event NoteOnSent(ByVal lngChannel As Long, ByVal lngNote As Long)

Public Event NoteOffSent(ByVal lngChannel As Long, ByVal lngNote As Long)


Public Function ReadFromMidiFile(ByVal strMidiFilePath As String) As MidiWorkShopExceptions
    Const PROCEDURE_NAME            As String = "ReadFromMidiFile"
    Dim objFileControl              As New MidiWorkShop.FileControl
    Dim objExceptionHandler         As New MidiWorkShop.ExceptionHandler
    Dim lngFileNumber               As Long
    Dim lngTrack                    As Long
    
    On Error GoTo ExceptionHandler
    
    If objFileControl.OpenBinaryFileForReading(strMidiFilePath) = EXCEPTION_NONE Then
        
        If ReadMidiFileHeader(objFileControl) = EXCEPTION_NONE Then
            ReDim arrTracks(lngNumberOfTracks - 1)
            Do While lngTrack < lngNumberOfTracks
                If arrTracks(lngTrack).ReadTrack(objFileControl) = EXCEPTION_NONE Then
                    ' Successfully read a track
                Else
                    ReadFromMidiFile = EXCEPTION_NOT_A_VALID_MIDI_FILE
                    objFileControl.CloseFile
                    Exit Function
                End If
                lngTrack = lngTrack + 1
            Loop
            ReadFromMidiFile = EXCEPTION_NONE
        Else
            ReadFromMidiFile = EXCEPTION_NOT_A_VALID_MIDI_FILE
        End If
        
    Else
        ReadFromMidiFile = EXCEPTION_FILE_NOT_FOUND
    End If
    objFileControl.CloseFile
    Exit Function
ExceptionHandler:
    ReadFromMidiFile = EXCEPTION_UNEXPECTED
    Call objExceptionHandler.LogException(err.Number, err.Source, err.Description, Erl, CLASS_NAME, PROCEDURE_NAME, "")
    Close
End Function


Private Function ReadMidiFileHeader(ByRef objFile As MidiWorkShop.FileControl) As MidiWorkShopExceptions
    Const PROCEDURE_NAME            As String = "ReadMidiFileHeader"
    Dim objExceptionHandler         As New MidiWorkShop.ExceptionHandler
    Dim objTools                    As New MidiWorkShop.Tools
    Dim arrData()                   As Byte
    Dim lngFileTypeTemporary        As Long
    Dim lngHeaderLength             As Long
    
    On Error GoTo ExceptionHandler
    
    If objFile.Read(MIDI_FILE_CHUNK_DESCRIPTOR_LENGTH, arrData) = EXCEPTION_NONE Then
        If objTools.CompareBinaryArrayToString(MIDI_FILE_HEADER_DESCRIPTOR, arrData) = EXCEPTION_NONE Then
            Erase arrData
            ' Read the header length
            If objFile.Read(MIDI_FILE_CHUNK_LENGTH_BYTE_COUNT, arrData) = EXCEPTION_NONE Then
                If objTools.CompareBinaryArrayToString(MIDI_FILE_SUPPORTED_HEADER_LENGTH, arrData) = EXCEPTION_NONE Then
                    ' Read the midi file descriptor MThd
                    If objFile.Read(MIDI_FILE_HEADER_FILE_TYPE_BYTE_COUNT, arrData) = EXCEPTION_NONE Then
                        lngFileTypeTemporary = (arrData(0) * &H100) + arrData(1)
                        Erase arrData
                        If lngFileTypeTemporary = MidiWorkShop.MidiFileTypes.MIDI_FILE_TYPE_0 Or _
                           lngFileTypeTemporary = MidiWorkShop.MidiFileTypes.MIDI_FILE_TYPE_1 Or _
                           lngFileTypeTemporary = MidiWorkShop.MidiFileTypes.MIDI_FILE_TYPE_2 Then
                            lngMidiFileType = lngFileTypeTemporary
                            ' Read Number Of Tracks
                            If objFile.Read(MIDI_FILE_HEADER_NUMBER_OF_TRACKS_BYTE_COUNT, arrData) = EXCEPTION_NONE Then
                                lngNumberOfTracks = (arrData(0) * &H100) + arrData(1)
                                Erase arrData
                                ' Midi file type 0 should have only one track, other type should have one or more
                                If (lngMidiFileType = MIDI_FILE_TYPE_0 And lngNumberOfTracks = 1) Or lngNumberOfTracks >= 1 Then
                                    'Read the Division
                                    If objFile.Read(MIDI_FILE_HEADER_DIVISION_BYTE_COUNT, arrData) = EXCEPTION_NONE Then
                                        If arrData(0) And &HF0 Then ' Bit 15 (most significant bit) is set
                                            ' Read SMPTE Format
                                            lngDeltaTimeFormat = MIDI_DELTA_TIME_TYPE_SMPTE
                                            lngTicksPerQuarterNote = 0
                                        
                                        
                                        Else ' Bit 15 (most significant bit) is not set
                                            ' Read Ticks per quarter note
                                            lngDeltaTimeFormat = MIDI_DELTA_TIME_TYPE_TICKS_PER_QUARTER_NOTE
                                            lngTicksPerFrame = 0
                                            lngNegativeSMPTEFormat = 0
                                            ' Bits 14 through 0 is the ticks per quarter note
                                            lngTicksPerQuarterNote = (arrData(0) * &H100) + arrData(1)
                                        End If
                                        ' Finalize by getting file size
                                        If objFile.GetSize(lngFileSize) = EXCEPTION_NONE Then
                                            ReadMidiFileHeader = EXCEPTION_NONE
                                        Else
                                            ReadMidiFileHeader = EXCEPTION_UNEXPECTED
                                        End If
                                    Else
                                        ReadMidiFileHeader = EXCEPTION_NOT_A_VALID_MIDI_FILE
                                    End If
                                Else
                                    ReadMidiFileHeader = EXCEPTION_NOT_A_VALID_MIDI_FILE
                                End If
                            Else
                                ReadMidiFileHeader = EXCEPTION_NOT_A_VALID_MIDI_FILE
                            End If
                        Else
                            ReadMidiFileHeader = EXCEPTION_NOT_A_VALID_MIDI_FILE
                        End If
                    Else
                        ReadMidiFileHeader = EXCEPTION_NOT_A_VALID_MIDI_FILE
                    End If
                Else
                    ReadMidiFileHeader = EXCEPTION_MIDI_FILE_TYPE_NOT_SUPPORTED
                End If
            Else
                ReadMidiFileHeader = EXCEPTION_NOT_A_VALID_MIDI_FILE
            End If
        Else
            ReadMidiFileHeader = EXCEPTION_NOT_A_VALID_MIDI_FILE
        End If
    Else
        ReadMidiFileHeader = EXCEPTION_NOT_A_VALID_MIDI_FILE
    End If

    Exit Function
ExceptionHandler:
    ReadMidiFileHeader = EXCEPTION_UNEXPECTED
    Call objExceptionHandler.LogException(err.Number, err.Source, err.Description, Erl, CLASS_NAME, PROCEDURE_NAME, "")
End Function

Public Property Get FileFormat() As MidiWorkShop.MidiFileTypes
    FileFormat = lngMidiFileType
End Property

Public Property Get NumberOfTracks() As Long
    NumberOfTracks = lngNumberOfTracks
End Property

Public Property Get DeltaTimeFormat() As MidiWorkShop.MidiDeltaTimeTypes
    DeltaTimeFormat = lngDeltaTimeFormat
End Property

Public Property Get TicksPerQuarterNote() As Long
    TicksPerQuarterNote = lngTicksPerQuarterNote
End Property

Public Property Get SMPTETimeCode() As MidiWorkShop.MidiSMPTETimeCodes
    SMPTETimeCode = lngNegativeSMPTEFormat
End Property

Public Property Get TicksPerFrame() As Long
    TicksPerFrame = lngTicksPerFrame
End Property

Public Property Get FileSize() As Long
    FileSize = lngFileSize
End Property


Public Function SortEvents()
    Dim lngTrack        As Long
    Dim lngTotalEvents  As Long
    Dim arrSeek()       As Long
    Dim arrTick()       As Long
    Dim lngCurrentTick  As Long
    Dim lngTestTick     As Long
    Dim lngMinTick      As Long
    Dim lngCurEvent     As Long
    Dim lngCurTick      As Long
    Dim lngLastTick     As Long
    
    For lngTrack = 0 To lngNumberOfTracks - 1
        lngTotalEvents = lngTotalEvents + arrTracks(lngTrack).NumberOfEvents
    Next
    
    ReDim arrEvents(lngTotalEvents - 1)
    ReDim arrSeek(lngNumberOfTracks - 1)
    ReDim arrTick(lngNumberOfTracks - 1)
    
    lngCurrentTick = 0
    
    ' Find the first event's track in midi file
    For lngTrack = 0 To lngNumberOfTracks - 1
        If arrTracks(lngTrack).NumberOfEvents > 0 Then
            arrTick(lngTrack) = arrTracks(lngTrack).Events(0).DeltaTicks
            lngTestTick = arrTracks(lngTrack).Events(0).DeltaTicks
            'If lngTestTick = 0 Then
            '    lngMinTick = 0
            '    Exit For
            'End If
            If lngMinTick > lngTestTick Then
                lngMinTick = lngTestTick
            End If
        End If
    Next
    
    lngCurEvent = -1
    lngCurTick = lngMinTick
    lngLastTick = 0
    Do While True
        For lngTrack = 0 To lngNumberOfTracks - 1
            lngTestTick = arrTick(lngTrack)
            Do While arrTick(lngTrack) = lngCurTick
                lngCurEvent = lngCurEvent + 1
                
                With arrTracks(lngTrack).Events(arrSeek(lngTrack))
                    If .EventDataLength > 0 Then
                        ReDim arrEvents(lngCurEvent).arrData(.EventDataLength)
                        CopyMemory arrEvents(lngCurEvent).arrData(0), .arrData(0), .EventDataLength
                    End If
                    arrEvents(lngCurEvent).Channel = .Channel
                    arrEvents(lngCurEvent).DeltaTicks = lngCurTick - lngLastTick ' SPECIAL HANDLING
                    arrEvents(lngCurEvent).EventDataLength = .EventDataLength
                    arrEvents(lngCurEvent).EventType = .EventType
                    arrEvents(lngCurEvent).MetaEventType = .MetaEventType
                    arrEvents(lngCurEvent).MidiEventType = .MidiEventType
                    arrEvents(lngCurEvent).NextEvent = .NextEvent
                    arrEvents(lngCurEvent).PreviousEvent = .PreviousEvent
                    'Debug.Print lngCurEvent & ": " & arrEvents(lngCurEvent).Channel
                End With
                If arrSeek(lngTrack) = arrTracks(lngTrack).NumberOfEvents - 1 Then
                    Exit Do
                End If
                arrSeek(lngTrack) = arrSeek(lngTrack) + 1
                arrTick(lngTrack) = arrTick(lngTrack) + arrTracks(lngTrack).Events(arrSeek(lngTrack)).DeltaTicks
                'lngTestTick = arrTick(lngTrack)
            Loop
            'arrSeek(lngTrack) = arrSeek(lngTrack) - 1
        Next lngTrack
        ' Find the minimum to continue
        lngLastTick = lngCurTick
        lngCurTick = 999999999#
        For lngTrack = 0 To lngNumberOfTracks - 1
            If arrSeek(lngTrack) < arrTracks(lngTrack).NumberOfEvents - 1 Then
                If lngCurTick > arrTick(lngTrack) Then
                    lngCurTick = arrTick(lngTrack)
                End If
            End If
        Next
        If lngCurTick = 999999999# Then
            Exit Function
        End If
    Loop
End Function


Public Function PlayMidi(ByVal lngHMidi As Long, ByVal lngStartTick As Long)
    Dim lngTrack            As Long
    Dim arrSeek()           As Long
    Dim arrTick()           As Long
    Dim lngTickToPlay       As Long
    Dim lngNextTickToPlay   As Long
    Dim lngSleepStep        As Long
    
    ReDim arrSeek(lngNumberOfTracks - 1)
    ReDim arrTick(lngNumberOfTracks - 1)
    blnStopped = False
    lngTickToPlay = 0
    lngNextTickToPlay = 999999999#
    Do While Not blnStopped
        For lngTrack = 0 To lngNumberOfTracks - 1
            If arrSeek(lngTrack) <= arrTracks(lngTrack).NumberOfEvents - 1 Then
                Do While arrTick(lngTrack) = lngTickToPlay
                    If arrChannels(arrTracks(lngTrack).Events(arrSeek(lngTrack)).Channel) Then
                        If arrTracks(lngTrack).Events(arrSeek(lngTrack)).EventType = MIDI_EVENT_TYPE_MIDI Then
                            midiOutShortMsg lngHMidi, PrepareShortMessage(arrTracks(lngTrack).Events(arrSeek(lngTrack)), lngTranspose)
                            If arrTracks(lngTrack).Events(arrSeek(lngTrack)).MidiEventType = MIDI_MIDI_EVENT_TYPE_NOTE_ON Then
                                If arrTracks(lngTrack).Events(arrSeek(lngTrack)).arrData(1) > 0 Then
                                    RaiseEvent NoteOnSent(arrTracks(lngTrack).Events(arrSeek(lngTrack)).Channel, arrTracks(lngTrack).Events(arrSeek(lngTrack)).arrData(0))
                                Else
                                    RaiseEvent NoteOffSent(arrTracks(lngTrack).Events(arrSeek(lngTrack)).Channel, arrTracks(lngTrack).Events(arrSeek(lngTrack)).arrData(0))
                                End If
                            ElseIf arrTracks(lngTrack).Events(arrSeek(lngTrack)).MidiEventType = MIDI_MIDI_EVENT_TYPE_NOTE_OFF Then
                                RaiseEvent NoteOffSent(arrTracks(lngTrack).Events(arrSeek(lngTrack)).Channel, arrTracks(lngTrack).Events(arrSeek(lngTrack)).arrData(0))
                            End If
                        ElseIf arrTracks(lngTrack).Events(arrSeek(lngTrack)).EventType = MIDI_EVENT_TYPE_META Then
                            If arrTracks(lngTrack).Events(arrSeek(lngTrack)).MetaEventType = MIDI_META_EVENT_TYPE_TEMPO_CHANGE Then
                                With arrTracks(lngTrack).Events(arrSeek(lngTrack))
                                    dblDeltaTickMultiplier = (.arrData(0) * &H10000 + .arrData(1) * CLng(&H100) + .arrData(2)) / 1000
                                    dblDeltaTickMultiplier = (dblDeltaTickMultiplier / lngTicksPerQuarterNote)
                                End With
                            End If
                        End If
                    End If
                    arrSeek(lngTrack) = arrSeek(lngTrack) + 1
                    If arrSeek(lngTrack) < arrTracks(lngTrack).NumberOfEvents - 1 Then
                        arrTick(lngTrack) = arrTick(lngTrack) + arrTracks(lngTrack).Events(arrSeek(lngTrack)).DeltaTicks
                        If arrTick(lngTrack) = lngTickToPlay Then
                            ' Play this event too
                        Else
                            ' Find ninimum delta ticks to the next event
                            If lngNextTickToPlay > arrTick(lngTrack) Then
                                lngNextTickToPlay = arrTick(lngTrack)
                            End If
                        End If
                    Else
                        arrTick(lngTrack) = 999999999#
                    End If
                Loop
                ' Find ninimum delta ticks to the next event
                If lngNextTickToPlay > arrTick(lngTrack) Then
                    lngNextTickToPlay = arrTick(lngTrack)
                End If
            End If
        Next
        If lngNextTickToPlay = 999999999# Then
            blnStopped = True
            Exit Do
        End If
        
        
        Sleep (lngNextTickToPlay - lngTickToPlay) * dblDeltaTickMultiplier
        DoEvents
        lngTickToPlay = lngNextTickToPlay
        lngNextTickToPlay = 999999999#
    Loop


End Function


Public Function StopMidi() As Long
    blnStopped = True
End Function

Private Sub Class_Initialize()
    Dim lngChannel As Long
    ' Supported Header Length
    MIDI_FILE_SUPPORTED_HEADER_LENGTH = ChrB(0) & ChrB(0) & ChrB(0) & ChrB(6)
    ' MThd:
    MIDI_FILE_HEADER_DESCRIPTOR = ChrB(&H4D) & ChrB(&H54) & ChrB(&H68) & ChrB(&H64)
    For lngChannel = 0 To 15
        arrChannels(lngChannel) = True
    Next lngChannel
End Sub


Public Function NumberOfEvents() As Long
    NumberOfEvents = UBound(arrEvents) - 1
End Function



Public Function Events(ByVal lngEvent As Long) As MidiEvent
    Events = arrEvents(lngEvent)
End Function

Public Function Tracks(ByVal lngTrack As Long) As MidiWorkShop.Track
    Set Tracks = arrTracks(lngTrack)
End Function

Public Function ChangeChannelStatus(ByVal lngChannel As Long, ByVal blnEnabled As Boolean, lngHMidi As Long) As Long
    arrChannels(lngChannel) = blnEnabled
    If Not blnEnabled Then
        midiOutShortMsg lngHMidi, &H7B * CLng(&H100) + &HB0 + lngChannel
    End If
End Function

Function Transpose() As Long
    Transpose = lngTranspose
End Function

Function SetTranspose(ByVal lngTrn As Long) As Long
    lngTranspose = lngTrn
End Function
