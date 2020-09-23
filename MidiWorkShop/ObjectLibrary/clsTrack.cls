VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Track"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0
Option Compare Binary

Private MIDI_FILE_TRACK_DESCRIPTOR                          As String
Private Const MIDI_FILE_CHUNK_LENGTH_BYTE_COUNT             As Long = 4
Private Const MIDI_FILE_CHUNK_DESCRIPTOR_LENGTH             As Long = 4
Private Const MIDI_EVENT_TYPE_DETERMINER_META               As Byte = &HFF
Private Const MIDI_EVENT_TYPE_DETERMINER_SYS_EX             As Byte = &HF0
Private Const MIDI_EVENT_TYPE_DETERMINER_SYS_EX_ESCAPE      As Byte = &HF7
Private lngCurrentTick                                      As Long
Private bytRunningStatus                                    As Byte

'----------------------------------------------------------------
' T R A C K   I N F O R M A T I O N
'----------------------------------------------------------------
Private lngTrackLength              As Long
Private lngNumberOfEvents           As Long
Private arrEvents()                 As MidiEvent
'----------------------------------------------------------------

Private lngMSPerQuarterNote         As Long

Friend Function ReadTrack(ByVal objFile As MidiWorkShop.FileControl) As MidiWorkShopExceptions
    Dim arrData()           As Byte
    Dim objTools            As New MidiWorkShop.Tools
    Dim lngTrackStart       As Long
    Dim lngStart            As Long
    If objFile.Read(MIDI_FILE_CHUNK_DESCRIPTOR_LENGTH, arrData) = EXCEPTION_NONE Then
        If objTools.CompareBinaryArrayToString(MIDI_FILE_TRACK_DESCRIPTOR, arrData) = EXCEPTION_NONE Then
            Erase arrData
            If objFile.Read(MIDI_FILE_CHUNK_LENGTH_BYTE_COUNT, arrData) = EXCEPTION_NONE Then
                If objTools.Convert4BytesToLong(arrData, lngTrackLength) = EXCEPTION_NONE Then
                    Erase arrData
                    ' Since each event needs at least 2 bytes (delta time min = 1 byte, event data min = 1 byte)
                    ReDim arrEvents(lngTrackLength / 2)
                    ' Read the whole track
                    lngTrackStart = objFile.GetSeek
                    If objFile.Read(lngTrackLength, arrData) = EXCEPTION_NONE Then
                        ' Read events one by one
                        Do While lngStart < lngTrackLength
                            If ReadNextEvent(arrData, lngStart) = EXCEPTION_NONE Then
                                ' Event successfully read
                            Else
                                ReadTrack = EXCEPTION_NOT_A_VALID_MIDI_FILE
                                Exit Function
                            End If
                        Loop
                        ReadTrack = EXCEPTION_NONE
                    Else
                        ReadTrack = EXCEPTION_UNEXPECTED
                    End If
                Else
                    ReadTrack = EXCEPTION_UNEXPECTED
                End If
            Else
                ReadTrack = EXCEPTION_NOT_A_VALID_MIDI_FILE
            End If
        Else
            ReadTrack = EXCEPTION_NOT_A_VALID_MIDI_FILE
        End If
    Else
        ReadTrack = EXCEPTION_NOT_A_VALID_MIDI_FILE
    End If
    Set objTools = Nothing
End Function




Private Function ReadNextEvent(ByRef arrData() As Byte, ByRef lngStart As Long) As MidiWorkShopExceptions
    Dim lngDeltaTime            As Long
    Dim objTools                As New MidiWorkShop.Tools
    Dim lngEventDataLength      As Long
    Dim bytEventDeterminer      As Byte
    Dim lngIdx                  As Long
    
    'If lngNumberOfEvents = 0 Then
    '    ReDim arrEvents(0)
    'Else
    '    ReDim Preserve arrEvents(lngNumberOfEvents)
    'End If
    If objTools.ReadVariableLengthQuantityFromArray(arrData, lngStart, lngDeltaTime) = EXCEPTION_NONE Then
        lngCurrentTick = lngCurrentTick + lngDeltaTime
        ' After reading the delta time, read the event type, which is allways 1 byte
        With arrEvents(lngNumberOfEvents)
            .DeltaTicks = lngDeltaTime
            .PreviousEvent = lngNumberOfEvents - 1
            .NextEvent = lngNumberOfEvents + 1
            If lngStart < UBound(arrData) Then
                If arrData(lngStart) And &H80 Then ' If bit 7 is set, change the running status. If not, use the running status
                    Select Case arrData(lngStart)
                        Case MIDI_EVENT_TYPE_DETERMINER_META
                            .EventType = MIDI_EVENT_TYPE_META
                            lngStart = lngStart + 1
                            .MetaEventType = arrData(lngStart)
                            lngStart = lngStart + 1
                            If objTools.ReadVariableLengthQuantityFromArray(arrData, lngStart, lngEventDataLength) = EXCEPTION_NONE Then
                                .EventDataLength = lngEventDataLength
                                If .EventDataLength > 0 And .MetaEventType <> MIDI_META_EVENT_TYPE_END_OF_TRACK Then
                                    Call objTools.CopyPartOfArray(arrData, .arrData, lngStart, .EventDataLength)
                                End If
                            Else
                                ReadNextEvent = EXCEPTION_NOT_A_VALID_MIDI_FILE
                                Exit Function
                            End If
                        Case MIDI_EVENT_TYPE_DETERMINER_SYS_EX, MIDI_EVENT_TYPE_DETERMINER_SYS_EX_ESCAPE
                            .EventType = MIDI_EVENT_TYPE_SYSTEM_EXCLUSIVE
                            lngStart = lngStart + 1
                            If objTools.ReadVariableLengthQuantityFromArray(arrData, lngStart, lngEventDataLength) = EXCEPTION_NONE Then
                                .EventDataLength = lngEventDataLength
                                Call objTools.CopyPartOfArray(arrData, .arrData, lngStart, .EventDataLength)
                            Else
                                ReadNextEvent = EXCEPTION_NOT_A_VALID_MIDI_FILE
                                Exit Function
                            End If
                        Case Else ' Midi Event
                            .EventType = MIDI_EVENT_TYPE_MIDI
                            bytRunningStatus = arrData(lngStart)
                            If bytRunningStatus Mod &H10 = 0 Then
                                bytEventDeterminer = arrData(lngStart)
                            Else
                                bytEventDeterminer = arrData(lngStart) - (arrData(lngStart) Mod &H10)
                            End If
                            lngStart = lngStart + 1
                            If bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_PITCH_WHEEL_CHANGE Then
                                .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_PITCH_WHEEL
                            ElseIf bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_CHANNEL_PRESSURE Then
                                .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_CHANNEL_PRESSURE
                            ElseIf bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_PROGRAM_CHANGE Then
                                .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_PROGRAM_CHANGE
                            ElseIf bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_CONTROL_CHANGE Then
                                .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_CONTROL_CHANGE
                            ElseIf bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_POLYPHONIC_KEY_PRESSURE Then
                                .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_POLYPHONIC_KEY_PRESSURE
                            ElseIf bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_NOTE_ON Then
                                .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_NOTE_ON
                            ElseIf bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_NOTE_OFF Then
                                .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_NOTE_OFF
                            Else
                                ReadNextEvent = EXCEPTION_MIDI_FILE_TYPE_NOT_SUPPORTED
                                Exit Function
                            End If
                            .MidiEventType = bytEventDeterminer
                            .Channel = bytRunningStatus - bytEventDeterminer
                            ' Read the data into array
                            Call objTools.CopyPartOfArray(arrData, .arrData, lngStart, .EventDataLength)
                    End Select
                Else ' Running status is used here. The event is a midi event
                    .EventType = MIDI_EVENT_TYPE_MIDI
                    bytEventDeterminer = bytRunningStatus
                    If bytRunningStatus Mod &H10 = 0 Then
                        bytEventDeterminer = bytRunningStatus
                    Else
                        bytEventDeterminer = bytRunningStatus - (bytRunningStatus Mod &H10)
                    End If
                    If bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_PITCH_WHEEL_CHANGE Then
                        .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_PITCH_WHEEL
                    ElseIf bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_CHANNEL_PRESSURE Then
                        .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_CHANNEL_PRESSURE
                    ElseIf bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_PROGRAM_CHANGE Then
                        .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_PROGRAM_CHANGE
                    ElseIf bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_CONTROL_CHANGE Then
                        .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_CONTROL_CHANGE
                    ElseIf bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_POLYPHONIC_KEY_PRESSURE Then
                        .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_POLYPHONIC_KEY_PRESSURE
                    ElseIf bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_NOTE_ON Then
                        .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_NOTE_ON
                    ElseIf bytEventDeterminer = MidiMidiEventTypes.MIDI_MIDI_EVENT_TYPE_NOTE_OFF Then
                        .EventDataLength = MidiMidiEventDataLengths.MIDI_MIDI_EVENT_DATA_LENGTH_NOTE_OFF
                    Else
                        ReadNextEvent = EXCEPTION_MIDI_FILE_TYPE_NOT_SUPPORTED
                        Exit Function
                    End If
                    .MidiEventType = bytEventDeterminer
                    .Channel = bytRunningStatus - bytEventDeterminer
                    ' Read the data into array
                    'lngStart = lngStart - 1
                    Call objTools.CopyPartOfArray(arrData, .arrData, lngStart, .EventDataLength)
               End If
            Else ' Use running status. No value bytes specified, use 0
                .EventType = MIDI_EVENT_TYPE_MIDI
                bytEventDeterminer = bytRunningStatus
                If bytRunningStatus Mod &H10 = 0 Then
                    bytEventDeterminer = bytRunningStatus
                Else
                    bytEventDeterminer = bytRunningStatus - (bytRunningStatus Mod &H10)
                End If
                .MidiEventType = bytEventDeterminer
                .Channel = bytRunningStatus - bytEventDeterminer
                .EventDataLength = arrEvents(lngNumberOfEvents - 1).EventDataLength
                If .EventDataLength > 0 Then
                    ReDim .arrData(.EventDataLength - 1)
                End If
                ' Let the values be 0
            End If
        End With
    Else
        ReadNextEvent = EXCEPTION_NOT_A_VALID_MIDI_FILE
        Exit Function
    End If
    
    lngNumberOfEvents = lngNumberOfEvents + 1
    Set objTools = Nothing
End Function

Private Sub Class_Initialize()
    ' MTrk:
    MIDI_FILE_TRACK_DESCRIPTOR = ChrB(&H4D) & ChrB(&H54) & ChrB(&H72) & ChrB(&H6B)
End Sub

Public Function NumberOfEvents() As Long
    NumberOfEvents = lngNumberOfEvents
End Function


Public Function Events(ByVal lngEvent As Long) As MidiEvent
    Events = arrEvents(lngEvent)
End Function



Public Function AddEvent(ByRef typMidiEvent As MidiEvent) As Long
    Dim lngIndexToCopy      As Long
    Dim lngLocalSize        As Long
    Dim lngNewSize          As Long
    Dim lngParamSize        As Long
    Dim lngDataSize         As Long
    Dim lngHMem             As Long
    lngIndexToCopy = GetEventIndexAtTick(typMidiEvent.DeltaTicks)
    lngDataSize = UBound(typMidiEvent.arrData)
    lngHMem = LocalHandle(arrEvents(lngIndexToCopy))
    lngLocalSize = LocalSize(lngHMem)
    ReDim Preserve arrEvents(lngNumberOfEvents)
    With arrEvents(lngNumberOfEvents)
        ReDim .arrData(lngDataSize)
        CopyMemory arrEvents(lngNumberOfEvents), typMidiEvent, LenB(typMidiEvent)
        CopyMemory .arrData(0), typMidiEvent.arrData(0), lngDataSize
    End With
    If lngIndexToCopy > -1 Then
        ' Shift events after "index to copy" to the right,
        CopyMemory arrEvents(lngIndexToCopy + 1), arrEvents(lngIndexToCopy), lngLocalSize
        ' Insert the new event into the location
        
    
    Else ' Add to the end
        ' Do nothing
    End If
End Function


Public Function GetEventIndexAtTick(ByVal lngTick As Long) As Long
    Dim lngIdx As Long
    Dim lngTotalTick As Long
    For lngIdx = 0 To lngNumberOfEvents - 1
        lngTotalTick = lngTotalTick + arrEvents(lngIdx).DeltaTicks
        If lngTotalTick >= lngTick Then
            GetEventIndexAtTick = lngIdx
            Exit Function
        End If
    Next lngIdx
    GetEventIndexAtTick = -1 ' No events at this tick or later
End Function


