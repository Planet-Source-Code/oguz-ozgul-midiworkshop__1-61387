VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8580
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13920
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   9180
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu muFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuMidi 
      Caption         =   "&Midi"
      Begin VB.Menu mnuMidiTranspose 
         Caption         =   "&Transpose"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Option Compare Binary

Dim WithEvents ox       As MidiWorkShop.MidiFile
Attribute ox.VB_VarHelpID = -1
Private lngHMidi        As Long

'Function PlayTrack(ByVal lngTrack As Long)
'    Dim lngH        As Long
'    Dim lngHMidiOut As Long
'    Dim lngCount    As Long
'    Dim typEvent    As MidiEvent
'    Dim curTick     As Long
'    Dim lngShortMsg As Long
'    MsgBox midiOutOpen(lngHMidiOut, 0, vbNull, 0, CALLBACK_NULL)
'    lngCount = ox.NumberOfEvents
'    For lngH = 0 To lngCount
'        typEvent = ox.Events(lngH)
'        With typEvent
'            If .EventType = MIDI_EVENT_TYPE_MIDI Then
'                'If .Channel <= 11 Then
'                    If .DeltaTicks > 0 Then
'                        Sleep (.DeltaTicks) * 1
'                    End If
'                    curTick = curTick + .DeltaTicks
'                    lngShortMsg = PrepareShortMessage(typEvent)
'                    midiOutShortMsg lngHMidiOut, lngShortMsg
'                'End If
'            End If
'        End With
'    Next lngH
'    ' Close all notes
'    For lngH = 0 To 15
'        midiOutShortMsg lngHMidiOut, &H7B * CLng(&H100) + &HB0 + lngH
'    Next lngH
'    MsgBox midiOutClose(lngHMidiOut)
'End Function





Private Sub MDIForm_Load()
    'frmPiano.Show
    'frmBoxes.Show
End Sub

Private Sub mnuFileClose_Click()
    ox.StopMidi
    Set ox = Nothing
End Sub

Function AllNotesOff()
    Dim lngChannel As Long
    ' Close all notes
    For lngChannel = 0 To 15
        midiOutShortMsg lngHMidi, &H7B * CLng(&H100) + &HB0 + lngChannel
    Next lngChannel
    For lngChannel = 0 To 15
        'frmPiano.AllNotesOff lngChannel
        frmBoxes.AllNotesOff lngChannel
    Next lngChannel
End Function

Private Sub mnuFileOpen_Click()
    Dim strMsg As String
    Dim strFile As String
    Dim x As Long
    Dim lngCount As Long
    Dim OC As MIDIOUTCAPS
    
    Set ox = New MidiWorkShop.MidiFile
    
    'lngCount = midiOutGetNumDevs()
    'For x = 0 To lngCount - 1
    '    If midiOutGetDevCaps(x, OC, LenB(OC)) = MMSYSERR_NOERROR Then
    '        If InStr(OC.szPname, Chr(0)) Then
    '            Debug.Print "Device " & x & ": " & Left(OC.szPname, InStr(OC.szPname, Chr(0)) - 1) & ", Voices: " & OC.wVoices & ", Tech: " & GetMidiOutTechnology(OC.wTechnology) & ", Supports: " & GetMidiOutSupport(OC.dwSupport)
    '        Else
    '            Debug.Print "Device " & x & ": " & OC.szPname & ", Voices: " & OC.wVoices & ", Tech: " & GetMidiOutTechnology(OC.wTechnology) & ", Supports: " & GetMidiOutSupport(OC.dwSupport)
    '        End If
    '    End If
    'Next x
    
    cd1.DefaultExt = "*.mid"
    cd1.FileName = "*.mid"
    cd1.ShowOpen
    
    strFile = cd1.FileName
    If LCase(Right(strFile, 4)) = ".mid" Then
        x = Timer
        If ox.ReadFromMidiFile(strFile) = EXCEPTION_NONE Then
            'MsgBox "MIDI FILE HAS BEEN READ IN " & Timer - x
            strMsg = "Midi File has been opened: " & strFile & vbCrLf & _
                     "The midi file type is : MIDI File Format " & ox.FileFormat & vbCrLf & _
                     "The Number Of Tracks: " & ox.NumberOfTracks & vbCrLf
            If ox.DeltaTimeFormat = MIDI_DELTA_TIME_TYPE_SMPTE Then
                strMsg = strMsg & "The delta time format is SMPTE Time Code" & vbCrLf
            Else
                strMsg = strMsg & "The delta time format is " & ox.TicksPerQuarterNote & " Ticks Per Quarter note"
            End If
            'MsgBox strMsg, vbInformation
            
            'ox.SortEvents
            
            'PlayTrack 1
            
            'If ox.FileFormat = MIDI_FILE_TYPE_1 Then
            Call midiOutOpen(lngHMidi, 1, vbNull, 0, CALLBACK_NULL)
            
            frmPiano.Show
            frmPiano.Top = 0
            frmPiano.Width = 9495
            frmPiano.Height = 8325
            frmBoxes.Show
            DoEvents
            
            For x = 0 To 15
                frmPiano.chkChannel(x).Value = vbChecked
            Next x
            
            ox.PlayMidi lngHMidi, 0
            
            For x = 0 To 15
                midiOutShortMsg lngHMidi, &H7B * CLng(&H100) + &HB0 + x
            Next x
            
            Call midiOutClose(lngHMidi)
            Unload frmPiano
            Unload frmBoxes
            'End If
            
            'End
        Else
            MsgBox "an Exception occured while reading midi file"
        End If
    End If

End Sub

Private Sub mnuMidiTranspose_Click()
    frmTranspose.SetMidiFile ox
    frmTranspose.Show
End Sub

Private Sub ox_NoteOffSent(ByVal lngChannel As Long, ByVal lngNote As Long)
    frmPiano.drawNoteOnOff lngChannel, lngNote, False
    frmBoxes.drawNoteOnOff lngChannel, lngNote, False
End Sub

Private Sub ox_NoteOnSent(ByVal lngChannel As Long, ByVal lngNote As Long)
    frmPiano.drawNoteOnOff lngChannel, lngNote, True
    frmBoxes.drawNoteOnOff lngChannel, lngNote, True
End Sub

Function ChangeChannelStatus(ByVal lngChannel As Long, ByVal blnEnabled As Boolean) As Long
    ox.ChangeChannelStatus lngChannel, blnEnabled, lngHMidi
    If Not blnEnabled Then
        frmPiano.AllNotesOff lngChannel
        frmBoxes.AllNotesOff lngChannel
    End If
End Function
