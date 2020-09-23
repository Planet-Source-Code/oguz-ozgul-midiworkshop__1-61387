VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "G -"
      Height          =   255
      Left            =   9060
      TabIndex        =   19
      Top             =   360
      Width           =   435
   End
   Begin VB.CommandButton Command8 
      Caption         =   "G +"
      Height          =   255
      Left            =   9600
      TabIndex        =   18
      Top             =   360
      Width           =   435
   End
   Begin VB.CommandButton Command7 
      Caption         =   "N -"
      Height          =   255
      Left            =   9600
      TabIndex        =   17
      Top             =   60
      Width           =   435
   End
   Begin VB.CommandButton Command6 
      Caption         =   "N +"
      Height          =   255
      Left            =   9060
      TabIndex        =   16
      Top             =   60
      Width           =   435
   End
   Begin VB.CommandButton Command5 
      Caption         =   "x -"
      Height          =   255
      Left            =   7620
      TabIndex        =   15
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "y -"
      Height          =   255
      Left            =   8220
      TabIndex        =   14
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "y +"
      Height          =   255
      Left            =   8220
      TabIndex        =   13
      Top             =   60
      Width           =   495
   End
   Begin VB.CheckBox chkAlignToGrid 
      Caption         =   "Align To Grid"
      Height          =   195
      Left            =   2460
      TabIndex        =   10
      Top             =   60
      Width           =   1575
   End
   Begin VB.ComboBox cmbGridFrequency 
      Height          =   315
      ItemData        =   "pRoll.frx":0000
      Left            =   960
      List            =   "pRoll.frx":0019
      TabIndex        =   8
      Top             =   300
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "x +"
      Height          =   255
      Left            =   7620
      TabIndex        =   7
      Top             =   60
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      ScaleHeight     =   225
      ScaleWidth      =   10665
      TabIndex        =   6
      Top             =   660
      Width           =   10695
   End
   Begin VB.PictureBox picKeyboard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   4875
      Left            =   0
      ScaleHeight     =   323
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   5
      Top             =   900
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6540
      TabIndex        =   4
      Text            =   "4/4"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   360
      Width           =   1395
   End
   Begin VB.PictureBox picHScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   11640
      ScaleHeight     =   4845
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   900
      Width           =   255
   End
   Begin VB.PictureBox picVScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      ScaleHeight     =   225
      ScaleWidth      =   10665
      TabIndex        =   1
      Top             =   5760
      Width           =   10695
   End
   Begin VB.PictureBox picPRoll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   960
      ScaleHeight     =   323
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   711
      TabIndex        =   0
      Top             =   900
      Width           =   10695
   End
   Begin VB.Label lblKeySignature 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   45
      TabIndex        =   12
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblTimeSignature 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4/4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   45
      TabIndex        =   11
      Top             =   60
      Width           =   855
   End
   Begin VB.Label lblGridFrequency 
      Caption         =   "Grids Per Beat"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long

Private Type anEvent
    StartTick       As Long
    Channel         As Byte
    StatusByte      As Byte
    DataByte1       As Byte ' Note for note on events
    DataByte2       As Byte ' Velocity for note on events
    NextEventIndex  As Long
    Length          As Long ' Ticks until Note Off. 0 for other type of events
End Type

Private Enum MouseAction
    MOUSE_MOVE_NONE = 0
    MOUSE_MOVE_HORIZONTAL = 1
    MOUSE_MOVE_VERTICAL = 2
    MOUSE_MOVE_RESIZE = 3
End Enum

Private arrNoteNames(11)    As String
Private arrNotes()          As anEvent


' The two variables below, demonstrates the time signature.
' like 4/4, 2/4, 3/4, 6/8, 9/8 etc...

' the top number tells us how many beats exists in a measure (or bar)
' the bottom number tells us what kind of note gets the beat

' for the time signature 6/8:

' There will be 6 beats per measure each of which is a 8th note

Private lngBeatsPerMeasure  As Long ' Song specific
Private lngBeatNoteType     As Long ' Song specific


' After deciding how to divide measures into beats, we need to decide
' how to draw the grids. The piano roll can not be useful without grids.

' Grid Width and Grid height both are specified in pixels

Private lngGridsPerBeat         As Long ' User editable
Private lngGridWidth            As Double ' User editable
Private lngGridHeight           As Double ' User editable

Const NOTES_PER_OCTAVE          As Long = 12 ' Static value
Const GRID_LINE_WIDTH           As Long = 1 ' Static value

Const GRID_COLOR_NORMAL         As Long = &HD0D0D0
Const GRID_COLOR_BEAT           As Long = &HA0A0A0
Const GRID_COLOR_MEASURE        As Long = &HE08080
Const GRID_COLOR_OCTAVE         As Long = &HE08080

Const NOTE_COLOR                As Long = &HE060E0
Const NOTE_COLOR_BORDER         As Long = &O202020

Const GRID_WIDTH_DEFAULT        As Long = 6
Const GRID_HEIGHT_DEFAULT       As Long = 6
Const GRIDS_PER_QUARTER_BEAT    As Long = 8

Private lngStartNote            As Long

' The variables and constants above, if all set, are enough to start drawing the piano roll background


' After drawing the background, we need to draw the notes on it
' therefore we need to know how many ticks exist in a quarter note, better, in a grid

Private lngTicksPerQuarter      As Long ' Song Specific
Private lngTicksPerGrid         As Double ' User editable


' Detect the start grid, therefore the start tick to determine
' which notes will be displayed and which will not
Private lngStartTick            As Long
Private lngStartGrid            As Long


Private blnLeftMousePressed     As Boolean
Private lngNoteBeingHold        As Long
Private lngMouseAction          As MouseAction
Private lngMouseDownOffsetX     As Long
Private lngMouseDownOffsetY     As Long

Const MOUSE_ACTION_VERTICAL_DISTANCE       As Byte = 20
Const MOUSE_ACTION_HORIZONTAL_DISTANCE     As Long = 10

Private lngHMidi As Long

Private blnKeyDown As Boolean


Private Sub cmbGridFrequency_Click()
    lngGridWidth = lngGridWidth * (lngGridsPerBeat / CLng(cmbGridFrequency.Text))
    lngGridsPerBeat = CLng(cmbGridFrequency.Text)
    lngTicksPerGrid = lngTicksPerQuarter / (lngGridsPerBeat * (lngBeatNoteType / 4))
    ClearPianoRoll
    DrawGrids
    DrawNotes
End Sub

Private Sub Command1_Click()
    Dim arr() As String
    
    lngGridHeight = GRID_HEIGHT_DEFAULT
    lngGridWidth = GRID_WIDTH_DEFAULT
    arr = Split(Text1.Text, "/") ' Get time signature
    lngBeatsPerMeasure = arr(0)
    lngBeatNoteType = arr(1)
    lngGridsPerBeat = GRIDS_PER_QUARTER_BEAT * 4 / lngBeatNoteType
    cmbGridFrequency.Text = lngGridsPerBeat
    lngTicksPerQuarter = 120
    lngTicksPerGrid = lngTicksPerQuarter / (lngGridsPerBeat * (lngBeatNoteType / 4))
    lblTimeSignature.Caption = lngBeatsPerMeasure & "/" & lngBeatNoteType
    lngStartNote = 64
    lngStartGrid = 0
    lngStartTick = lngStartGrid * lngTicksPerGrid
    
    ClearPianoRoll
    DrawGrids
    DrawKeyboard False
    DrawNotes
End Sub

Private Function ClearPianoRoll()
    picPRoll.Cls
End Function

Private Function DrawGrids()
    Dim lngVGrid                    As Long
    Dim lngHGrid                    As Long
    Dim lngHeight                   As Long
    Dim lngWidth                    As Long
    Dim lngNumberOfVerticalGrids    As Long
    Dim lngNumberOfHorizontalGrids  As Long
    
    lngHeight = picPRoll.Height
    lngWidth = picPRoll.Width
    lngNumberOfHorizontalGrids = Int(lngHeight / (lngGridHeight + GRID_LINE_WIDTH))
    lngNumberOfVerticalGrids = Int(lngWidth / (lngGridWidth + GRID_LINE_WIDTH))
    
    For lngHGrid = 1 To lngNumberOfVerticalGrids
        If (lngHGrid + lngStartGrid) Mod lngGridsPerBeat = 0 Then ' Beat line
            If (lngHGrid + lngStartGrid) Mod lngGridsPerBeat * lngBeatsPerMeasure = 0 Then
                picPRoll.Line (lngHGrid * (lngGridWidth + GRID_LINE_WIDTH), 0)-(lngHGrid * (lngGridWidth + GRID_LINE_WIDTH), lngHeight), GRID_COLOR_MEASURE, BF
            Else
                picPRoll.Line (lngHGrid * (lngGridWidth + GRID_LINE_WIDTH), 0)-(lngHGrid * (lngGridWidth + GRID_LINE_WIDTH), lngHeight), GRID_COLOR_BEAT, BF
            End If
        Else
            picPRoll.Line (lngHGrid * (lngGridWidth + GRID_LINE_WIDTH), 0)-(lngHGrid * (lngGridWidth + GRID_LINE_WIDTH), lngHeight), GRID_COLOR_NORMAL, BF
        End If
    Next lngHGrid
    
    For lngVGrid = 1 To lngNumberOfHorizontalGrids
        If (lngStartNote - lngVGrid + 1) Mod NOTES_PER_OCTAVE = 0 Then
            picPRoll.Line (0, lngVGrid * (lngGridHeight + GRID_LINE_WIDTH))-(lngWidth, lngVGrid * (lngGridHeight + GRID_LINE_WIDTH)), GRID_COLOR_OCTAVE, BF
        Else
            picPRoll.Line (0, lngVGrid * (lngGridHeight + GRID_LINE_WIDTH))-(lngWidth, lngVGrid * (lngGridHeight + GRID_LINE_WIDTH)), GRID_COLOR_NORMAL, BF
        End If
    Next lngVGrid

End Function






Private Function DrawNotes()
    Dim lngIdx          As Long
    Dim lngX1           As Double
    Dim lngX2           As Double
    Dim lngY1           As Double
    Dim lngY2           As Double
    
    For lngIdx = 0 To UBound(arrNotes)
        With arrNotes(lngIdx)
            lngX1 = ((.StartTick - lngStartTick) / lngTicksPerGrid) * (lngGridWidth + GRID_LINE_WIDTH) + 1
            lngX2 = lngX1 + (.Length / lngTicksPerGrid) * (lngGridWidth + GRID_LINE_WIDTH) - 1
            lngY1 = (lngStartNote - .DataByte1) * (lngGridHeight + GRID_LINE_WIDTH) + 1
            lngY2 = lngY1 + lngGridHeight - 1
'            .X1 = lngX1
'            .X2 = lngX2
'            .Y1 = lngY1
'            .Y2 = lngY2
            picPRoll.Line (lngX1, lngY1)-(lngX2, lngY2), NOTE_COLOR_BORDER, B
            picPRoll.Line (lngX1 + 1, lngY1 + 1)-(lngX2 - 1, lngY2 - 1), NOTE_COLOR, BF
        End With
    Next lngIdx
End Function



Function Initialize(ByVal lngGridsPerBeat As Long, _
                    ByVal lngBeatsPerMeasure As Long, _
                    ByVal lngStartTick As Long)








End Function






Private Function DrawKeyboard(blnDisplayNoteNames As Boolean)
    Dim lngHGrid                    As Long
    Dim lngHeight                   As Long
    Dim lngWidth                    As Long
    Dim lngNumberOfHorizontalGrids  As Long
    Dim lngCurrentLineY             As Long
    Dim blnBlackNote                As Boolean
    Dim blnOnlyWhiteNote            As Boolean
    Dim strCurrentNoteName          As String
    Dim lngCurrentOctave            As Long
    Dim lngCurrentNote              As Long
    
    picKeyboard.Cls
    
    lngHeight = picKeyboard.Height
    lngWidth = picKeyboard.Width
    lngNumberOfHorizontalGrids = Int(lngHeight / (lngGridHeight + 1))
    
    lngCurrentOctave = Int(lngStartNote / 12) - 1
    
    For lngHGrid = 1 To lngNumberOfHorizontalGrids
        lngCurrentNote = lngStartNote - lngHGrid + 1
        lngCurrentLineY = lngHGrid * (lngGridHeight + 1)
        blnBlackNote = (lngCurrentNote Mod 12 = 1) Or (lngCurrentNote Mod 12 = 3) Or (lngCurrentNote Mod 12 = 6) Or (lngCurrentNote Mod 12 = 8) Or (lngCurrentNote Mod 12 = 10)
        If lngCurrentNote Mod 12 = 11 Then
            lngCurrentOctave = lngCurrentOctave - 1
        End If
        strCurrentNoteName = arrNoteNames(lngCurrentNote Mod 12) & lngCurrentOctave & " "
        If blnBlackNote Then
            picKeyboard.Line (0, lngCurrentLineY - lngGridHeight)-(picKeyboard.Width * 0.7, lngCurrentLineY), RGB(&H10, &H10, &H10), BF
            picKeyboard.Line (0, (lngCurrentLineY - lngGridHeight) + (lngGridHeight / 2))-(picKeyboard.Width, (lngCurrentLineY - lngGridHeight) + (lngGridHeight / 2)), RGB(&H10, &H10, &H10)
            picKeyboard.ForeColor = RGB(&HA0, &HA0, &HA0)
        Else
            picKeyboard.Line (0, lngCurrentLineY)-(picKeyboard.Width * 0.7, lngCurrentLineY), RGB(&H10, &H10, &H10)
            picKeyboard.ForeColor = RGB(&H80, &H80, &H80)
        End If
        blnOnlyWhiteNote = (lngCurrentNote Mod 12 = 0) Or (lngCurrentNote Mod 12 = 5)
        If blnOnlyWhiteNote Then
            picKeyboard.Line (0, lngCurrentLineY)-(picKeyboard.Width, lngCurrentLineY), RGB(&H10, &H10, &H10)
        End If
        If blnDisplayNoteNames Then
            If blnBlackNote Then
                picKeyboard.CurrentX = (picKeyboard.Width * 0.7) - 3 - picKeyboard.TextWidth(strCurrentNoteName)
                picKeyboard.CurrentY = (lngCurrentLineY - lngGridHeight) + ((lngGridHeight - picKeyboard.TextHeight(strCurrentNoteName)) / 2)
            Else
                picKeyboard.CurrentX = picKeyboard.Width - 3 - picKeyboard.TextWidth(strCurrentNoteName)
                If blnOnlyWhiteNote Then
                    picKeyboard.CurrentY = (lngCurrentLineY - lngGridHeight) + ((lngGridHeight - picKeyboard.TextHeight(strCurrentNoteName)) / 2)
                Else
                    picKeyboard.CurrentY = (lngCurrentLineY - lngGridHeight * 0.5) + ((lngGridHeight - picKeyboard.TextHeight(strCurrentNoteName)) / 2)
                End If
            End If
            picKeyboard.Print strCurrentNoteName
        End If
    Next lngHGrid



End Function

Private Sub Command2_Click()
    lngGridWidth = lngGridWidth + 1
    ClearPianoRoll
    DrawGrids
    DrawKeyboard False
    DrawNotes
End Sub

Private Sub Command3_Click()
    lngGridHeight = lngGridHeight + 1
    ClearPianoRoll
    DrawGrids
    DrawKeyboard False
    DrawNotes
End Sub

Private Sub Command4_Click()
    lngGridHeight = lngGridHeight - 1
    ClearPianoRoll
    DrawGrids
    DrawKeyboard False
    DrawNotes
End Sub

Private Sub Command5_Click()
    lngGridWidth = lngGridWidth - 1
    ClearPianoRoll
    DrawGrids
    DrawKeyboard False
    DrawNotes
End Sub

Private Sub Command6_Click()
    lngStartNote = lngStartNote + 1
    ClearPianoRoll
    DrawGrids
    DrawKeyboard False
    DrawNotes
End Sub

Private Sub Command7_Click()
    lngStartNote = lngStartNote - 1
    ClearPianoRoll
    DrawGrids
    DrawKeyboard False
    DrawNotes
End Sub

Private Sub Command9_Click()
    lngStartGrid = lngStartGrid - 1
    lngStartTick = lngStartGrid * lngTicksPerGrid
    ClearPianoRoll
    DrawGrids
    DrawNotes
End Sub

Private Sub Command8_Click()
    lngStartGrid = lngStartGrid + 1
    lngStartTick = lngStartGrid * lngTicksPerGrid
    ClearPianoRoll
    DrawGrids
    DrawNotes
End Sub


Private Sub Form_Load()
    
    MsgBox midiOutOpen(lngHMidi, 0, 0, 0, 0)
    
    lngNoteBeingHold = -1
    
    arrNoteNames(0) = "C"
    arrNoteNames(1) = "C#"
    arrNoteNames(2) = "D"
    arrNoteNames(3) = "Eb"
    arrNoteNames(4) = "E"
    arrNoteNames(5) = "F"
    arrNoteNames(6) = "F#"
    arrNoteNames(7) = "G"
    arrNoteNames(8) = "G#"
    arrNoteNames(9) = "A"
    arrNoteNames(10) = "Bb"
    arrNoteNames(11) = "B"
    ReDim arrNotes(24)
    Dim i As Long
    Randomize Timer
    arrNotes(0).DataByte1 = 50
    arrNotes(0).StartTick = 120
    'arrNotes(0).Length = 120
    For i = 1 To 24
        arrNotes(i).StartTick = 200 + i * 10
        'arrNotes(i).Length = Int(Rnd() * 600)
        arrNotes(i).DataByte1 = Int(Rnd() * 60)
        arrNotes(i).DataByte2 = 127
    Next i

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    midiOutClose lngHMidi
End Sub



Private Sub picKeyboard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    midiOutShortMsg lngHMidi, 127 * &H10000 + (lngStartNote - CLng(Y / (lngGridHeight + GRID_LINE_WIDTH))) * CLng(&H100) + &H99
    
End Sub

Private Sub picKeyboard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    midiOutShortMsg lngHMidi, &H7B * CLng(&H100) + &HB9

End Sub

Private Sub picPRoll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        blnLeftMousePressed = True
        If lngNoteBeingHold > -1 And lngMouseAction <> MOUSE_MOVE_NONE Then
            With arrNotes(lngNoteBeingHold)
                Select Case lngMouseAction
                    Case MOUSE_MOVE_HORIZONTAL
                        ' Get distance from left-top
                        lngMouseDownOffsetX = X - .X1
                    Case MOUSE_MOVE_VERTICAL
                        ' Get distance from center-top
                        lngMouseDownOffsetX = X - (.X1 + .X2) / 2
                    Case MOUSE_MOVE_RESIZE
                        lngMouseDownOffsetX = X - .X2
                End Select
                lngMouseDownOffsetY = Y - .Y1
                'Debug.Print lngMouseAction & ": " & lngMouseDownOffsetX & " - " & lngMouseDownOffsetY
            End With
        End If
    End If
End Sub

Private Sub picPRoll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Dim lngPoint As Long
    If Not blnLeftMousePressed Then
        'lngPoint = picPRoll.Point(X, Y)
        'If lngPoint = NOTE_COLOR Or lngPoint = NOTE_COLOR_BORDER Then
        lngNoteBeingHold = GetNoteByPosition(X, Y)
        If lngNoteBeingHold >= 0 Then
            With arrNotes(lngNoteBeingHold)
                If Abs(X - (.X1 + .X2) / 2) < MOUSE_ACTION_VERTICAL_DISTANCE Then 'And lngPoint = NOTE_COLOR Then
                    Form1.MousePointer = MousePointerConstants.vbSizeNS
                    lngMouseAction = MOUSE_MOVE_VERTICAL
                ElseIf Abs(X - .X1) < MOUSE_ACTION_HORIZONTAL_DISTANCE Then
                    Form1.MousePointer = MousePointerConstants.vbSizeWE
                    lngMouseAction = MOUSE_MOVE_HORIZONTAL
                ElseIf Abs(X - .X2) < MOUSE_ACTION_HORIZONTAL_DISTANCE Then
                    Form1.MousePointer = MousePointerConstants.vbSizeWE
                    lngMouseAction = MOUSE_MOVE_RESIZE
                Else
                    Form1.MousePointer = MousePointerConstants.vbArrow
                    lngMouseAction = MOUSE_MOVE_NONE
                End If
            End With
        Else
            Form1.MousePointer = MousePointerConstants.vbDefault
            lngNoteBeingHold = -1
        End If
    ElseIf lngNoteBeingHold > -1 And lngMouseAction <> MOUSE_MOVE_NONE Then
        ' Move or Resize Note
        With arrNotes(lngNoteBeingHold)
            Select Case lngMouseAction
                Case MOUSE_MOVE_VERTICAL
                    If Y < .Y1 Then
                        .DataByte1 = .DataByte1 + 1
                        midiOutShortMsg lngHMidi, &H7B * CLng(&H100) + &HB9
                        midiOutShortMsg lngHMidi, 127 * &H10000 + .DataByte1 * &H100 + &H99
                        .Y1 = .Y1 - lngGridHeight - GRID_LINE_WIDTH
                        ClearPianoRoll
                        DrawGrids
                        DrawNotes
                    ElseIf Y > .Y2 Then
                        .DataByte1 = .DataByte1 - 1
                        midiOutShortMsg lngHMidi, &H7B * CLng(&H100) + &HB9
                        midiOutShortMsg lngHMidi, 127 * &H10000 + .DataByte1 * &H100 + &H99
                        .Y1 = .Y1 + lngGridHeight + GRID_LINE_WIDTH
                        ClearPianoRoll
                        DrawGrids
                        DrawNotes
                    End If
                Case MOUSE_MOVE_HORIZONTAL
                    If X <> .X1 + lngMouseDownOffsetX Then
                        .X1 = X - lngMouseDownOffsetX
                        .StartTick = CalculateTickByHorizontalPosition(.X1)
                        ClearPianoRoll
                        DrawGrids
                        DrawNotes
                    End If
                
                Case MOUSE_MOVE_RESIZE
                    If X <> .X2 + lngMouseDownOffsetX Then
                        .X2 = X - lngMouseDownOffsetX
                        .Length = CalculateTickByHorizontalPosition(.X2) - CalculateTickByHorizontalPosition(.X1)
                        ClearPianoRoll
                        DrawGrids
                        DrawNotes
                    End If
            
            End Select
        End With
    End If
End Sub


Private Function GetNoteByPosition(ByVal X As Single, ByVal Y As Single)
    Dim lngIdx As Long
    For lngIdx = 0 To UBound(arrNotes)
        With arrNotes(lngIdx)
            If X >= .X1 And X <= .X2 And Y > .Y1 And Y < .Y2 Then
                GetNoteByPosition = lngIdx
                Exit Function
            End If
        End With
    Next lngIdx
    GetNoteByPosition = -1
End Function

Private Function CalculateTickByHorizontalPosition(ByVal X As Long) As Long
    CalculateTickByHorizontalPosition = lngStartTick + (X / (lngGridWidth + GRID_LINE_WIDTH)) * lngTicksPerGrid
End Function

Private Function ACalculateTickByHorizontalPosition(ByVal X As Long) As Long
    ACalculateTickByHorizontalPosition = lngStartTick + (X / (lngGridWidth + GRID_LINE_WIDTH)) * lngTicksPerGrid
End Function


Private Sub picPRoll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        midiOutShortMsg lngHMidi, &H7B * CLng(&H100) + &HB9
        If chkAlignToGrid.Value = vbChecked Then
            
        End If
        blnLeftMousePressed = False
        lngNoteBeingHold = -1
        lngMouseAction = MOUSE_MOVE_NONE
        lngMouseDownOffsetX = 0
        lngMouseDownOffsetY = 0
    End If
End Sub
