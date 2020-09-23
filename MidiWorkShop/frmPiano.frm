VERSION 5.00
Begin VB.Form frmPiano 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   9405
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   15
      Left            =   780
      TabIndex        =   42
      Top             =   7560
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   15
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   41
      Top             =   7440
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   14
      Left            =   780
      TabIndex        =   40
      Top             =   7080
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   14
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   39
      Top             =   6960
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   13
      Left            =   780
      TabIndex        =   38
      Top             =   6600
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   13
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   37
      Top             =   6480
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   12
      Left            =   780
      TabIndex        =   36
      Top             =   6120
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   12
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   35
      Top             =   6000
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   11
      Left            =   780
      TabIndex        =   34
      Top             =   5640
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   11
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   33
      Top             =   5520
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   10
      Left            =   780
      TabIndex        =   32
      Top             =   5160
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   10
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   31
      Top             =   5040
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   9
      Left            =   780
      TabIndex        =   30
      Top             =   4680
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   9
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   29
      Top             =   4560
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   8
      Left            =   780
      TabIndex        =   27
      Top             =   4200
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   8
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   26
      Top             =   4080
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   7
      Left            =   780
      TabIndex        =   18
      Top             =   3720
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   7
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   17
      Top             =   3600
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   6
      Left            =   780
      TabIndex        =   16
      Top             =   3240
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   6
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   15
      Top             =   3120
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   5
      Left            =   780
      TabIndex        =   14
      Top             =   2760
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   5
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   13
      Top             =   2640
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   4
      Left            =   780
      TabIndex        =   12
      Top             =   2280
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   4
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   11
      Top             =   2160
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   3
      Left            =   780
      TabIndex        =   10
      Top             =   1800
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   3
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   9
      Top             =   1680
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   2
      Left            =   780
      TabIndex        =   8
      Top             =   1320
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   2
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   7
      Top             =   1200
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   1
      Left            =   780
      TabIndex        =   6
      Top             =   840
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   1
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   5
      Top             =   720
      Width           =   7725
   End
   Begin VB.CheckBox chkChannel 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   0
      Left            =   780
      TabIndex        =   1
      Top             =   360
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox picPiano 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   0
      Left            =   1140
      ScaleHeight     =   405
      ScaleWidth      =   7695
      TabIndex        =   0
      Top             =   240
      Width           =   7725
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   15
      Left            =   8940
      TabIndex        =   66
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   14
      Left            =   8940
      TabIndex        =   65
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   13
      Left            =   8940
      TabIndex        =   64
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   12
      Left            =   8940
      TabIndex        =   63
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   11
      Left            =   8940
      TabIndex        =   62
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   10
      Left            =   8940
      TabIndex        =   61
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   9
      Left            =   8940
      TabIndex        =   60
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   8
      Left            =   8940
      TabIndex        =   59
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   7
      Left            =   8940
      TabIndex        =   58
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   6
      Left            =   8940
      TabIndex        =   57
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   5
      Left            =   8940
      TabIndex        =   56
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   8940
      TabIndex        =   55
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   8940
      TabIndex        =   54
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   8940
      TabIndex        =   53
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   8940
      TabIndex        =   52
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H80000003&
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   8940
      TabIndex        =   51
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Notes"
      Height          =   255
      Left            =   8880
      TabIndex        =   50
      Top             =   60
      Width           =   495
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "16"
      Height          =   195
      Index           =   15
      Left            =   180
      TabIndex        =   49
      Top             =   7560
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "15"
      Height          =   195
      Index           =   14
      Left            =   180
      TabIndex        =   48
      Top             =   7080
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "14"
      Height          =   195
      Index           =   13
      Left            =   180
      TabIndex        =   47
      Top             =   6600
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "13"
      Height          =   195
      Index           =   12
      Left            =   180
      TabIndex        =   46
      Top             =   6180
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "12"
      Height          =   195
      Index           =   11
      Left            =   180
      TabIndex        =   45
      Top             =   5700
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "11"
      Height          =   195
      Index           =   10
      Left            =   180
      TabIndex        =   44
      Top             =   5220
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "10"
      Height          =   195
      Index           =   9
      Left            =   180
      TabIndex        =   43
      Top             =   4680
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "9"
      Height          =   195
      Index           =   8
      Left            =   180
      TabIndex        =   28
      Top             =   4200
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "8"
      Height          =   195
      Index           =   7
      Left            =   180
      TabIndex        =   25
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "7"
      Height          =   195
      Index           =   6
      Left            =   180
      TabIndex        =   24
      Top             =   3240
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "6"
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   23
      Top             =   2760
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "5"
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   22
      Top             =   2340
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "4"
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   21
      Top             =   1860
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "3"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   20
      Top             =   1380
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "2"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   19
      Top             =   840
      Width           =   315
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "1"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   360
      Width           =   315
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000003&
      Caption         =   "Channel"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "Play"
      Height          =   255
      Left            =   780
      TabIndex        =   2
      Top             =   60
      Width           =   375
   End
End
Attribute VB_Name = "frmPiano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrNoteWidths(11) As Single
Dim arrPositions(127) As Single

Private Const NOTE_WIDTH As Long = 60
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long



Private Sub chkChannel_Click(Index As Integer)
    If chkChannel(Index).Value = vbChecked Then
        mdiMain.ChangeChannelStatus Index, True
    Else
        mdiMain.ChangeChannelStatus Index, False
    End If
End Sub

Private Sub Form_Load()
    
    Dim lngNote As Long
    Dim lngPos As Long
    
    arrNoteWidths(0) = 1.5
    arrNoteWidths(1) = 1
    arrNoteWidths(2) = 2
    arrNoteWidths(3) = 1
    arrNoteWidths(4) = 1.5
    arrNoteWidths(5) = 1.5
    arrNoteWidths(6) = 1
    arrNoteWidths(7) = 2
    arrNoteWidths(8) = 1
    arrNoteWidths(9) = 2
    arrNoteWidths(10) = 1
    arrNoteWidths(11) = 1.5

    For lngNote = 0 To 127
        arrPositions(lngNote) = lngPos
        Select Case lngNote Mod 12
            Case 1, 3, 6, 8, 10
            Case Else
                lngPos = lngPos + arrNoteWidths(lngNote Mod 12) * NOTE_WIDTH
        End Select
    Next lngNote
    
    For lngNote = 0 To 15
        picPiano(lngNote).Picture = LoadPicture(App.Path & "\graphics\proll.bmp")
    Next lngNote

End Sub

Private Sub Form_Paint()
    Dim lngChn As Long
    Dim lngNote As Long
    
    For lngChn = 0 To 15
        With picPiano(lngChn)
            
            For lngNote = 0 To 127
                'If lngNote = 48 Then MsgBox ""
                Select Case lngNote Mod 12
                    Case 1, 3, 6, 8, 10
                        picPiano(lngChn).Line (arrPositions(lngNote) - NOTE_WIDTH / 2, 0)-(arrPositions(lngNote) + NOTE_WIDTH / 2, .Height * 0.6), RGB(0, 0, 0), BF
                    Case Else
                        picPiano(lngChn).Line (arrPositions(lngNote), 0)-(arrPositions(lngNote) + arrNoteWidths(lngNote Mod 12) * NOTE_WIDTH, .Height), RGB(0, 0, 0), B
                End Select
            Next lngNote
        End With
    Next lngChn


End Sub

Function drawNoteOnOff(ByVal lngChannel As Long, ByVal lngNote As Long, ByVal blnNoteOn As Boolean)
    On Error Resume Next
    With picPiano(lngChannel)
        Select Case lngNote Mod 12
            Case 1, 3, 6, 8, 10
                If blnNoteOn Then
                    picPiano(lngChannel).Line (arrPositions(lngNote) - NOTE_WIDTH / 2 + 15, 15)-(arrPositions(lngNote) + NOTE_WIDTH / 2 - 15, .Height * 0.6 - 15), RGB(&HFF, &H80, &HFF), BF
                    lblNotes(lngChannel).Caption = lblNotes(lngChannel).Caption + 1
                Else
                    picPiano(lngChannel).Line (arrPositions(lngNote) - NOTE_WIDTH / 2 + 15, 15)-(arrPositions(lngNote) + NOTE_WIDTH / 2 - 15, .Height * 0.6 - 15), RGB(0, 0, 0), BF
                    If lblNotes(lngChannel).Caption > 0 Then
                        lblNotes(lngChannel).Caption = lblNotes(lngChannel).Caption - 1
                    End If
                End If
                DoEvents
                Exit Function
            Case Else
                If blnNoteOn Then
                    picPiano(lngChannel).Line (arrPositions(lngNote) + 15, 15)-(arrPositions(lngNote) + arrNoteWidths(lngNote Mod 12) * NOTE_WIDTH - 15, .Height - 15), RGB(&HFF, &H80, &HFF), BF
                    lblNotes(lngChannel).Caption = lblNotes(lngChannel).Caption + 1
                Else
                    picPiano(lngChannel).Line (arrPositions(lngNote) + 15, 15)-(arrPositions(lngNote) + arrNoteWidths(lngNote Mod 12) * NOTE_WIDTH - 15, .Height - 15), RGB(&HFF, &HFF, &HFF), BF
                    If lblNotes(lngChannel).Caption > 0 Then
                        lblNotes(lngChannel).Caption = lblNotes(lngChannel).Caption - 1
                    End If
                    'Exit Function
                End If
        End Select
    
        Select Case (lngNote - 1) Mod 12
            Case 1, 3, 6, 8, 10
                picPiano(lngChannel).Line (arrPositions(lngNote) - NOTE_WIDTH / 2, 0)-(arrPositions(lngNote) + NOTE_WIDTH / 2, .Height * 0.6), RGB(0, 0, 0), BF
            Case Else
                'picPiano(lngChannel).Line (arrPositions(lngNote - 1) + arrNoteWidths(lngNote Mod 12) * NOTE_WIDTH, 0)-(arrPositions(lngNote - 1) + arrNoteWidths(lngNote Mod 12) * NOTE_WIDTH, .Height), RGB(0, 0, 0)
        End Select
        Select Case (lngNote + 1) Mod 12
            Case 1, 3, 6, 8, 10
                picPiano(lngChannel).Line (arrPositions(lngNote + 1) - NOTE_WIDTH / 2, 0)-(arrPositions(lngNote + 1) + NOTE_WIDTH / 2, .Height * 0.6), RGB(0, 0, 0), BF
            Case Else
                'picPiano(lngChannel).Line (arrPositions(lngNote + 1), 0)-(arrPositions(lngNote + 1), .Height), RGB(0, 0, 0)
        End Select
   
    End With
    DoEvents
End Function

Function AllNotesOff(ByVal lngChannel As Long) As Long
    picPiano(lngChannel).Cls
    lblNotes(lngChannel).Caption = 0
End Function
