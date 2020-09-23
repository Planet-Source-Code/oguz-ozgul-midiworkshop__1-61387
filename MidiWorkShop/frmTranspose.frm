VERSION 5.00
Begin VB.Form frmTranspose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtTranpose 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   90
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Transpose"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmTranspose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objMidi As MidiFile

Private Sub cmdApply_Click()
    If IsNumeric(txtTranpose.Text) Then
        mdiMain.AllNotesOff
        objMidi.SetTranspose txtTranpose.Text
        Unload Me
    Else
        txtTranpose.SetFocus
    End If
End Sub

Private Sub Form_Load()
    txtTranpose = objMidi.Transpose
End Sub

Function SetMidiFile(ByRef objMidiFile As MidiFile) As Long
    Set objMidi = objMidiFile
End Function
