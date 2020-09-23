Attribute VB_Name = "modMidiException"
Option Explicit
Option Base 0
Option Compare Binary

Public Const MAXERRORLENGTH = 128  '  max error text length (including final NULL)

Public Const MMSYSERR_BASE = 0

' General error return values
Public Const MMSYSERR_NOERROR = 0  '  no error
Public Const MMSYSERR_ERROR = (MMSYSERR_BASE + 1)  '  unspecified error
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)  '  device ID out of range
Public Const MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)  '  driver failed enable
Public Const MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)  '  device already allocated
Public Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)  '  device handle is invalid
Public Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)  '  no device driver present
Public Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)  '  memory allocation error
Public Const MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8)  '  function isn't supported
Public Const MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)  '  error value out of range
Public Const MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10) '  invalid flag passed
Public Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11) '  invalid parameter passed
Public Const MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12) '  handle being used simultaneously on another thread (eg callback)
Public Const MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13) '  "Specified alias not found in WIN.INI
Public Const MMSYSERR_LASTERROR = (MMSYSERR_BASE + 13) '  last error in range

' GET EXCEPTION
Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long


Public Function GetMidiExceptionMessage(ByVal lngException As Long) As String
    Dim strException As String * MAXERRORLENGTH
    
    midiOutGetErrorText lngException, strException, MAXERRORLENGTH

    If InStr(strException, Chr(0)) Then
        GetMidiExceptionMessage = Left(strException, InStr(strException, Chr(0)) - 1)
    Else
        GetMidiExceptionMessage = strException
    End If

End Function


