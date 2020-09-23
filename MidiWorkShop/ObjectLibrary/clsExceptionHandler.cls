VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExceptionHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0
Option Compare Binary

Public Enum MidiWorkShopExceptions
    EXCEPTION_NONE = 0
    EXCEPTION_UNEXPECTED = 1
    EXCEPTION_FILE_NOT_FOUND = 2
    EXCEPTION_CAN_NOT_OPEN_MULTIPLE_FILES = 3
    EXCEPTION_NO_OPENED_FILES = 4
    EXCEPTION_SEEK_POSITION_OVER_FILE_SIZE = 5
    EXCEPTION_SEEK_POSITION_LESS_THAN_1 = 6
    EXCEPTION_NOT_A_VALID_MIDI_FILE = 7
    EXCEPTION_COMPARE_NOT_MATCH = 8
    EXCEPTION_MIDI_FILE_TYPE_NOT_SUPPORTED = 9
    



End Enum

Public Function LogException(ByVal lngExceptionNumber As Long, _
                             ByVal strExceptionSource As String, _
                             ByVal strExceptionMessage As String, _
                             ByVal lngExceptionLine As Long, _
                             ByVal strExceptionClass As String, _
                             ByVal strExceptionProcedure As String, _
                             ByVal strExceptionExtraInfo As String) As Long

    




End Function
