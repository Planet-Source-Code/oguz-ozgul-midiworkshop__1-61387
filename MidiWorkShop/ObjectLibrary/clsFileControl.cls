VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0
Option Compare Binary

Private lngFileNumber       As Long
Private lngFileSeek         As Long
Private blnFileIsOpened     As Boolean

Public Function OpenBinaryFileForReading(ByVal strFilePath As String) As MidiWorkShopExceptions
    
    On Error GoTo ExceptionHandler
    
    If Not blnFileIsOpened Then
        If FileExists(strFilePath) = EXCEPTION_NONE Then
            lngFileNumber = FreeFile
            Open strFilePath For Binary Access Read As lngFileNumber
            blnFileIsOpened = True
            lngFileSeek = 1
            OpenBinaryFileForReading = EXCEPTION_NONE
        Else
            OpenBinaryFileForReading = EXCEPTION_FILE_NOT_FOUND
        End If
    Else
        OpenBinaryFileForReading = EXCEPTION_CAN_NOT_OPEN_MULTIPLE_FILES
    End If
    Exit Function
ExceptionHandler:
    OpenBinaryFileForReading = EXCEPTION_UNEXPECTED
End Function

Public Function Read(ByVal lngByteCount As Long, ByRef arrReadInto() As Byte) As MidiWorkShopExceptions
    
    On Error GoTo ExceptionHandler
    
    If blnFileIsOpened Then
        ReDim arrReadInto(lngByteCount - 1)
        Get #lngFileNumber, lngFileSeek, arrReadInto
        lngFileSeek = lngFileSeek + lngByteCount
        Read = EXCEPTION_NONE
    Else
        Read = EXCEPTION_NO_OPENED_FILES
    End If
    Exit Function
ExceptionHandler:
    Read = EXCEPTION_UNEXPECTED
End Function

Public Function GetSeek() As Long
    GetSeek = lngFileSeek
End Function

Public Function SetSeek(ByVal lngPosition As Long) As MidiWorkShopExceptions
    
    On Error GoTo ExceptionHandler
    
    If blnFileIsOpened Then
        If lngPosition <= LOF(lngFileNumber) Then
            lngFileSeek = lngPosition
            SetSeek = EXCEPTION_NONE
        ElseIf lngPosition < 1 Then
            SetSeek = EXCEPTION_SEEK_POSITION_LESS_THAN_1
        Else
            SetSeek = EXCEPTION_SEEK_POSITION_OVER_FILE_SIZE
        End If
    Else
        SetSeek = EXCEPTION_NO_OPENED_FILES
    End If
    Exit Function
ExceptionHandler:
    SetSeek = EXCEPTION_UNEXPECTED
End Function

Public Function CloseFile() As MidiWorkShopExceptions
    If blnFileIsOpened Then
        Close lngFileNumber
        lngFileNumber = 0
        lngFileSeek = 0
        blnFileIsOpened = False
        CloseFile = EXCEPTION_NONE
    Else
        CloseFile = EXCEPTION_NO_OPENED_FILES
    End If
End Function

Public Function GetSize(ByRef lngSize) As MidiWorkShopExceptions
    If blnFileIsOpened Then
        lngSize = LOF(lngFileNumber)
        GetSize = EXCEPTION_NONE
    Else
        GetSize = EXCEPTION_NO_OPENED_FILES
    End If
End Function

Public Function FileExists(ByVal strFilePath As String) As MidiWorkShopExceptions
    Dim objFileSystemObject     As New Scripting.FileSystemObject
    
    On Error GoTo ExceptionHandler
    
    If objFileSystemObject.FileExists(strFilePath) Then
        FileExists = EXCEPTION_NONE
    Else
        FileExists = EXCEPTION_FILE_NOT_FOUND
    End If
    Set objFileSystemObject = Nothing
    Exit Function
ExceptionHandler:
    FileExists = EXCEPTION_UNEXPECTED
End Function
