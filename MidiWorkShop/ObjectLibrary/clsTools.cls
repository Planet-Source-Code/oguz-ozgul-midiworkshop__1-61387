VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0
Option Compare Binary

Private Const CLASS_NAME    As String = "MidiWorkShop.Tools"

Private arrMultipliers(3)   As Long





Friend Function Convert4BytesToLong(ByRef arrData() As Byte, ByRef lngResult As Long) As MidiWorkShopExceptions
    lngResult = arrData(0) * 8388608# + arrData(1) * 65536# + arrData(2) * 256# + arrData(3)
    Convert4BytesToLong = EXCEPTION_NONE
End Function



Friend Function ReadVariableLengthQuantity(ByRef objFile As MidiWorkShop.FileControl, ByRef lngValue As Long) As MidiWorkShopExceptions
    Dim lngTotal        As Long
    Dim lngResult(3)    As Long
    Dim lngPos          As Long
    Dim lngIdx          As Long
    Dim arrData()       As Byte
    Dim blnReadNext     As Boolean
    blnReadNext = True
    Do While blnReadNext
        Erase arrData
        If objFile.Read(1, arrData) = EXCEPTION_NONE Then
            If arrData(0) And &H80 Then ' Check most significant bit
                lngResult(lngPos) = (arrData(0) And &H7F) ' Remove most significant bit
                lngPos = lngPos + 1
            Else
                lngResult(lngPos) = arrData(0)
                blnReadNext = False
            End If
        End If
    Loop
    For lngIdx = 0 To lngPos
        lngTotal = lngTotal + lngResult(lngIdx) * arrMultipliers(lngPos - lngIdx)
    Next lngIdx
    lngValue = lngTotal
    ReadVariableLengthQuantity = EXCEPTION_NONE
End Function


Friend Function ReadVariableLengthQuantityFromArray(ByRef arrData() As Byte, ByRef lngStart As Long, ByRef lngValue As Long) As MidiWorkShopExceptions
    Dim lngTotal        As Long
    Dim lngResult(3)    As Long
    Dim lngPos          As Long
    Dim lngIdx          As Long
    Dim blnReadNext     As Boolean
    blnReadNext = True
    Do While blnReadNext
        If arrData(lngStart) And &H80 Then ' Check most significant bit
            lngResult(lngPos) = (arrData(lngStart) And &H7F) ' Remove most significant bit
            lngPos = lngPos + 1
            lngStart = lngStart + 1
        Else
            lngResult(lngPos) = arrData(lngStart)
            lngStart = lngStart + 1
            blnReadNext = False
        End If
    Loop
    For lngIdx = 0 To lngPos
        lngTotal = lngTotal + lngResult(lngIdx) * arrMultipliers(lngPos - lngIdx)
    Next lngIdx
    lngValue = lngTotal
    ReadVariableLengthQuantityFromArray = EXCEPTION_NONE
End Function

Friend Function CopyPartOfArray(ByRef arrSource() As Byte, ByRef arrDestination() As Byte, ByRef lngStart As Long, ByVal lngCount As Long) As MidiWorkShopExceptions
    Dim lngIdx      As Long
    ReDim arrDestination(lngCount - 1)
    For lngIdx = lngStart To lngStart + lngCount - 1
        If lngIdx <= UBound(arrSource) Then
            arrDestination(lngIdx - lngStart) = arrSource(lngIdx)
        End If
    Next lngIdx
    lngStart = lngStart + lngCount
    CopyPartOfArray = EXCEPTION_NONE
End Function


Friend Function CompareBinaryArrayToString(ByVal strText As String, ByRef arrData() As Byte) As MidiWorkShopExceptions
    Const PROCEDURE_NAME            As String = "CompareBinaryArrayToString"
    Dim objExceptionHandler         As New MidiWorkShop.ExceptionHandler
    Dim lngIdx                      As Long
    
    On Error GoTo ExceptionHandler
    
    If UBound(arrData) - LBound(arrData) + 1 = LenB(strText) Then
        For lngIdx = LBound(arrData) To UBound(arrData)
            If arrData(lngIdx) <> AscB(MidB(strText, lngIdx + 1, 1)) Then
                CompareBinaryArrayToString = EXCEPTION_COMPARE_NOT_MATCH
                Exit Function
            End If
        Next lngIdx
        CompareBinaryArrayToString = EXCEPTION_NONE
    Else
        CompareBinaryArrayToString = EXCEPTION_COMPARE_NOT_MATCH
    End If
    Exit Function
ExceptionHandler:
    CompareBinaryArrayToString = EXCEPTION_UNEXPECTED
    Call objExceptionHandler.LogException(Err.Number, Err.Source, Err.Description, Erl, CLASS_NAME, PROCEDURE_NAME, "")
End Function

Private Sub Class_Initialize()
    arrMultipliers(0) = 1 '         first 7 bits        stay the same
    arrMultipliers(1) = 128 '       second 7 bits are   shift left  7 bits = 2^7
    arrMultipliers(2) = 32768 '     third 7 bits are    shift left 14 bits = 2^14
    arrMultipliers(3) = 4194304 '   fourth 7 bits are   shift left 21 bits = 2^21
End Sub
