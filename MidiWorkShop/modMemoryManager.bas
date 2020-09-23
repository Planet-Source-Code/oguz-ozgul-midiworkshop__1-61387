Attribute VB_Name = "modMemoryManager"
Option Explicit

Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalHandle Lib "kernel32" (wMem As Any) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal wBytes As Long, ByVal wFlags As Long) As Long
Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

' Local Memory Flags
Public Const LMEM_FIXED = &H0
Public Const LMEM_MOVEABLE = &H2
Public Const LMEM_NOCOMPACT = &H10
Public Const LMEM_NODISCARD = &H20
Public Const LMEM_ZEROINIT = &H40
Public Const LMEM_MODIFY = &H80
Public Const LMEM_DISCARDABLE = &HF00
Public Const LMEM_VALID_FLAGS = &HF72
Public Const LMEM_INVALID_HANDLE = &H8000

Public Const LHND = (LMEM_MOVEABLE + LMEM_ZEROINIT)
Public Const LPTR = (LMEM_FIXED + LMEM_ZEROINIT)

Public Const NONZEROLHND = (LMEM_MOVEABLE)
Public Const NONZEROLPTR = (LMEM_FIXED)

' Flags returned by LocalFlags (in addition to LMEM_DISCARDABLE)
Public Const LMEM_DISCARDED = &H4000
Public Const LMEM_LOCKCOUNT = &HFF


