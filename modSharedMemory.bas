Attribute VB_Name = "modSharedMemory"
Option Explicit

Global Const MEMMSG_RUNSTRESS As Long = &H5
Global Const MEMMSG_STOPSTRESS As Long = &H7
Global Const MEMMSG_EXIT As Long = &H11

Global Const MEMSTATUS_IDLE As Long = &HA0
Global Const MEMSTATUS_RUNNING As Long = &HA2
Global Const MEMSTATUS_EXITING As Long = &HA4

Private Const PAGE_READWRITE As Long = &H4&
Private Const FILE_MAP_ALL_ACCESS As Long = &HF001F

Global Const SHAREDMEM_NAME As String = "Local\UbeCPUStress"

Private Const SHAREDMEM_INSTANCES As Long = 512
Private Const SHAREDMEM_DATASIZE As Long = 16
Private Const SHAREDMEM_SIZE As Long = SHAREDMEM_INSTANCES * SHAREDMEM_DATASIZE

Public Type SHAREDMEM_DATA
  mProcessID As Long
  mAssignedCore As Long
  mCommand As Long
  mStatus As Long
End Type

Public Type SHAREDMEMORY_LAYOUT
  Instances(0 To (SHAREDMEM_INSTANCES - 1)) As SHAREDMEM_DATA
End Type

Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappingAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
    
Dim SharedMemHandle As Long
Dim SharedMemBase As Long

Global SharedMemOffset As Long
Global SharedMemory As SHAREDMEMORY_LAYOUT

Global ActiveClients As Long

Public Function OpenSharedMemory() As Boolean
  Dim e As Boolean
  SharedMemHandle = CreateFileMapping(INVALID_HANDLE_VALUE, 0, PAGE_READWRITE, 0, SHAREDMEM_SIZE, SHAREDMEM_NAME)
  If SharedMemHandle = 0 Then Exit Function
  If Err.LastDllError() = ERROR_ALREADY_EXISTS Then e = True
  SharedMemBase = MapViewOfFile(SharedMemHandle, FILE_MAP_ALL_ACCESS, 0, 0, 0)
  If SharedMemBase = 0 Then
    Call CloseSharedMemory
    Exit Function
  End If
  If e = False Then
    ClearSharedMemory
  Else
    Call ReadFromSharedMemory(True)
  End If
  OpenSharedMemory = True
End Function

Public Function CloseSharedMemory() As Boolean
  Dim r As Boolean
  If SharedMemBase <> 0 Then
    Call UnmapViewOfFile(SharedMemBase)
    SharedMemBase = 0
    r = True
  End If
  If SharedMemHandle <> 0 Then
    Call CloseHandle(SharedMemHandle)
    SharedMemHandle = 0
    r = True
  End If
  CloseSharedMemory = r
End Function

Public Function WriteToSharedMemory(Optional WriteAllData As Boolean = False, Optional bOffset As Long = -1) As Boolean
  If SharedMemBase = 0 Then Exit Function
  If WriteAllData = True Then
    CopyMemoryByVal SharedMemBase, SharedMemory, LenB(SharedMemory)
  Else
    Dim mAddr As Long, mOff As Long
    mOff = IIf(bOffset = -1, SharedMemOffset, bOffset)
    If (mOff < LBound(SharedMemory.Instances) Or mOff > UBound(SharedMemory.Instances)) Then Exit Function
    mAddr = (SharedMemBase + (mOff * SHAREDMEM_DATASIZE))
    CopyMemoryByVal mAddr, SharedMemory.Instances(mOff), LenB(SharedMemory.Instances(mOff))
  End If
  WriteToSharedMemory = True
End Function

Public Function ReadFromSharedMemory(Optional ReadAllData As Boolean = False, Optional bOffset As Long = -1) As Boolean
  If SharedMemBase = 0 Then Exit Function
  If ReadAllData = True Then
    CopyMemory SharedMemory, ByVal SharedMemBase, LenB(SharedMemory)
  Else
    Dim mAddr As Long, mOff As Long
    mOff = IIf(bOffset = -1, SharedMemOffset, bOffset)
    If (mOff < LBound(SharedMemory.Instances) Or mOff > UBound(SharedMemory.Instances)) Then Exit Function
    mAddr = (SharedMemBase + (mOff * SHAREDMEM_DATASIZE))
    CopyMemory SharedMemory.Instances(mOff), ByVal mAddr, LenB(SharedMemory.Instances(mOff))
  End If
  ReadFromSharedMemory = True
End Function

Public Sub ClearSharedMemoryIndex(mOffset As Long)
  If (mOffset < 0 Or mOffset > (TotalCores - 1)) Then Exit Sub
  With SharedMemory.Instances(mOffset)
    .mAssignedCore = 0
    .mCommand = 0
    .mProcessID = 0
    .mStatus = 0
  End With
  Call WriteToSharedMemory(False, mOffset)
End Sub

Private Sub ClearSharedMemory()
  ZeroMemory SharedMemory, LenB(SharedMemory)
  Call WriteToSharedMemory(True)
End Sub
