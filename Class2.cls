VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComputerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class       : CComputerInfo
' Description : This class implements routines for obtaining computer information.
' Source      : Total VB SourceBook 6

Private Declare Sub GetSystemInfo _
  Lib "kernel32" _
  (lpSystemInfo As SYSTEM_INFO)

Private Declare Function IsProcessorFeaturePresent _
  Lib "kernel32" _
  (ByVal ProcessorFeature As Long) _
  As Long
  
Private Declare Function GetSystemMetrics _
  Lib "user32" _
  (ByVal nIndex As Long) _
  As Long
  
Private Type SYSTEM_INFO
  wProcessorArchitecture As Integer
  wReserved As Integer
  dwPageSize As Long
  lpMinimumApplicationAddress As Long
  lpMaximumApplicationAddress As Long
  dwActiveProcessorMask As Long
  dwNumberOfProcessors As Long
  dwProcessorType As Long
  dwAllocationGranularity As Long
  wProcessorLevel As Integer
  wProcessorRevision As Integer
End Type

Private Const SM_SLOWMACHINE = 73

Private Const PF_FLOATING_POINT_PRECISION_ERRATA = 0
Private Const PF_FLOATING_POINT_EMULATED = 1
Private Const PF_COMPARE_EXCHANGE_DOUBLE = 2
Private Const PF_MMX_INSTRUCTIONS_AVAILABLE = 3

Public Enum EnumProcessorType
  cmiIntel386 = 386
  cmiIntel486 = 486
  cmiIntelPENTIUM = 586
  cmiMIPSR4000 = 4000
  cmiALPHA21064 = 21064
End Enum

Public Enum EnumProcessorArchitecture
  cmiIntel = 0
  cmiMIPS = 1
  cmiALPHA = 2
  cmiPPC = 3
  cmiUnknown = &HFFFF
End Enum

Property Get ActiveProcessorMask() As Long
  ' Returns : The active processors in the system
  ' Source: Total VB SourceBook 6
  Dim si As SYSTEM_INFO

  On Error GoTo PROC_ERR

  GetSystemInfo si
  
  ActiveProcessorMask = si.dwActiveProcessorMask

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ActiveProcessorMask"
  Resume PROC_EXIT

End Property

Property Get AllocationGranularity() As Long
  ' Returns : The granularity with which virtual memory is allocated.
  ' Source: Total VB SourceBook 6
  Dim si As SYSTEM_INFO

  On Error GoTo PROC_ERR

  GetSystemInfo si
  
  AllocationGranularity = si.dwAllocationGranularity

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "AllocationGranularity"
  Resume PROC_EXIT

End Property

Property Get CompareExchangeDouble() As Boolean
  ' Returns : True if the compare and exchange double operation is available
  '           False if it is not
  ' Source: Total VB SourceBook 6
  
  On Error GoTo PROC_ERR

  CompareExchangeDouble = _
    IsProcessorFeaturePresent(PF_COMPARE_EXCHANGE_DOUBLE)

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "CompareExchangeDouble"
  Resume PROC_EXIT

End Property

Property Get FloatingPointEmulated() As Boolean
  ' Returns : True if floating point emulation is used, False if it is not
  ' Source: Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  FloatingPointEmulated = _
    IsProcessorFeaturePresent(PF_FLOATING_POINT_EMULATED)

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "FloatingPointEmulated"
  Resume PROC_EXIT

End Property

Property Get FloatingPointError() As Boolean
  ' Returns : True if the pentium floating point bug exists in this processor
  '           False if it does not
  ' Source: Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  FloatingPointError = _
    IsProcessorFeaturePresent(PF_FLOATING_POINT_PRECISION_ERRATA)

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "FloatingPointError"
  Resume PROC_EXIT

End Property

Property Get LowMemory() As Boolean
  ' Returns   : True if the computer is considered a low memory machine, False
  '             if it is not
  ' Source: Total VB SourceBook 6
  On Error GoTo PROC_ERR

  ' The following flags are defined for the return value of the
  ' GetSystemMetrics(SM_SLOWMACHINE) function
  '
  ' &H0001 - CPU is a 386
  ' &H0002 - low memory machine (less than 5 megabytes)
  ' &H0004 - slow (non-accelerated) display card

  LowMemory = (GetSystemMetrics(SM_SLOWMACHINE) And 2) > 0

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LowMemory"
  Resume PROC_EXIT

End Property

Property Get MaxAppAddress() As Long
  ' Returns : the highest memory address accessible to applications and DLLs.
  ' Source: Total VB SourceBook 6
  Dim si As SYSTEM_INFO

  On Error GoTo PROC_ERR

  GetSystemInfo si
  
  MaxAppAddress = si.lpMaximumApplicationAddress

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "MaxAppAddress"
  Resume PROC_EXIT

End Property

Property Get MinAppAddress() As Long
  ' Returns   : the lowest memory address accessible to applications and DLLs.
  ' Source: Total VB SourceBook 6
  Dim si As SYSTEM_INFO

  On Error GoTo PROC_ERR
  
  GetSystemInfo si
  
  MinAppAddress = si.lpMinimumApplicationAddress

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "MinAppAddress"
  Resume PROC_EXIT

End Property

Property Get MMXAvailable() As Boolean
  ' Returns : True if the processor supports MMX, False if it does not
  ' Source: Total VB SourceBook 6
  On Error GoTo PROC_ERR

  MMXAvailable = _
    IsProcessorFeaturePresent(PF_MMX_INSTRUCTIONS_AVAILABLE)

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "MMXAvailable"
  Resume PROC_EXIT

End Property

Property Get NumberOfProcessors() As Long
  ' Returns : The number of processors in the system
  ' Source: Total VB SourceBook 6
  Dim si As SYSTEM_INFO

  On Error GoTo PROC_ERR

  GetSystemInfo si
  
  NumberOfProcessors = si.dwNumberOfProcessors

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "NumberOfProcessors"
  Resume PROC_EXIT

End Property

Property Get PageSize() As Long
  ' Returns : Indicate the page size.
  ' Source: Total VB SourceBook 6
  '
  Dim si As SYSTEM_INFO

  On Error GoTo PROC_ERR
  
  GetSystemInfo si
  
  PageSize = si.dwPageSize

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "PageSize"
  Resume PROC_EXIT

End Property

Property Get ProcessorArchitecture() As EnumProcessorArchitecture
  ' Returns : The processor architecture
  ' Source: Total VB SourceBook 6
  Dim si As SYSTEM_INFO

  On Error GoTo PROC_ERR

  GetSystemInfo si
  
  ProcessorArchitecture = si.wProcessorArchitecture

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ProcessorArchitecture"
  Resume PROC_EXIT

End Property

Property Get ProcessorLevel() As Long
  ' Returns : The system's architecture-dependent processor level
  ' Source: Total VB SourceBook 6
  Dim si As SYSTEM_INFO

  On Error GoTo PROC_ERR

  GetSystemInfo si
  
  ProcessorLevel = si.wProcessorLevel

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ProcessorLevel"
  Resume PROC_EXIT

End Property

Property Get ProcessorRevision() As Long
  ' Returns : The architecture-dependent processor revision.
  ' Source: Total VB SourceBook 6
  Dim si As SYSTEM_INFO

  On Error GoTo PROC_ERR

  GetSystemInfo si
  
  ProcessorRevision = si.wProcessorRevision

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ProcessorRevision"
  Resume PROC_EXIT

End Property

Property Get ProcessorType() As EnumProcessorType
  ' Returns : The processor type
  ' Source: Total VB SourceBook 6
  Dim si As SYSTEM_INFO

  On Error GoTo PROC_ERR
  
  GetSystemInfo si
  
  ProcessorType = si.dwProcessorType

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ProcessorType"
  Resume PROC_EXIT

End Property

Property Get SlowGraphics() As Boolean
  ' Returns   : True if the graphics are considered slow, False if they are not
  ' Source: Total VB SourceBook 6
  On Error GoTo PROC_ERR

  ' The following flags are defined for the return value of the
  ' GetSystemMetrics(SM_SLOWMACHINE) function
  '
  ' &H0001 - CPU is a 386
  ' &H0002 - low memory machine (less than 5 megabytes)
  ' &H0004 - slow (non-accelerated) display card

  SlowGraphics = (GetSystemMetrics(SM_SLOWMACHINE) And 4) > 0

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "SlowGraphics"
  Resume PROC_EXIT

End Property

Property Get SlowMachine() As Boolean
  ' Returns : True if the computer is considered a slow machine, False if it
  '           is not
  ' Source: Total VB SourceBook 6
  On Error GoTo PROC_ERR

  ' The following flags are defined for the return value of the
  ' GetSystemMetrics(SM_SLOWMACHINE) function
  '
  ' &H0001 - CPU is a 386
  ' &H0002 - low memory machine (less than 5 megabytes)
  ' &H0004 - slow (non-accelerated) display card
  
  SlowMachine = (GetSystemMetrics(SM_SLOWMACHINE) And 1) > 0

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "SlowMachine"
  Resume PROC_EXIT

End Property


