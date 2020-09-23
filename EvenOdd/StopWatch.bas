Attribute VB_Name = "StopWatch"
Option Explicit

' -----------------------------------------------------------------------------
'
' A small module used for timing high resolution pieces of code. Most of this
' code was botched together from the example by the KPD team at mentalis.org
' (allapi.net)
'
' -----------------------------------------------------------------------------

Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type


Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Private m_frequency     As Currency
Private m_start         As LARGE_INTEGER
Private m_stop          As LARGE_INTEGER



Public Function CheckSupport() As Boolean

  Dim liFrequency As LARGE_INTEGER
  
    CheckSupport = QueryPerformanceFrequency(liFrequency) > 0
    m_frequency = LargeIntToCurrency(liFrequency)
End Function


Public Sub StartTimer()
    QueryPerformanceCounter m_start
End Sub

Public Sub StopTimer()
    QueryPerformanceCounter m_stop
End Sub


Public Function GetTime() As Currency
    GetTime = (LargeIntToCurrency(m_stop) - LargeIntToCurrency(m_start)) / m_frequency
End Function


Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
    CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
    LargeIntToCurrency = LargeIntToCurrency * 10000
End Function



