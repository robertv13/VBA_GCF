VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' YouTube Video: Make Your VBA Code Run 1000 Times Faster (Part 2)
' Link: https://youtu.be/RNqd89K_bbU
' Website: https://excelmacromastery.com/

' MicroTimer taken from Charles Williams website: https://bit.ly/2Yz4br3

#If VBA7 Then
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias _
"QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias _
"QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
Private Declare Function getFrequency Lib "kernel32" Alias _
"QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare Function getTickCount Lib "kernel32" Alias _
"QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If

Private m_Time As Double
Private m_lastTime As Double

Private m_fromStartTime As Double
Private m_fromLastTime As Double

Enum eResultType
    efromstart
    efromlast
End Enum

Public Property Get fromStartTime() As Double
    Call GetTime
    fromStartTime = m_fromStartTime
End Property

Public Property Get fromLastTime() As Double
    Call GetTime
    fromLastTime = m_fromLastTime
End Property

Private Function MicroTimer() As Double
    '
    ' returns seconds
    ' uses Windows API calls to the high resolution timer
    '
    Dim cyTicks1 As Currency
    Dim cyTicks2 As Currency
    Static cyFrequency As Currency
    '
    MicroTimer = 0
    '
    ' get frequency
    '
    If cyFrequency = 0 Then getFrequency cyFrequency
    '
    ' get ticks
    '
    getTickCount cyTicks1
    getTickCount cyTicks2
    If cyTicks2 < cyTicks1 Then cyTicks2 = cyTicks1
    '
    ' calc seconds
    '
    If cyFrequency Then MicroTimer = cyTicks2 / cyFrequency
End Function
    
Public Sub start()
    m_Time = MicroTimer
    m_lastTime = m_Time
End Sub
Public Function PrintTime(Optional ByVal text As String, Optional resultType As eResultType = efromlast)

    'Debug.Print GetTime(resultType) * 1000 & "," & text
    Call GetTime
    
    If resultType = efromstart Then
        Debug.Print m_fromStartTime * 1000 & "," & text
    Else
        Debug.Print m_fromLastTime * 1000 & "," & text
    End If
    
End Function

Private Function GetTime()
    
    ' Get time
    Dim currentTime As Double
    currentTime = MicroTimer
    
    m_fromStartTime = (currentTime - m_Time)
    m_fromLastTime = (currentTime - m_lastTime)
    
    m_lastTime = currentTime
    
End Function

