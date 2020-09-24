Attribute VB_Name = "modMain"
Option Explicit

'API
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount& Lib "kernel32" ()

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


Public Sub WaitAsync(msTime As Long)
    Dim FirstTick As Long
    FirstTick = GetTickCount
    While (GetTickCount - FirstTick) < msTime
        DoEvents
    Wend
End Sub


Public Function getRnd(rMax As Long) As Long
    getRnd = (Rnd * rMax) \ 1
End Function
