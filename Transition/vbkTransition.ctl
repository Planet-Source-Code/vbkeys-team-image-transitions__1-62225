VERSION 5.00
Begin VB.UserControl vbkTransition 
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   552
   Begin VB.PictureBox picB 
      AutoRedraw      =   -1  'True
      Height          =   5670
      Left            =   4125
      ScaleHeight     =   374
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   376
      TabIndex        =   0
      Top             =   3465
      Visible         =   0   'False
      Width           =   5700
   End
End
Attribute VB_Name = "vbkTransition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum enmTransition
    trHRandomLine
    trVRandomLine
    trCrossRandomLine
    trRandomPoint
    trVWipeCenterStep
    trHWipeCenterStep
    trWipeCenterStep
    trvWipeCenter
    trHWipeCenter
    trWipeCenter
    trCenter2Border
    trWindowOpen
    trWindowOpenUp
End Enum

Public Enum enmPic
    A = 1
    B = 2
End Enum

Dim mHDCSrc As Long
Dim mHDCDest  As Long

'Default Property Values:
Const m_def_Thickness = 1
Const m_def_Delay = 1
Const m_def_TranType = 1
'Property Variables:
Dim m_Thickness As Long
Dim m_Delay As Long
Dim m_TranType As enmTransition
Dim mW As Long, mH As Long

Event EndTran()

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1
Public Property Get TranType() As enmTransition
    TranType = m_TranType
End Property

Public Property Let TranType(ByVal New_TranType As enmTransition)
    m_TranType = New_TranType
    PropertyChanged "TranType"
End Property


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_TranType = m_def_TranType
    m_Delay = m_def_Delay
    m_Thickness = m_def_Thickness
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_TranType = PropBag.ReadProperty("TranType", m_def_TranType)
    m_Delay = PropBag.ReadProperty("Delay", m_def_Delay)
    m_Thickness = PropBag.ReadProperty("Thickness", m_def_Thickness)
End Sub

Private Sub UserControl_Resize()
    picB.Width = UserControl.Width
    picB.Height = UserControl.Height
    mW = UserControl.ScaleWidth
    mH = UserControl.ScaleHeight
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("TranType", m_TranType, m_def_TranType)
    Call PropBag.WriteProperty("Delay", m_Delay, m_def_Delay)
    Call PropBag.WriteProperty("Thickness", m_Thickness, m_def_Thickness)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,500
Public Property Get Delay() As Long
    Delay = m_Delay
End Property

Public Property Let Delay(ByVal New_Delay As Long)
    m_Delay = New_Delay
    PropertyChanged "Delay"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1
Public Property Get Thickness() As Long
    Thickness = m_Thickness
End Property

Public Property Let Thickness(ByVal New_Thickness As Long)
    m_Thickness = New_Thickness
    PropertyChanged "Thickness"
End Property


Public Property Set objPic(Pic As enmPic, ByVal Src As Object)
    If Pic = A Then
        mW = Src.Width \ 15
        mH = Src.Height \ 15
        mHDCDest = Src.hDC
    Else
        mHDCSrc = Src.hDC
    End If
End Property

Public Sub Start()
    Randomize
    Select Case m_TranType
        Case trHRandomLine
            dotrHRandomLine
        Case trVRandomLine
            dotrVRandomLine
        Case trCrossRandomLine
            dotrCrossRandomLine
        Case trRandomPoint
            dotrRandomPoint
        Case trVWipeCenterStep
            dovWipeCenterStep
        Case trHWipeCenterStep
            dohWipeCenterStep
        Case trWipeCenterStep
            doWipeCenterStep
        Case trHWipeCenter
            dohWipeCenter
        Case trvWipeCenter
            dovWipeCenter
        Case trWipeCenter
            doWipeCenter
        Case trCenter2Border
            doCenter2Border
        Case trWindowOpen
            doWindowOpen
        Case trWindowOpenUp
            doWindowOpenUp
    End Select
End Sub

Private Sub dotrHRandomLine()
    Dim I As Long
    Dim Y As Long
    For I = 0 To mH * 1.2
        Y = getRnd(mH)
        BitBlt mHDCDest, 0, Y, mW, m_Thickness, mHDCSrc, 0, Y, SRCCOPY
        WaitAsync m_Delay
        DoEvents
    Next I
    BitBlt mHDCDest, 0, 0, mW, mH, mHDCSrc, 0, 0, SRCCOPY
    RaiseEvent EndTran
End Sub

Private Sub dotrVRandomLine()
    Dim I As Long
    Dim X As Long
    For I = 0 To mH * 1.2
        X = getRnd(mW)
        BitBlt mHDCDest, X, 0, m_Thickness, mH, mHDCSrc, X, 0, SRCCOPY
        WaitAsync m_Delay
        DoEvents
    Next I
    BitBlt mHDCDest, 0, 0, mW, mH, mHDCSrc, 0, 0, SRCCOPY
    RaiseEvent EndTran
End Sub

Private Sub dotrCrossRandomLine()
    Dim I As Long
    Dim X As Long, Y As Long
    For I = 0 To mW * 1.2
        X = getRnd(mW)
        Y = getRnd(mH)
        BitBlt mHDCDest, X, 0, m_Thickness, mH, mHDCSrc, X, 0, SRCCOPY
        BitBlt mHDCDest, 0, Y, mW, m_Thickness, mHDCSrc, 0, Y, SRCCOPY
        WaitAsync m_Delay
        DoEvents
    Next I
    BitBlt mHDCDest, 0, 0, mW, mH, mHDCSrc, 0, 0, SRCCOPY
    RaiseEvent EndTran
End Sub

Private Sub dotrRandomPoint()
    Dim I As Long
    Dim X As Long, Y As Long
    For I = 0 To mW * mH * 1.2
        X = getRnd(mW)
        Y = getRnd(mH)
        BitBlt mHDCDest, X, Y, m_Thickness, m_Thickness, mHDCSrc, X, Y, SRCCOPY
        DoEvents
    Next I
    BitBlt mHDCDest, 0, 0, mW, mH, mHDCSrc, 0, 0, SRCCOPY
    RaiseEvent EndTran
End Sub

Private Sub dovWipeCenterStep()
    Dim I As Long
    Dim ToAd As Long
    ToAd = IIf(Int(mW / m_Thickness) Mod 2 = 1, 0, m_Thickness)
    
    For I = 0 To Int(mW / m_Thickness) Step 2
        BitBlt mHDCDest, I * m_Thickness, 0, m_Thickness, mH, mHDCSrc, I * m_Thickness, 0, SRCCOPY
        BitBlt mHDCDest, mW - I * m_Thickness + ToAd, 0, m_Thickness, mH, mHDCSrc, mW - I * m_Thickness + ToAd, 0, SRCCOPY
        WaitAsync m_Delay
        DoEvents
    Next I
    BitBlt mHDCDest, 0, 0, mW, mH, mHDCSrc, 0, 0, SRCCOPY
    RaiseEvent EndTran
End Sub


Private Sub dohWipeCenterStep()
    Dim I As Long
    Dim ToAd As Long
    
    ToAd = IIf(Int(mH / m_Thickness) Mod 2 = 1, 0, m_Thickness)
    
    For I = 0 To Int(mH / m_Thickness) Step 2
        BitBlt mHDCDest, 0, I * m_Thickness, mW, m_Thickness, mHDCSrc, 0, I * m_Thickness, SRCCOPY
        BitBlt mHDCDest, 0, mH - I * m_Thickness + ToAd, mW, m_Thickness, mHDCSrc, 0, mH - I * m_Thickness + ToAd, SRCCOPY
        WaitAsync m_Delay
        DoEvents
    Next I
    BitBlt mHDCDest, 0, 0, mW, mH, mHDCSrc, 0, 0, SRCCOPY
    RaiseEvent EndTran
End Sub

Private Sub doWipeCenterStep()
    Dim I As Long
    Dim ToAdx As Long
    Dim ToAdy As Long
    
    ToAdx = IIf(Int(mH / m_Thickness) Mod 2 = 1, 0, m_Thickness)
    ToAdy = IIf(Int(mW / m_Thickness) Mod 2 = 1, 0, m_Thickness)
    
    For I = 0 To Int(mH / m_Thickness) Step 2
        BitBlt mHDCDest, 0, I * m_Thickness, mW, m_Thickness, mHDCSrc, 0, I * m_Thickness, SRCCOPY
        BitBlt mHDCDest, 0, mH - I * m_Thickness + ToAdx, mW, m_Thickness, mHDCSrc, 0, mH - I * m_Thickness + ToAdx, SRCCOPY
        
        BitBlt mHDCDest, I * m_Thickness, 0, m_Thickness, mH, mHDCSrc, I * m_Thickness, 0, SRCCOPY
        BitBlt mHDCDest, mW - I * m_Thickness + ToAdy, 0, m_Thickness, mH, mHDCSrc, mW - I * m_Thickness + ToAdy, 0, SRCCOPY
        
        WaitAsync m_Delay
        DoEvents
    Next I
    BitBlt mHDCDest, 0, 0, mW, mH, mHDCSrc, 0, 0, SRCCOPY
    RaiseEvent EndTran
End Sub

Private Sub dovWipeCenter()
    Dim I As Long
        
    For I = 0 To Int(mW / m_Thickness) \ 2
        BitBlt mHDCDest, I * m_Thickness, 0, m_Thickness, mH, mHDCSrc, I * m_Thickness, 0, SRCCOPY
        BitBlt mHDCDest, mW - I * m_Thickness, 0, m_Thickness, mH, mHDCSrc, mW - I * m_Thickness, 0, SRCCOPY
        WaitAsync m_Delay
        DoEvents
    Next I
    BitBlt mHDCDest, 0, 0, mW, mH, mHDCSrc, 0, 0, SRCCOPY
    RaiseEvent EndTran
End Sub

Private Sub dohWipeCenter()
    Dim I As Long
    
    For I = 0 To Int(mH / m_Thickness) \ 2
        BitBlt mHDCDest, 0, I * m_Thickness, mW, m_Thickness, mHDCSrc, 0, I * m_Thickness, SRCCOPY
        BitBlt mHDCDest, 0, mH - I * m_Thickness, mW, m_Thickness, mHDCSrc, 0, mH - I * m_Thickness, SRCCOPY
        WaitAsync m_Delay
        DoEvents
    Next I
    BitBlt mHDCDest, 0, 0, mW, mH, mHDCSrc, 0, 0, SRCCOPY
    RaiseEvent EndTran
End Sub

Private Sub doWipeCenter()
    Dim I As Long
    Dim MinS As Long
    MinS = mH
    If mW < mH Then MinS = mW
    For I = 0 To Int(MinS / m_Thickness) \ 2
        BitBlt mHDCDest, 0, I * m_Thickness, mW, m_Thickness, mHDCSrc, 0, I * m_Thickness, SRCCOPY
        BitBlt mHDCDest, 0, mH - I * m_Thickness, mW, m_Thickness, mHDCSrc, 0, mH - I * m_Thickness, SRCCOPY
        
        BitBlt mHDCDest, I * m_Thickness, 0, m_Thickness, mH, mHDCSrc, I * m_Thickness, 0, SRCCOPY
        BitBlt mHDCDest, mW - I * m_Thickness, 0, m_Thickness, mH, mHDCSrc, mW - I * m_Thickness, 0, SRCCOPY
        
        WaitAsync m_Delay
        DoEvents
    Next I
    BitBlt mHDCDest, 0, 0, mW, mH, mHDCSrc, 0, 0, SRCCOPY
    RaiseEvent EndTran
End Sub

Private Sub doCenter2Border()
    Dim TheLimit As Long
    Dim I As Long
    TheLimit = mW
    If mH > mW Then TheLimit = mH
    For I = Int(TheLimit / 2) To 0 Step -1
        BitBlt mHDCDest, 0, I, mW, 1, mHDCSrc, 0, I, SRCCOPY
        BitBlt mHDCDest, 0, mH - I, mW, 1, mHDCSrc, 0, mH - I, SRCCOPY
        BitBlt mHDCDest, I, 0, 1, mH, mHDCSrc, I, 0, SRCCOPY
        BitBlt mHDCDest, mW - I, 0, 1, mH, mHDCSrc, mW - I, 0, SRCCOPY
        
        WaitAsync m_Delay
        DoEvents
    Next I
    BitBlt mHDCDest, 0, 0, mW, mH, mHDCSrc, 0, 0, SRCCOPY
    RaiseEvent EndTran
End Sub

Private Sub doWindowOpen()
    Dim I As Long
    
    For I = Int(mW / (2 * m_Thickness)) To 0 Step -1
        BitBlt mHDCDest, I * m_Thickness, 0, m_Thickness, mH, mHDCSrc, I * m_Thickness, 0, SRCCOPY
        BitBlt mHDCDest, mW - I * m_Thickness, 0, m_Thickness, mH, mHDCSrc, mW - I * m_Thickness, 0, SRCCOPY
        
        WaitAsync m_Delay
        DoEvents
    Next I
    BitBlt mHDCDest, 0, 0, mW, mH, mHDCSrc, 0, 0, SRCCOPY
    RaiseEvent EndTran
End Sub

Private Sub doWindowOpenUp()
    Dim I As Long
    For I = Int(mH / (2 * m_Thickness)) To 0 Step -1
        BitBlt mHDCDest, 0, I * m_Thickness, mW, m_Thickness, mHDCSrc, 0, I * m_Thickness, SRCCOPY
        BitBlt mHDCDest, 0, mH - I * m_Thickness, mW, m_Thickness, mHDCSrc, 0, mH - I * m_Thickness, SRCCOPY
        
        WaitAsync m_Delay
        DoEvents
    Next I
    BitBlt mHDCDest, 0, 0, mW, mH, mHDCSrc, 0, 0, SRCCOPY
    RaiseEvent EndTran
End Sub
