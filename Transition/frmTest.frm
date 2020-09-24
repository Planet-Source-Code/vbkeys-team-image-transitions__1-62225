VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Transition"
   ClientHeight    =   7425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9165
      MouseIcon       =   "frmTest.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmTest.frx":0CCA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   30
      Width           =   240
   End
   Begin TransitionTest.vbkTitle vbkTitle1 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   529
      Picture         =   "frmTest.frx":10E7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtThickness 
      Height          =   315
      Left            =   3615
      TabIndex        =   5
      Text            =   "1"
      Top             =   6990
      Width           =   525
   End
   Begin VB.ComboBox cboTrans 
      Height          =   315
      ItemData        =   "frmTest.frx":A50B
      Left            =   60
      List            =   "frmTest.frx":A536
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6990
      Width           =   2430
   End
   Begin VB.PictureBox picB 
      AutoRedraw      =   -1  'True
      Height          =   6810
      Left            =   4515
      ScaleHeight     =   6750
      ScaleWidth      =   7320
      TabIndex        =   3
      Top             =   8205
      Visible         =   0   'False
      Width           =   7380
   End
   Begin VB.PictureBox picA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6555
      Left            =   60
      ScaleHeight     =   6525
      ScaleWidth      =   9300
      TabIndex        =   2
      Top             =   330
      Width           =   9330
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   345
      Left            =   4305
      TabIndex        =   1
      Top             =   6975
      Width           =   1320
   End
   Begin TransitionTest.vbkTransition vbkTransition1 
      Height          =   645
      Left            =   6855
      TabIndex        =   0
      Top             =   8175
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1138
      TranType        =   12
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H009A3500&
      Height          =   1020
      Left            =   0
      Top             =   0
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Thickness:"
      Height          =   210
      Left            =   2700
      TabIndex        =   6
      Top             =   7065
      Width           =   855
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copryright VBKeys.com
Option Explicit

Private Sub cmdStart_Click()
    cmdStart.Enabled = False
    cboTrans.Enabled = False
    txtThickness.Enabled = False
    vbkTransition1.TranType = cboTrans.ListIndex
    vbkTransition1.Thickness = Val(txtThickness.Text)
    vbkTransition1.Start
End Sub

Private Sub Form_Load()
    Shape1.Move 0, 0, Me.Width, Me.Height
    Set picA.Picture = LoadPicture(App.Path & "\images\picA.jpg")
    Set picB.Picture = LoadPicture(App.Path & "\images\picb.jpg")
    picB.Width = picA.Width
    picB.Height = picA.Height
    Set vbkTransition1.objPic(A) = picA
    Set vbkTransition1.objPic(B) = picB
    cboTrans.ListIndex = 0
End Sub

Private Sub picClose_Click()
    Unload Me
End Sub

Private Sub vbkTransition1_EndTran()
    If cmdStart.Tag = "1" Then
        Set picA.Picture = LoadPicture(App.Path & "\images\picA.jpg")
        Set picB.Picture = LoadPicture(App.Path & "\images\picb.jpg")
        picA.Refresh
        picB.Refresh
        'Set vbkTransition1.objPic(A) = picA
        'Set vbkTransition1.objPic(B) = picB
        cmdStart.Tag = ""
    Else
        Set picA.Picture = LoadPicture(App.Path & "\images\picb.jpg")
        Set picB.Picture = LoadPicture(App.Path & "\images\pica.jpg")
        picA.Refresh
        picB.Refresh
    
        'Set vbkTransition1.objPic(B) = picA
        'Set vbkTransition1.objPic(A) = picB
        cmdStart.Tag = "1"
    End If
    Set vbkTransition1.objPic(A) = picA
    Set vbkTransition1.objPic(B) = picB
    cmdStart.Enabled = True
    cboTrans.Enabled = True
    txtThickness.Enabled = True
End Sub
