VERSION 5.00
Begin VB.UserControl AniGif 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   573
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   0
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.PictureBox Picture7 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   -15
      ScaleHeight     =   175
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   204
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   3060
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7980
      Top             =   2880
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   5745
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   184
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   2760
      Begin VB.PictureBox Image1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Index           =   0
         Left            =   0
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   129
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   3015
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   1
      Top             =   3540
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   0
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   0
      Top             =   3540
      Visible         =   0   'False
      Width           =   2865
   End
End
Attribute VB_Name = "AniGif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_File As String
Private FrameCount As Long
Private m_StopAtFirstFrame As Boolean
Private m_SpinFast As Boolean
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Click()
Public Sub CopyTransImage()
Image1(FrameCount).BackColor = vbWhite
CreateMaskImage Image1(FrameCount), Picture5
Image1(FrameCount).BackColor = 0
CreateMaskImage Image1(FrameCount), Picture2
BitBlt Picture5.hDC, 0, 0, LogicalWidth, LogicalHeight, Picture2.hDC, 0, 0, vbSrcAnd

BitBlt Picture6.hDC, 0, 0, Image1(FrameCount).ScaleWidth, Image1(FrameCount).ScaleHeight, Picture5.hDC, 0, 0, vbSrcCopy
BitBlt Picture6.hDC, 0, 0, Image1(FrameCount).ScaleWidth, Image1(FrameCount).ScaleHeight, Image1(FrameCount).hDC, 0, 0, vbSrcErase
Picture6.Refresh

BitBlt Picture7.hDC, Image1(FrameCount).Left, Image1(FrameCount).Top, Image1(FrameCount).ScaleWidth, Image1(FrameCount).ScaleHeight, Picture5.hDC, 0, 0, vbSrcAnd
BitBlt Picture7.hDC, Image1(FrameCount).Left, Image1(FrameCount).Top, Image1(FrameCount).ScaleWidth, Image1(FrameCount).ScaleHeight, Picture6.hDC, 0, 0, vbSrcPaint
End Sub


Public Sub LoadFile(ByVal FileName As String, ByVal DoNotStart As Boolean)
m_StopAtFirstFrame = False
m_SpinFast = False
Picture7.Cls
Timer1.Enabled = False
m_File = Trim(FileName)
  If Len(FileName) = 0 Then
    Timer1.Enabled = False
    m_File = ""
    TotalFrames = 0
    Exit Sub
  End If

Dim nFrames As Long
nFrames = LoadGif(m_File, Image1)
   
If nFrames > 0 Then
UserControl.Width = LogicalWidth * Screen.TwipsPerPixelX
UserControl.Height = LogicalHeight * Screen.TwipsPerPixelY
Picture5.Width = LogicalWidth
Picture6.Width = LogicalWidth
Picture2.Width = LogicalWidth
Picture5.Height = LogicalHeight
Picture6.Height = LogicalHeight
Picture2.Height = LogicalHeight
FrameCount = 0
  If nFrames > 1 And Not DoNotStart Then
    Timer1.Enabled = True
  Else
    Timer1_Timer
  End If
End If
End Sub

Private Sub Picture7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseUp(Button, Shift, x, y)
End Sub


Private Sub Timer1_Timer()
Dim i As Long
FrameCount = FrameCount + 1
If FrameCount > TotalFrames Then FrameCount = 1
      
   'redraw?
   i = FrameCount - 1
   If i = 0 Then i = TotalFrames
   Select Case Image1(i).DrawStyle
   Case 0, 1
   CopyTransImage
   Case 2
   Picture7.Line (Image1(i).Left, Image1(i).Top)-Step(Image1(i).Width - 1, Image1(i).Height - 1), Picture7.BackColor, BF
   CopyTransImage
   End Select
   Picture7.Refresh
   
   If m_SpinFast Then
   Timer1.Interval = 1
   Else
     Timer1.Interval = CLng(Image1(FrameCount).Tag)
   End If
   
If m_StopAtFirstFrame And FrameCount = 1 Then
  Timer1.Enabled = False
  m_StopAtFirstFrame = False
  m_SpinFast = False
End If

If TotalFrames = 1 Then
  Timer1.Interval = 1
  Timer1.Enabled = False
End If
End Sub



Public Sub StopAnimate(ByVal StopAtFirstFrame As Boolean, ByVal StopAtOnce As Boolean)
m_StopAtFirstFrame = False
m_SpinFast = False

If Not Timer1.Enabled Then Exit Sub
If StopAtFirstFrame Then
   If StopAtOnce Then
      Timer1.Enabled = False
      FrameCount = 0
      Picture7.Cls
      Timer1_Timer
      Exit Sub
   Else
      m_StopAtFirstFrame = True
      Exit Sub
   End If
End If
Timer1.Enabled = False
End Sub

Public Sub RestartAnimate()
If TotalFrames = 0 Then Exit Sub
m_StopAtFirstFrame = False
m_SpinFast = False
Timer1.Enabled = True
End Sub

Public Sub NextFrame()
If TotalFrames = 0 Then Exit Sub
m_StopAtFirstFrame = False
m_SpinFast = False
Timer1_Timer
End Sub

Private Sub UserControl_Initialize()
Picture7.BackColor = vbWhite
myBackColor = vbWhite
'gif specs
sGifMagic = Chr$(0) & Chr$(&H21) & Chr$(&HF9)
Trailer = Chr(59)
End Sub

Private Sub UserControl_Resize()
Picture7.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub





Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
Picture7.BackColor = vNewValue
myBackColor = vNewValue
End Property

Public Property Get TimerStatus() As Boolean
TimerStatus = Timer1.Enabled
End Property


Public Sub FinishCycleFast()
If Not Timer1.Enabled Then Exit Sub
m_StopAtFirstFrame = True
m_SpinFast = True
End Sub

