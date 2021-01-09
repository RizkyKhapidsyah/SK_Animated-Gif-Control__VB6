VERSION 5.00
Object = "{E6C4280E-288E-41E1-B348-A0E583B65166}#1.1#0"; "AnimatedGif.ocx"
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   2160
   ClientTop       =   2070
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   1095
      TabIndex        =   8
      Top             =   5895
      Width           =   3495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SpinDown"
      Height          =   495
      Left            =   5355
      TabIndex        =   7
      Top             =   7305
      Width           =   1215
   End
   Begin AnimatedGif.AniGif AniGif1 
      Height          =   1875
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3307
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Backcolor"
      Height          =   495
      Left            =   9525
      TabIndex        =   5
      Top             =   7305
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Next"
      Height          =   495
      Left            =   9510
      TabIndex        =   4
      Top             =   6675
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Restart"
      Height          =   495
      Left            =   8085
      TabIndex        =   3
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "stop"
      Height          =   495
      Left            =   6675
      TabIndex        =   2
      Top             =   7305
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "unload"
      Height          =   495
      Left            =   8085
      TabIndex        =   1
      Top             =   6675
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "load"
      Height          =   495
      Left            =   6705
      TabIndex        =   0
      Top             =   6675
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LB_DIR As Long = &H18D
Private Const DDL_ARCHIVE As Long = &H20
Private Const DDL_EXCLUSIVE As Long = &H8000
Private Const DDL_FLAGS As Long = DDL_ARCHIVE Or DDL_EXCLUSIVE
 
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Dim MyPath As String

Sub AddDirSep(strPathName As String)
    If Right$(RTrim$(strPathName), 1) <> "\" Then
    strPathName = RTrim$(strPathName) & "\"
    End If
End Sub

Private Sub Command1_Click()
AniGif1.LoadFile MyPath & List1.Text, False
End Sub


Private Sub Command2_Click()
AniGif1.LoadFile "", False
End Sub


Private Sub Command3_Click()
AniGif1.StopAnimate True, True
End Sub


Private Sub Command4_Click()
AniGif1.RestartAnimate
End Sub


Private Sub Command5_Click()
AniGif1.NextFrame
End Sub




Private Sub Command6_Click()
AniGif1.BackColor = vbYellow
End Sub


Private Sub Command7_Click()
AniGif1.FinishCycleFast
End Sub



Private Sub Form_Load()
MyPath = App.Path
AddDirSep MyPath
   Call SendMessage(List1.hwnd, _
                    LB_DIR, _
                    DDL_FLAGS, _
                    ByVal MyPath & "*.gif")
End Sub


