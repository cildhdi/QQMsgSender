VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "QQBomp"
   ClientHeight    =   5520
   ClientLeft      =   7320
   ClientTop       =   3780
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   Picture         =   "Main.frx":0000
   ScaleHeight     =   5520
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   4410
      Top             =   1590
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2505
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Main.frx":5EC3
      Top             =   4695
      Width           =   1920
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1905
      TabIndex        =   6
      Top             =   2610
      Width           =   5190
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   3390
      ScaleHeight     =   510
      ScaleWidth      =   525
      TabIndex        =   2
      Top             =   1260
      Width           =   525
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发送内容："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   570
      TabIndex        =   8
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label Label5 
      Caption         =   "停止"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   4710
      Width           =   675
   End
   Begin VB.Label Label4 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4950
      TabIndex        =   4
      Top             =   4710
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "窗口标题："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   570
      TabIndex        =   3
      Top             =   2100
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "按住下方蓝色方块，将鼠标移到要发送的聊天窗口上松开。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   15
      TabIndex        =   1
      Top             =   840
      Width           =   7800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   7125
      X2              =   7350
      Y1              =   420
      Y2              =   165
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   7095
      X2              =   7335
      Y1              =   150
      Y2              =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   15
      X2              =   7455
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QQ消息轰炸机V 2.0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3750
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'必须的API函数以及全局常量和变量声明
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Const WM_PASTE As Long = &H302
Private Const CF_TEXT As Long = 1
Private Const WM_KEYDOWN = &H100
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Dim cPoint As POINTAPI
Dim mhwnd As Long
Dim num As Integer
Dim str As String
Dim bool As Boolean
Dim mCaption As String
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const WM_SYSCOMMAND As Long = &H112
Private Const SC_MOVE As Long = &HF010&
Private Const HTCAPTION As Long = 2
Private Const n As Integer = 100
Dim prgb As Long
Dim mx As Long
Dim my As Long


Private Sub Form_Load()
Dim rtn As Long
Dim pra As Integer
'将窗口设置为全透明
rtn = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hWnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Me.hWnd, 0, 0, LWA_ALPHA
Me.Hide
'检测系统字体目录里是否有该字体，没有就写入
If Dir("C:\WINDOWS\Fonts\造字工房尚黑粗体.otf") = "" Then
    Dim ttf() As Byte
    ttf = LoadResData(101, "CUSTOM")
    Open "C:\WINDOWS\Fonts\造字工房尚黑粗体.otf" For Binary As #1
    Put #1, , ttf
    Close #1
    Sleep 1000
    Me.Show
Else
    Me.Show
End If
'以下为一些窗口控件的属性设置
Label1.Font = "造字工房尚黑 G0v1 粗体"
Label1.Left = 0
Label1.Top = 0
Line1.X1 = 0
Line1.X2 = Me.Width
Label2.Font = "造字工房尚黑 G0v1 粗体"
Label3.Font = "造字工房尚黑 G0v1 粗体"
Label4.Font = "造字工房尚黑 G0v1 粗体"
Label5.Font = "造字工房尚黑 G0v1 粗体"
Label6.Font = "造字工房尚黑 G0v1 粗体"
Text1.Font = "造字工房尚黑 G0v1 粗体"
Text2.Font = "造字工房尚黑 G0v1 粗体"
Label4.Visible = False
Label5.Visible = False
Label4.BackColor = RGB(150, 0, 0)
Label5.BackColor = RGB(150, 0, 0)
Picture1.BackColor = vbBlue
'画边框
Me.Line (0, 0)-(Me.Width, 0), 5
Me.Line (0, 0)-(0, Me.Height), 5
Me.Line (Me.Width, 0)-(Me.Width, Me.Height), 5
Me.Line (0, Me.Height)-(Me.Width, Me.Height), 5
'画关闭的叉
Line2.X1 = Me.Width - Line1.Y1 / 2 - n
Line3.X1 = Me.Width - Line1.Y1 / 2 - n
Line2.X2 = Me.Width - Line1.Y1 / 2 + n
Line3.X2 = Me.Width - Line1.Y1 / 2 + n
Line2.Y1 = Line1.Y1 / 2 - n
Line3.Y1 = Line1.Y1 / 2 + n
Line2.Y2 = Line1.Y1 / 2 + n
Line3.Y2 = Line1.Y1 / 2 - n


'窗体渐渐出现
For pra = 0 To 240 Step 10
    rtn = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes Me.hWnd, 0, pra, LWA_ALPHA
    DoEvents
    Sleep 50
Next pra
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < Line1.Y1 Then '窗口可移动
ReleaseCapture
SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
    If X > (Me.Width - Line1.Y1) And Y < Line1.Y1 Then '点击关闭时
        Dim rtn As Long                                '窗口渐渐消失
        Dim pra As Integer
        For pra = 240 To 0 Step -10
            rtn = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
            rtn = rtn Or WS_EX_LAYERED
            SetWindowLong hWnd, GWL_EXSTYLE, rtn
            SetLayeredWindowAttributes Me.hWnd, 0, pra, LWA_ALPHA
            DoEvents                 '如果不加这一句，窗体在变化过程中为黑色
            Sleep 50
        Next pra
    End
    End If
End If
End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > (Me.Width - Line1.Y1) And Y < Line1.Y1 And X < Me.Width And Y > 0 Then
    Line2.BorderColor = vbRed
    Line3.BorderColor = vbRed
Else
    Line2.BorderColor = vbWhite
    Line3.BorderColor = vbWhite
End If
Label4.BackColor = RGB(150, 0, 0)
Label5.BackColor = RGB(150, 0, 0)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture                                                  '按在标题上也可拖动
SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
End Sub


Private Sub Label4_Click()
str = Text1.Text
Timer1.Interval = Val(Text2.Text)
bool = True
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BackColor = RGB(255, 0, 0)
End Sub

Private Sub Label5_Click()
bool = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = RGB(255, 0, 0)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BackColor = vbRed
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BackColor = vbBlue
mCaption = ".............................................."
GetCursorPos cPoint
mhwnd = WindowFromPoint(cPoint.X, cPoint.Y)
GetWindowText mhwnd, mCaption, 1000
Label3.Caption = "窗口标题：" + mCaption
Label4.Visible = True
Label5.Visible = True
End Sub
Private Sub Text2_Change()
If Val(Text2.Text) > 65535 Then Text2.Text = 65535
End Sub
Private Sub Text2_GotFocus()
Text2.Text = ""
End Sub
Private Sub Timer1_Timer()
If bool Then
CopyTextToClip (str) '将文本复制到系统全局粘贴板
SendMessage mhwnd, WM_PASTE, 0, 0 '粘贴文本
SendMessage mhwnd, WM_KEYDOWN, vbKeyReturn, 0 '发送
End If
End Sub

Private Sub CopyTextToClip(sData As String) '该过程将文本复制到系统全局粘贴板
   If CBool(OpenClipboard(0)) Then
      Dim hMemHandle As Long, lpData As Long
      hMemHandle = GlobalAlloc(0, LenB(sData) + 2)
      If CBool(hMemHandle) Then
         lpData = GlobalLock(hMemHandle)
         If lpData <> 0 Then
            CopyMemory ByVal lpData, ByVal sData, LenB(sData)
            GlobalUnlock hMemHandle
            EmptyClipboard
            SetClipboardData CF_TEXT, hMemHandle
         End If
      End If
      Call CloseClipboard
   End If
End Sub

