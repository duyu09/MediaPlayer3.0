VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DuyuPlayer3.1"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   435
   ClientWidth     =   15165
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   15165
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer4 
      Interval        =   800
      Left            =   240
      Top             =   2400
   End
   Begin VB.Timer Timer3 
      Interval        =   20
      Left            =   240
      Top             =   1560
   End
   Begin VB.Frame Frame1 
      Caption         =   "DuyuPlayer控制器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Top             =   5760
      Width           =   15135
      Begin VB.HScrollBar HScroll3 
         Height          =   135
         Left            =   12000
         Max             =   80
         TabIndex        =   20
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12120
         Sorted          =   -1  'True
         TabIndex        =   19
         Text            =   "Combo1"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   375
         Left            =   9960
         Max             =   100
         TabIndex        =   17
         Top             =   1080
         Value           =   50
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         Caption         =   "静音"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   16
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   15
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   1080
         Width           =   5895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "停止"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "播放"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   200
         Left            =   120
         Max             =   100
         TabIndex        =   10
         Top             =   480
         Width           =   11295
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   12120
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label6 
         Caption         =   "00:00\00:00"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.ListBox List2 
      Height          =   1680
      Left            =   5880
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   840
   End
   Begin VB.ListBox Listtemp 
      Height          =   2040
      Left            =   8040
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   3120
      Left            =   9240
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "歌曲信息栏"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   5400
      Width           =   15015
   End
   Begin VB.Image Image1 
      Height          =   5775
      Left            =   0
      Top             =   0
      Width           =   1935
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   2475
      Left            =   0
      TabIndex        =   5
      Tag             =   "无URL"
      Top             =   0
      Width           =   2400
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   4233
      _cy             =   4366
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   6240
      Width           =   8655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   8655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   8655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function ChangeVn(FileAllName As String) As String
For zxc = 1 To Len(FileAllName)
    If Mid(FileAllName, zxc, 1) = "." Then
       adf = zxc
    End If
Next zxc
ChangeVn = Left(FileAllName, adf) & "lrc"
End Function
Private Function Bjzfc(S1 As String, S2 As String) As Boolean
Dim I As Integer, T1 As Long, T2 As Long
For I = 1 To Len(S1)
If Val(Mid(S1, I, 1)) > 0 Then T1 = Val(Mid(S1, I)): Exit For
Next I
For I = 1 To Len(S1)
If Val(Mid(S2, I, 1)) > 0 Then T2 = Val(Mid(S2, I)): Exit For
Next I
If T1 > T2 Then Bjzfc = False Else Bjzfc = True
End Function
Private Sub LoadLRC(LRCfilefullpath As String)
List1.Clear
List2.Clear
Listtemp.Clear
TITLE = "未知标题"
ALBUM = "未知专辑"
AR = "未知艺术家"
OFFSET = ""
NumOfLine = 1
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Dim ConL As Integer
Open LRCfilefullpath For Input As #1
     ConL = 1
     Do While Not EOF(1)
        Line Input #1, LineLy(ConL)
        ConL = ConL + 1
     Loop
Close #1

For a = 1 To ConL - 1
    DoEvents
    On Error Resume Next
    If Left(LineLy(a), 4) = "[ti:" Then
       TITLE = Mid(LineLy(a), 5, InStr(1, LineLy(a), "]") - 5)
    ElseIf Left(LineLy(a), 4) = "[ar:" Then
       AR = Mid(LineLy(a), 5, InStr(1, LineLy(a), "]") - 5)
    ElseIf Left(LineLy(a), 4) = "[al:" Then
       ALBUM = Mid(LineLy(a), 5, InStr(1, LineLy(a), "]") - 5)
    ElseIf Left(LineLy(a), 8) = "[offset:" Then
       OFFSET = Mid(LineLy(a), 9, InStr(1, LineLy(a), "]") - 9)
    End If
Next a

Dim Bb As Integer, Cc As Integer
Cc = 1
For Bb = 1 To ConL - 1
    DoEvents
    If Left(LineLy(Bb), 2) = "[0" Then
       List1.AddItem LineLy(Bb) & " "
       Listtemp.AddItem CStr(Val(Mid(LineLy(Bb), 2, 2)) * 60 + Val(Mid(LineLy(Bb), 5, 2)) + Val(Mid(LineLy(Bb), 8, 2)) * 0.01)
       Cc = Cc + 1
    End If
Next Bb

For I = 1 To Listtemp.ListCount - 1
  For j = 0 To I - 1
   If Bjzfc(Listtemp.List(j), Listtemp.List(I)) Then
       k = Listtemp.List(I)
       Listtemp.List(I) = Listtemp.List(j)
       Listtemp.List(j) = k
   End If
  Next j
Next I
For sdfl = 0 To Listtemp.ListCount - 1
    List2.AddItem Listtemp.List(Listtemp.ListCount - 1 - sdfl)
Next sdfl
End Sub

Private Sub Combo1_Change()
Me.Label1.FontName = Me.Combo1.Text
Me.Label2.FontName = Me.Combo1.Text
Me.Label3.FontName = Me.Combo1.Text
Me.Label4.FontName = Me.Combo1.Text
Me.Label5.FontName = Me.Combo1.Text
End Sub

Private Sub Combo1_Click()
Me.Label1.FontName = Me.Combo1.Text
Me.Label2.FontName = Me.Combo1.Text
Me.Label3.FontName = Me.Combo1.Text
Me.Label4.FontName = Me.Combo1.Text
Me.Label5.FontName = Me.Combo1.Text
End Sub

Private Sub Command1_Click()
Timer2.Enabled = False
Timer3.Enabled = False
Me.WindowsMediaPlayer1.Controls.stop
Me.WindowsMediaPlayer1.URL = ""
Me.WindowsMediaPlayer1.Tag = "无URL"
List1.Clear
List2.Clear
Listtemp.Clear
TITLE = "未知标题"
ALBUM = "未知专辑"
AR = "未知艺术家"
OFFSET = ""
NumOfLine = 1
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label7.Caption = "歌曲信息栏"
Label6.Caption = "00:00\00:00"
Command2.Caption = "播放"
Me.Caption = "DuyuPlayer3.1"
For sd = 1 To 32767
    LineLy(sd) = ""
Next sd
tt = 0
NumOfLine = 1
End Sub

Private Sub Command2_Click()
Timer3.Enabled = True
If Command2.Caption = "播放" Then
   
   Command2.Caption = "暂停"
   Me.Caption = "DuyuPlayer3.1 - 正在播放:" & Text1.Text
  
   
   If Me.WindowsMediaPlayer1.Tag = "已有URL" Then
      Me.WindowsMediaPlayer1.Controls.play
      Timer2.Enabled = True
   Else
      If Dir(ChangeVn(Text1.Text)) <> "" Then
         LoadLRC (ChangeVn(Text1.Text))
      End If
      xb = InStr(List1.List(NumOfLine), "]")
      On Error Resume Next
      Label4.Caption = Right(List1.List(0), Len(List1.List(0)) - xb)
      On Error Resume Next
      Label5.Caption = Right(List1.List(1), Len(List1.List(1)) - xb)
      Label7.Caption = "标题：" & TITLE & " 艺术家：" & AR & " 专辑：" & ALBUM
   
      Timer2.Enabled = True
      Me.WindowsMediaPlayer1.URL = Text1.Text
      Me.WindowsMediaPlayer1.Tag = "已有URL"
   End If
Else
   Command2.Caption = "播放"
   Me.WindowsMediaPlayer1.Controls.pause
   Me.Caption = "DuyuPlayer3.1 - 已暂停"
   Timer2.Enabled = False
End If
End Sub

Private Sub Command5_Click()
Dim ofn As OPENFILENAME
Dim rtn As String, fp As String
ofn.lStructSize = Len(ofn)
ofn.hwndOwner = Me.hWnd
ofn.hInstance = App.hInstance
ofn.lpstrFilter = "所有文件(*.*)"
ofn.lpstrFile = Space(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = App.Path
ofn.lpstrTitle = "请选择音频文件 - DuyuPlayer"
ofn.flags = 6148
rtn = GetOpenFileName(ofn)
If rtn >= 1 Then
Text1.Text = ofn.lpstrFile
End If
End Sub

Private Sub Command6_Click()
Me.HScroll2.Value = 0
End Sub

Private Sub Form_Load()
WindowsMediaPlayer1.settings.volume = HScroll2.Value
Me.Label1.Width = Me.Width - 480
Me.Label2.Width = Me.Width - 480
Me.Label3.Width = Me.Width - 480
Me.Label4.Width = Me.Width - 480
Me.Label5.Width = Me.Width - 480
Me.HScroll1.Width = Me.Width - 495
Frame1.Top = Abs(Me.Height - Frame1.Height)
Frame1.Width = Me.Width - 100
Me.HScroll1.Width = Frame1.Width - 333
Label6.Left = Abs(Frame1.Width - 1400)
Label7.Top = Abs(Frame1.Top - Label7.Height)
TITLE = "未知标题"
ALBUM = "未知专辑"
AR = "未知艺术家"
Me.Combo1.Text = Me.Label1.FontName
For Counter1 = 0 To Screen.FontCount - 1
    Me.Combo1.AddItem Screen.Fonts(Counter1)
Next
Dim RealCom As String
If Command <> "" Then
    If Left(Command(), 1) = Chr(34) Then
       RealCom = Mid(Command(), 2, Len(Command()) - 2)
    Else
       RealCom = Command()
    End If
    Text1.Text = RealCom
    Command2_Click
End If
End Sub

Private Sub Form_Resize()
Me.Label1.Width = Me.Width - 480
Me.Label2.Width = Me.Width - 480
Me.Label3.Width = Me.Width - 480
Me.Label4.Width = Me.Width - 480
Me.Label5.Width = Me.Width - 480
Me.HScroll1.Width = Me.Width - 495
Frame1.Top = Abs(Me.Height - Frame1.Height)
Frame1.Width = Me.Width - 100
Me.HScroll1.Width = Frame1.Width - 333
Label6.Left = Abs(Frame1.Width - 1400)
Label7.Top = Abs(Frame1.Top - Label7.Height)
Me.WindowsMediaPlayer1.Top = 0
Me.WindowsMediaPlayer1.Left = 0
Me.WindowsMediaPlayer1.Width = Me.Width * 0.01 * Me.HScroll3.Value
Me.WindowsMediaPlayer1.Height = Me.Height * 0.01 * Me.HScroll3.Value
End Sub


Private Sub Frame1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.Text = Data.Files(1)
End Sub

Private Sub HScroll1_Scroll()
On Error Resume Next
WindowsMediaPlayer1.Controls.currentPosition = Me.HScroll1.Value * WindowsMediaPlayer1.currentMedia.duration / 100
NumOfLine = 1
End Sub

Private Sub HScroll2_Change()
WindowsMediaPlayer1.settings.volume = HScroll2.Value
End Sub

Private Sub HScroll3_Scroll()
Me.WindowsMediaPlayer1.Top = 0
Me.WindowsMediaPlayer1.Left = 0
Me.WindowsMediaPlayer1.Width = Me.Width * 0.01 * Me.HScroll3.Value
Me.WindowsMediaPlayer1.Height = Me.Height * 0.01 * Me.HScroll3.Value
End Sub

Private Sub Timer1_Timer()
If Len(Label3.Caption) > 48 Then
   Label3.Caption = Mid(Label3.Caption, 2, Len(Label3.Caption) - 1) + Left(Label3.Caption, 1)
End If
End Sub

Private Sub Timer2_Timer()
xb = InStr(List1.List(NumOfLine), "]")
tt = WindowsMediaPlayer1.Controls.currentPosition + Val(OFFSET) / 1000
If tt >= Val(List2.List(NumOfLine)) Then
   On Error Resume Next
   Label3.Caption = Right(List1.List(NumOfLine), Len(List1.List(NumOfLine)) - xb)
   
   On Error Resume Next
   Label4.Caption = Right(List1.List(NumOfLine + 1), Len(List1.List(NumOfLine + 1)) - xb)
   If Err.Number > 0 Then Label4.Caption = ""
      
   On Error Resume Next
   Label5.Caption = Right(List1.List(NumOfLine + 2), Len(List1.List(NumOfLine + 2)) - xb)
   If Err.Number > 0 Then Label5.Caption = ""
   
   On Error Resume Next
   Label1.Caption = Right(List1.List(NumOfLine - 2), Len(List1.List(NumOfLine - 2)) - xb)
   If Err.Number > 0 Then Label1.Caption = ""
   
   On Error Resume Next
   Label2.Caption = Right(List1.List(NumOfLine - 1), Len(List1.List(NumOfLine - 1)) - xb)
   If Err.Number > 0 Then Label2.Caption = ""
   
   NumOfLine = NumOfLine + 1
End If
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Label6.Caption = WindowsMediaPlayer1.Controls.currentPositionString & "\" & WindowsMediaPlayer1.currentMedia.durationString
If Err.Number > 0 Then
   Label6.Caption = "00:00\00:00"
End If
End Sub

Private Sub Timer4_Timer()
On Error Resume Next
Me.HScroll1.Value = 100 * WindowsMediaPlayer1.Controls.currentPosition / (WindowsMediaPlayer1.currentMedia.duration)
End Sub


Private Sub WindowsMediaPlayer1_StatusChange()
If WindowsMediaPlayer1.playState = wmppsMediaEnded Then
   Timer2.Enabled = False
Me.WindowsMediaPlayer1.Controls.stop
Me.WindowsMediaPlayer1.URL = ""
Me.WindowsMediaPlayer1.Tag = "无URL"
List1.Clear
List2.Clear
Listtemp.Clear
TITLE = "未知标题"
ALBUM = "未知专辑"
AR = "未知艺术家"
OFFSET = ""
NumOfLine = 1
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label7.Caption = "歌曲信息栏"
Label6.Caption = "00:00\00:00"
Command2.Caption = "播放"
Me.Caption = "DuyuPlayer3.1"
For sd = 1 To 32767
    LineLy(sd) = ""
Next sd
tt = 0
NumOfLine = 1
End If
End Sub
