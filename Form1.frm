VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   7200
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer4 
      Interval        =   60000
      Left            =   360
      Top             =   3240
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   2280
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1200
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1800
      Top             =   1560
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   840
      Picture         =   "Form1.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   1320
      Picture         =   "Form1.frx":20D1
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   0
      Picture         =   "Form1.frx":42F6
      Top             =   0
      Width           =   240
   End
   Begin VB.Menu menu 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu item1 
         Caption         =   "图标显示"
         Begin VB.Menu icons 
            Caption         =   "鼠标移过时"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu icons 
            Caption         =   "总是显示"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu icons 
            Caption         =   "不显示"
            Checked         =   -1  'True
            Index           =   2
         End
      End
      Begin VB.Menu ani 
         Caption         =   "动画"
         Checked         =   -1  'True
      End
      Begin VB.Menu Exit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'因为vb6蛋疼，信息提示的气泡效果需要对编译出的exe这样修改：
'  用exescope打开编译得到的exe，然后点“更改到XP样式”即可
Const SHOWMESSAGE = True                                                        '设置成true的话会在某些时候显示信息哦

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const flag = SWP_NOMOVE Or SWP_NOSIZE
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const HTCAPTION = 2
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA As Long = &H2
Private Const WS_EX_LAYERED As Long = &H80000
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Dim x1, y1, x2, y2
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim pth As String
Dim cs(10, 3) As Long
Dim nc As Integer, stg As Integer, nm As Integer, HeartMode As Integer
Dim MouseIn As Boolean
Dim tip As New CTips
Dim msg1 As String, msg2 As String, ymsg As Boolean
Dim mcount As Integer, dcount As Integer
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub ani_Click()
    ani.Checked = Not ani.Checked
End Sub

Private Sub Exit_Click()
    End
End Sub

Private Sub form_Initialize()
InitCommonControls

End Sub

Private Sub Form_Load()
    Randomize
    cs(0, 0) = 255
    cs(1, 0) = 255
    cs(1, 1) = 255
    cs(2, 1) = 255
    cs(3, 1) = 255
    cs(3, 2) = 255
    cs(4, 2) = 255
    cs(5, 2) = 255
    cs(5, 0) = 255
    Form1.BorderStyle = 0
    Form1.Width = 705
    Form1.Height = 350
    pth = App.Path
    If right(pth, 1) <> "\" Then pth = pth + "\"
    Form1.left = Val(GetSetting("Ideary", "Data1", "Left", CStr(Screen.Width - Form1.Width - 200 * Screen.TwipsPerPixelX)))
    Form1.top = Val(GetSetting("Ideary", "Data1", "Top", CStr(20 * Screen.TwipsPerPixelY)))
    HeartMode = Val(GetSetting("Ideary", "Data1", "HeartMode", "0"))
    For i = 0 To 2
        icons(i).Checked = (HeartMode = i)
    Next i
    '这里就是蛋疼的消息提示了
    If Hour(Now) >= 6 And Hour(Now) <= 10 Then
        msg1 = "早上好"
        msg2 = "新的一天开始啦，开心地出发吧"
    ElseIf Hour(Now) >= 11 And Hour(Now) <= 13 Then
        msg1 = "中午好"
        msg2 = "上午之后，累了吗？也许睡个午觉不错哦"
    ElseIf Hour(Now) >= 14 And Hour(Now) <= 18 Then
        msg1 = "下午好"
        msg2 = "有很多事情要忙吗？一定要加油啊"
    ElseIf Hour(Now) >= 19 And Hour(Now) <= 23 Then
        msg1 = "晚上好"
        msg2 = "今天的心情怎样呢？发短信说说吧，累了请快休息"
    Else
        msg1 = "半夜...好"
        msg2 = "熬夜的话，少喝点咖啡吧，加油加油"
    End If
    ymsg = True
    Me.Show
    x1 = ScaleX(Me.left, vbTwips, vbPixels)
    y1 = ScaleY(Me.top, vbTwips, vbPixels)
    x2 = ScaleX(Me.Width, vbTwips, vbPixels)
    y2 = ScaleY(Me.Height, vbTwips, vbPixels)
    ani.Enabled = True
    Timer1_Timer
    Dim p As Long
    p = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, p Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(Me.hwnd, 0, 160, LWA_ALPHA)
End Sub

Private Sub icons_Click(Index As Integer)
    For i = 0 To 2
        icons(i).Checked = (i = Index)
    Next i
    HeartMode = Index
    SaveSetting "Ideary", "Data1", "HeartMode", Trim(Index)
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        '如果按下鼠标左键
        Dim ReturnVal As Long
        X = ReleaseCapture()
        ReturnVal = SendMessage(Form1.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0)
        'On Error Resume Next
        SaveSetting "Ideary", "Data1", "Left", Form1.left
        SaveSetting "Ideary", "Data1", "Top", Form1.top
        x1 = ScaleX(Me.left, vbTwips, vbPixels)
        y1 = ScaleY(Me.top, vbTwips, vbPixels)
        x2 = ScaleX(Me.Width, vbTwips, vbPixels)
        y2 = ScaleY(Me.Height, vbTwips, vbPixels)
        
        If ymsg And X <= Image4.Width And Y <= Image4.Height And X >= 0 And Y >= 0 And SHOWMESSAGE Then
            ymsg = False
            Image4.Visible = False
            tip.Style = TTBalloon
            tip.Icon = TTIconinfo                                               '图标类型
            tip.Title = msg1                                                    '标题
            tip.TipText = msg2                                                  '内容
            tip.PopupOnDemand = True
            tip.CreateToolTip Me.hwnd                                           '气泡所在窗口句柄
            tip.Show
            dcount = 0
            Timer3.Enabled = True
        End If
    Else
        PopupMenu menu
    End If
End Sub

Function fe(fn As String) As Boolean
    On Error Resume Next
    fe = Dir$(fn, vbNormal + vbReadOnly + vbHidden + vbSystem + vbVolume) <> ""
    If Err.Number <> 0 Then fe = False
End Function

Private Sub Timer1_Timer()
    Dim d1 As Date, d2 As Date
    d1 = #12/31/2009#
    d2 = Now
    Label1.Caption = Trim(Int(d2 - d1))
    Form1.Width = 535 + (Len(Label1.Caption) - 2) * 170
    Label1.Width = Form1.Width - 240
    If Hour(Now) = 12 And Minute(Now) = 0 And Second(Now) < 10 Then
        ymsg = True
        msg1 = "午餐时间到"
        msg2 = "忙了一上午，该吃一顿丰盛的午餐了，这样才有能量嘛"
        mcount = 0
    End If
    If Hour(Now) = 7 And Minute(Now) = 30 And Second(Now) < 10 Then
        ymsg = True
        msg1 = "早餐时间到"
        msg2 = "一日之计在于晨，美味的早餐能带来一天的好心情哦"
        mcount = 0
    End If
    If Hour(Now) = 7 And Minute(Now) = 30 And Second(Now) < 10 Then
        ymsg = True
        msg1 = "晚餐时间到"
        msg2 = "晚饭不一定多但一定要有营养，千万不要不吃 >_<"
        mcount = 0
    End If
    If Hour(Now) = 0 And Minute(Now) = 0 And Second(Now) < 10 Then
        ymsg = True
        msg1 = "凌晨了"
        msg2 = "早点睡吧...熬夜要注意身体，少喝咖啡"
        mcount = 0
    End If
    If right(Label1.Caption, 2) = "00" Then
        Image2.Visible = True
    Else
        Image2.Visible = False
        If (HeartMode = 0 And MouseIn) Or HeartMode = 1 Then
            Image1.Visible = True
        Else
            Image1.Visible = False
        End If
    End If
    If Image1.Visible = True Then
        Image1.left = Form1.Width
        Form1.Width = Form1.Width + Image1.Width
    End If
    If Image2.Visible = True Then
        Image2.left = Form1.Width
        Form1.Width = Form1.Width + Image2.Width
    End If
    If ymsg And SHOWMESSAGE Then
        Image4.Visible = True
    Else
        Image4.Visible = False
    End If
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
End Sub

Private Sub Timer2_Timer()
    '处理颜色变化之类的
    If ani.Checked = False Then Exit Sub
    If nm = 0 Then
        stg = stg + 1
        If stg = 16 Then
            stg = 0
            nc = nc + 1
        End If
        If nc = 6 Then
            nc = 0
            If Rnd() > 0.7 Then nm = 1
        End If
        Label1.ForeColor = RGB(CLng((16 - stg) / 16 * cs(nc, 0) + stg / 16 * cs((nc + 1) Mod 6, 0)), CLng((16 - stg) / 16 * cs(nc, 1) + stg / 16 * cs((nc + 1) Mod 6, 1)), CLng((16 - stg) / 16 * cs(nc, 2) + stg / 16 * cs((nc + 1) Mod 6, 2)))
    ElseIf nm = 1 Then
        Label1.ForeColor = RGB(CLng((16 - stg) / 16 * cs(nc, 0) + stg / 16 * 255), CLng((16 - stg) / 16 * cs(nc, 1) + stg / 16 * 255), CLng((16 - stg) / 16 * cs(nc, 2) + stg / 16 * 255))
        stg = stg + 1
        If stg = 16 Then
            stg = 0
            nm = 2
            nc = nc + 1
        End If
        If nc = 6 Then
            nc = 0
            nm = 3
        End If
    ElseIf nm = 2 Then
        Label1.ForeColor = RGB(CLng((16 - stg) / 16 * 255 + stg / 16 * cs(nc, 0)), CLng((16 - stg) / 16 * 255 + stg / 16 * cs(nc, 1)), CLng((16 - stg) / 16 * 255 + stg / 16 * cs(nc, 2)))
        stg = stg + 1
        If stg = 16 Then
            stg = 0
            nm = 1
        End If
        If nc = 6 Then
            nc = 0
            If Rnd() > 0.7 Then nm = 0
        End If
    ElseIf nm = 3 Then
        Label1.ForeColor = RGB(255, CLng(16 - stg) / 16 * 255, CLng(16 - stg) / 16 * 255)
        stg = stg + 1
        If stg = 16 Then
            stg = 0
            nm = 0
        End If
    End If
    
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    Dim a As Boolean
    a = MouseIn
    If lpPoint.X > x1 And lpPoint.X < x1 + x2 And lpPoint.Y > y1 And lpPoint.Y < y1 + y2 Then
        MouseIn = True
    Else
        MouseIn = False
    End If
    If a <> MouseIn Then
        Timer1_Timer
        If MouseIn Then
            p = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
            Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, p Or WS_EX_LAYERED)
            Call SetLayeredWindowAttributes(Me.hwnd, 0, 255, LWA_ALPHA)
        Else
            p = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
            Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, p Or WS_EX_LAYERED)
            Call SetLayeredWindowAttributes(Me.hwnd, 0, 160, LWA_ALPHA)
            
        End If
    End If
End Sub

'气泡提示5秒消失
Private Sub Timer3_Timer()
    dcount = dcount + 1
    If dcount > 5 Then
        Timer3.Enabled = False
        tip.Destroy
        dcount = 0
    End If
End Sub

'这里是管消息的自动消失的，30分钟
Private Sub Timer4_Timer()
    If Not ymsg Then Exit Sub
    mcount = mcount + 1
    If mcount > 30 Then
        mcount = 0
        ymsg = False
    End If
End Sub
