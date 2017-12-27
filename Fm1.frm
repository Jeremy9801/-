VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Fm1 
   BackColor       =   &H0080C0FF&
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12330
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   12330
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "检测单词"
      Height          =   495
      Left            =   10920
      TabIndex        =   28
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox t8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9960
      TabIndex        =   27
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox t7 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6840
      TabIndex        =   26
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox t5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   480
      TabIndex        =   24
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox t4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9960
      TabIndex        =   23
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox t3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6840
      TabIndex        =   22
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox t2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3600
      TabIndex        =   21
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox t1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   360
      TabIndex        =   20
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton p4 
      Height          =   2295
      Left            =   9480
      Picture         =   "Fm1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton p3 
      Height          =   2295
      Left            =   6360
      Picture         =   "Fm1.frx":6D87
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton p2 
      Height          =   2295
      Left            =   9480
      Picture         =   "Fm1.frx":10798
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton p1 
      Height          =   2295
      Left            =   6360
      Picture         =   "Fm1.frx":17DC2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cd6 
      Caption         =   "下一页"
      Height          =   495
      Left            =   3960
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   10
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cd5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "上一页"
      Height          =   495
      Left            =   600
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   9
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton CM5 
      Caption         =   "重来"
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Cd4 
      Height          =   2295
      Left            =   3240
      Picture         =   "Fm1.frx":234C0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton Cd3 
      Height          =   2295
      Left            =   120
      Picture         =   "Fm1.frx":2CA7B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   2535
   End
   Begin VB.CommandButton Cd2 
      Height          =   2295
      Left            =   3240
      Picture         =   "Fm1.frx":35675
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Cd1 
      Height          =   2295
      Left            =   120
      Picture         =   "Fm1.frx":3E075
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox t6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3720
      TabIndex        =   25
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label lb7 
      BackStyle       =   0  'Transparent
      Caption         =   "panda"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lb8 
      BackStyle       =   0  'Transparent
      Caption         =   "cow"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   19
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image i8 
      Height          =   1080
      Left            =   11520
      Picture         =   "Fm1.frx":47419
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lb4 
      BackStyle       =   0  'Transparent
      Caption         =   "tortoise"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image i6 
      Height          =   1080
      Left            =   5280
      Picture         =   "Fm1.frx":47ED7
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image i7 
      Height          =   1080
      Left            =   8400
      Picture         =   "Fm1.frx":48995
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lb2 
      BackStyle       =   0  'Transparent
      Caption         =   "horse"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lb5 
      BackStyle       =   0  'Transparent
      Caption         =   "cat"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lb6 
      BackStyle       =   0  'Transparent
      Caption         =   "lion"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lb3 
      BackStyle       =   0  'Transparent
      Caption         =   "pig"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Lb1 
      BackStyle       =   0  'Transparent
      Caption         =   "bird"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image i5 
      Height          =   1080
      Left            =   2040
      Picture         =   "Fm1.frx":49453
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image i4 
      Height          =   1080
      Left            =   11520
      Picture         =   "Fm1.frx":49F11
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image i3 
      Height          =   1080
      Left            =   8400
      Picture         =   "Fm1.frx":4A9CF
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image i2 
      Height          =   1080
      Left            =   5160
      Picture         =   "Fm1.frx":4B48D
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image i1 
      Height          =   1080
      Left            =   1920
      Picture         =   "Fm1.frx":4BF4B
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "共3页"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11160
      TabIndex        =   33
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "页码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11160
      TabIndex        =   32
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   31
      Top             =   7080
      Width           =   615
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   1800
      TabIndex        =   30
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
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
      uiMode          =   "full"
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
      _cx             =   873
      _cy             =   873
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "在白框输入单词点击右侧按钮可检测拼写对错→"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   9000
      TabIndex        =   29
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "点击图片出现相应单词及语音"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   5040
      TabIndex        =   11
      Top             =   7920
      Width           =   3975
   End
End
Attribute VB_Name = "Fm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cd1_Click()
Lb1.Visible = True
Select Case Label3.Caption
Case "1"
    WindowsMediaPlayer1.URL = App.Path & "\bird.wma"
Case "2"
    WindowsMediaPlayer1.URL = App.Path & "\blue.wma"
Case Else
    WindowsMediaPlayer1.URL = App.Path & "\apple.wma"
End Select
End Sub

Private Sub Cd2_Click()
lb2.Visible = True
Select Case Label3.Caption
Case "1"
    WindowsMediaPlayer1.URL = App.Path & "\horse.wma"
Case "2"
    WindowsMediaPlayer1.URL = App.Path & "\red.wma"
Case Else
    WindowsMediaPlayer1.URL = App.Path & "\mango.wma"
End Select
End Sub

Private Sub Cd3_Click()
lb3.Visible = True
Select Case Label3.Caption
Case "1"
    WindowsMediaPlayer1.URL = App.Path & "\pig.wma"
Case "2"
    WindowsMediaPlayer1.URL = App.Path & "\green.wma"
Case Else
    WindowsMediaPlayer1.URL = App.Path & "\mangosteen.wma"
End Select
End Sub

Private Sub Cd4_Click()
lb4.Visible = True
Select Case Label3.Caption
Case "1"
    WindowsMediaPlayer1.URL = App.Path & "\tortoise.wma"
Case "2"
    WindowsMediaPlayer1.URL = App.Path & "\yellow.wma"
Case Else
    WindowsMediaPlayer1.URL = App.Path & "\orange.wma"
End Select
End Sub


Private Sub cd5_Click()
Call qing
Select Case Label3.Caption
Case "2"
    Call pic("/bird.jpg", "/horse.jpg", "/pig.jpg", "/tortoise.jpg", "/cat.jpg", "/lion.jpg", "/panda.jpg", "/cow.jpg")
    Call vis
    Call cl("bird", "horse", "pig", "tortoise", "cat", "lion", "panda", "cow")
Case "3"
    Call pic("/blue.jpg", "/red.jpg", "/green.jpg", "/yellow.jpg", "/gray.jpg", "/orange.jpg", "/purple.jpg", "/black.jpg")
    Call cl("blue", "red", "green", "yellow", "gray", "orange", "purple", "blacl")
    Call vis
End Select
Call yema
Call ima
End Sub

Private Sub cd6_Click()
Call qing
Select Case Label3.Caption
Case "1"
    Call pic("/blue.jpg", "/red.jpg", "/green.jpg", "/yellow.jpg", "/gray.jpg", "/orange.jpg", "/purple.jpg", "/black.jpg")
    Call cl("blue", "red", "green", "yellow", "gray", "orange", "purple", "blacl")
    Call vis
Case "2"
    Call pic("/apple.jpg", "/mango.jpg", "/mangosteen.jpg", "/orange2.jpg", "/peach.jpg", "/pear.jpg", "/pineapple.jpg", "/watemelon.jpg")
    Call cl("apple", "mango", "mangosteen", "orange", "peach", "pear", "pineapple", "watemelon")
    Call vis
End Select
Call yema
Call ima
End Sub

Private Sub CM5_Click()
Call yema
Call qing
Call vis
Call ima
End Sub

Private Sub Command1_Click()
If t1.Text = Lb1.Caption Then
    t1.ForeColor = &HFF00&
    i1.Picture = LoadPicture(App.Path & "/flw.jpg")
Else
    t1.ForeColor = &HFF&
    i1.Picture = LoadPicture(App.Path & "/cry.jpg")
End If

If t2.Text = lb2.Caption Then
    t2.ForeColor = &HFF00&
    i2.Picture = LoadPicture(App.Path & "/flw.jpg")
Else
    t2.ForeColor = &HFF&
    i2.Picture = LoadPicture(App.Path & "/cry.jpg")
End If

If t3.Text = lb5.Caption Then
    t3.ForeColor = &HFF00&
    i3.Picture = LoadPicture(App.Path & "/flw.jpg")
Else
    t3.ForeColor = &HFF&
    i3.Picture = LoadPicture(App.Path & "/cry.jpg")
End If

If t4.Text = lb6.Caption Then
    t4.ForeColor = &HFF00&
    i4.Picture = LoadPicture(App.Path & "/flw.jpg")
Else
    t4.ForeColor = &HFF&
    i4.Picture = LoadPicture(App.Path & "/cry.jpg")
End If

If t5.Text = lb3.Caption Then
    t5.ForeColor = &HFF00&
    i5.Picture = LoadPicture(App.Path & "/flw.jpg")
Else
    t5.ForeColor = &HFF&
    i5.Picture = LoadPicture(App.Path & "/cry.jpg")
End If

If t6.Text = lb4.Caption Then
    t6.ForeColor = &HFF00&
    i6.Picture = LoadPicture(App.Path & "/flw.jpg")
Else
    t6.ForeColor = &HFF&
    i6.Picture = LoadPicture(App.Path & "/cry.jpg")
End If

If t7.Text = lb7.Caption Then
    t7.ForeColor = &HFF00&
    i7.Picture = LoadPicture(App.Path & "/flw.jpg")
Else
    t7.ForeColor = &HFF&
    i7.Picture = LoadPicture(App.Path & "/cry.jpg")
End If

If t8.Text = lb8.Caption Then
    t8.ForeColor = &HFF00&
    i8.Picture = LoadPicture(App.Path & "/flw.jpg")
Else
    t8.ForeColor = &HFF&
    i8.Picture = LoadPicture(App.Path & "/cry.jpg")
End If
i1.Visible = True
i2.Visible = True
i3.Visible = True
i4.Visible = True
i5.Visible = True
i6.Visible = True
i7.Visible = True
i8.Visible = True

End Sub

Private Sub Form_Load()
Fm1.Caption = "趣味看图记单词(作者：温振杰)"
MsgBox "欢迎使用看图记单词！" + Chr(13) + "第一页为动物类！" + Chr(13) + "第二页为颜色类！" + Chr(13) + "第三页为水果类！" + Chr(13) + "点击图片可听取单词语音！" + Chr(13) + "使用愉快！", , "欢迎使用看图记单词"
End Sub

Public Sub vis()
Lb1.Visible = False
lb2.Visible = False
lb3.Visible = False
lb4.Visible = False
lb5.Visible = False
lb6.Visible = False
lb7.Visible = False
lb8.Visible = False
End Sub

Public Function pic(g1 As String, g2 As String, g3 As String, g4 As String, g5 As String, g6 As String, g7 As String, g8 As String)
Cd1.Picture = LoadPicture(App.Path & g1)
Cd2.Picture = LoadPicture(App.Path & g2)
Cd3.Picture = LoadPicture(App.Path & g3)
Cd4.Picture = LoadPicture(App.Path & g4)
p1.Picture = LoadPicture(App.Path & g5)
p2.Picture = LoadPicture(App.Path & g6)
p3.Picture = LoadPicture(App.Path & g7)
p4.Picture = LoadPicture(App.Path & g8)
End Function

Public Function cl(c1 As String, c2 As String, c3 As String, c4 As String, c5 As String, c6 As String, c7 As String, c8 As String)
Lb1.Caption = c1
lb2.Caption = c2
lb3.Caption = c3
lb4.Caption = c4
lb5.Caption = c5
lb6.Caption = c6
lb7.Caption = c7
lb8.Caption = c8
End Function

Private Sub p1_Click()
lb5.Visible = True
Select Case Label3.Caption
Case "1"
    WindowsMediaPlayer1.URL = App.Path & "\cat.wma"
Case "2"
    WindowsMediaPlayer1.URL = App.Path & "\gray.wma"
Case Else
    WindowsMediaPlayer1.URL = App.Path & "\peach.wma"
End Select
End Sub

Private Sub p2_Click()
lb6.Visible = True
Select Case Label3.Caption
Case "1"
    WindowsMediaPlayer1.URL = App.Path & "\lion.wma"
Case "2"
    WindowsMediaPlayer1.URL = App.Path & "\orange.wma"
Case Else
    WindowsMediaPlayer1.URL = App.Path & "\pear.wma"
End Select
End Sub

Private Sub p3_Click()
lb7.Visible = True
Select Case Label3.Caption
Case "1"
    WindowsMediaPlayer1.URL = App.Path & "\panda.wma"
Case "2"
    WindowsMediaPlayer1.URL = App.Path & "\purple.wma"
Case Else
    WindowsMediaPlayer1.URL = App.Path & "\pineapple.wma"
End Select
End Sub

Private Sub p4_Click()
lb8.Visible = True
Select Case Label3.Caption
Case "1"
    WindowsMediaPlayer1.URL = App.Path & "\cow.wma"
Case "2"
    WindowsMediaPlayer1.URL = App.Path & "\black.wma"
Case Else
    WindowsMediaPlayer1.URL = App.Path & "\watermelon.wma"
End Select
End Sub

Private Sub t1_Change()
t1.ForeColor = &H0&
End Sub

Private Sub t2_Change()
t2.ForeColor = &H0&
End Sub

Private Sub t3_Change()
t3.ForeColor = &H0&
End Sub

Private Sub t4_Change()
t4.ForeColor = &H0&
End Sub

Private Sub t5_Change()
t5.ForeColor = &H0&
End Sub

Private Sub t6_Change()
t6.ForeColor = &H0&
End Sub

Private Sub t7_Change()
t7.ForeColor = &H0&
End Sub

Private Sub t8_Change()
t8.ForeColor = &H0&
End Sub

Public Sub yema()
Select Case Lb1.Caption
Case "bird"
    Label3.Caption = "1"
Case "blue"
    Label3.Caption = "2"
Case Else
    Label3.Caption = "3"
End Select
End Sub

Public Sub qing()
t1.Text = ""
t2.Text = ""
t3.Text = ""
t4.Text = ""
t5.Text = ""
t6.Text = ""
t7.Text = ""
t8.Text = ""
End Sub

Public Sub ima()
i1.Visible = False
i2.Visible = False
i3.Visible = False
i4.Visible = False
i5.Visible = False
i6.Visible = False
i7.Visible = False
i8.Visible = False
End Sub


