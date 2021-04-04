VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTendPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "记录单选项"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4920
   Icon            =   "frmTendPara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3735
      TabIndex        =   17
      Top             =   3720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2580
      TabIndex        =   16
      Top             =   3720
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   3510
      Index           =   1
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   4800
      Begin VB.PictureBox picControl 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1845
         Left            =   1950
         ScaleHeight     =   1845
         ScaleWidth      =   2295
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1485
         Visible         =   0   'False
         Width           =   2295
         Begin VB.CommandButton cmdUnVisible 
            Height          =   315
            Left            =   1815
            Picture         =   "frmTendPara.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "取消"
            Top             =   1500
            Width           =   450
         End
         Begin VB.PictureBox PicColorCollect 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1350
            Left            =   60
            Picture         =   "frmTendPara.frx":0596
            ScaleHeight     =   1350
            ScaleWidth      =   2160
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   90
            Width           =   2160
            Begin VB.Shape shpValue 
               BorderColor     =   &H00C56A31&
               FillColor       =   &H00FF8080&
               Height          =   270
               Left            =   0
               Top             =   0
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Shape shpBorder 
               BorderColor     =   &H00C56A31&
               FillColor       =   &H00FF8080&
               Height          =   270
               Left            =   1890
               Top             =   1080
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   200
            Left            =   90
            ScaleHeight     =   165
            ScaleWidth      =   165
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1575
            Width           =   200
         End
         Begin VB.Label lblColor 
            Caption         =   "&HFFFFFF"
            Height          =   195
            Left            =   405
            TabIndex        =   22
            Top             =   1575
            UseMnemonic     =   0   'False
            Width           =   1365
         End
      End
      Begin VB.PictureBox picLineColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1950
         ScaleHeight     =   180
         ScaleWidth      =   2265
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "点击选择颜色"
         Top             =   1275
         Width           =   2295
      End
      Begin VB.CheckBox chk 
         Caption         =   "预览、打印时同一页相同日期显示一次"
         Height          =   180
         Index           =   4
         Left            =   300
         TabIndex        =   13
         Top             =   2580
         Width           =   3540
      End
      Begin VB.ComboBox cboOperSing 
         Height          =   300
         ItemData        =   "frmTendPara.frx":0D0C
         Left            =   1950
         List            =   "frmTendPara.frx":0D0E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   540
         Width           =   2670
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1755
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "1"
         Top             =   1590
         Width           =   420
      End
      Begin VB.ComboBox cboNodule 
         Height          =   300
         ItemData        =   "frmTendPara.frx":0D10
         Left            =   1950
         List            =   "frmTendPara.frx":0D12
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   900
         Width           =   2670
      End
      Begin VB.CheckBox chk 
         Caption         =   "只在当前页中显示跨页数据（不勾两页均显示）"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   14
         Top             =   2865
         Width           =   4215
      End
      Begin VB.CheckBox chk 
         Caption         =   "住院病人同一时间需要记录多份护理文件"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   11
         Top             =   1965
         Width           =   3645
      End
      Begin VB.CheckBox chk 
         Caption         =   "护理文件页码按文件顺序编号"
         Height          =   180
         Index           =   3
         Left            =   300
         TabIndex        =   15
         Top             =   3165
         Width           =   3135
      End
      Begin VB.CheckBox chk 
         Caption         =   "预览、打印时签名人显示签名图片"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   12
         Top             =   2280
         Width           =   3645
      End
      Begin VB.ComboBox cboSinger 
         Height          =   300
         ItemData        =   "frmTendPara.frx":0D14
         Left            =   1950
         List            =   "frmTendPara.frx":0D16
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   2670
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   0
         Left            =   2160
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1575
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt(0)"
         BuddyDispid     =   196622
         BuddyIndex      =   0
         OrigLeft        =   2175
         OrigTop         =   930
         OrigRight       =   2430
         OrigBottom      =   1200
         Max             =   30
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblLineColor 
         AutoSize        =   -1  'True
         Caption         =   "小结标识颜色"
         Height          =   180
         Left            =   810
         TabIndex        =   18
         Top             =   1290
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "护士、签名列显示模式"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   3
         Top             =   600
         Width           =   1800
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "允许录入超过当前        天的护理记录数据"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   8
         Top             =   1635
         Width           =   3600
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "小结缺省标识"
         Height          =   180
         Index           =   1
         Left            =   810
         TabIndex        =   5
         Top             =   975
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审签模式"
         Height          =   180
         Index           =   0
         Left            =   1170
         TabIndex        =   1
         Top             =   240
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmTendPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmMain As Object
Private mblnOK As Boolean
Private mstrPrivs As String

Private mvarColor As OLE_COLOR
Private Const tomAutoColor As Long = -9999997
'设定一个窗体捕获鼠标，即所有鼠标输入消息都发往该窗体
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
'取消鼠标捕获
Private Declare Function ReleaseCapture Lib "user32" () As Long

Public Function ShowPara(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strSQL As String, strPar As String
    Dim curDate As Date, intDay As Integer
    Dim intStart As Integer
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    '初始体温单标记
    '------------------------------------------------------------------------------------------------------------------
    
    '43588,刘鹏飞,2012-09-13,添加记录单审签模式
    cboSinger.Clear
    cboSinger.AddItem "0-聘任职务+审签权限"
    cboSinger.AddItem "1-审签权限"
    
    cboNodule.Clear
    cboNodule.AddItem "0-不处理"
    cboNodule.AddItem "1-上下画横线标识"
    cboNodule.AddItem "2-汇总值下方画双横线标识"
    cboNodule.AddItem "3-上方画横线标识"
    '72664:刘鹏飞,2014-07-18,添加小结标识
    cboNodule.AddItem "4-汇总值下方画单横线标识"
    
    '58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
    cboOperSing.Clear
    cboOperSing.AddItem "0-所有行显示"
    cboOperSing.AddItem "1-首行显示"
    cboOperSing.AddItem "2-首尾行显示"
    cboOperSing.AddItem "3-尾行显示"
    
    '43588:刘鹏飞,2012-09-13,添加记录单审签模式
    strTmp = Val(zlDatabase.GetPara("记录单审签模式", glngSys, 1255, "0", Array(cboSinger, lbl(0)), InStr(mstrPrivs, "护理选项设置") > 0))
    If Val(strTmp) >= 0 And Val(strTmp) <= 1 Then
        cboSinger.ListIndex = CInt(Val(strTmp))
    Else
        cboSinger.ListIndex = 0
    End If
    
    strTmp = zlDatabase.GetPara("小结缺省格式", glngSys, 1255, "0", Array(cboNodule, lbl(1)), InStr(mstrPrivs, "护理选项设置") > 0)
    If Val(strTmp) >= 0 And Val(strTmp) <= 4 Then
        cboNodule.ListIndex = Val(strTmp)
    Else
        cboNodule.ListIndex = 0
    End If
    
    '58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
    strTmp = Val(zlDatabase.GetPara("护士、签名列显示模式", glngSys, 1255, "2", Array(cboOperSing, lbl(3)), InStr(mstrPrivs, "护理选项设置") > 0))
    If Val(strTmp) >= 0 And Val(strTmp) <= 3 Then
        cboOperSing.ListIndex = CInt(Val(strTmp))
    Else
        cboOperSing.ListIndex = 2
    End If
    
    txt(0).Text = Val(zlDatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1", Array(txt(0), lbl(2)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(0).Value = Val(zlDatabase.GetPara("对应多份护理文件", glngSys, 1255, "0", Array(chk(0)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(1).Value = Val(zlDatabase.GetPara("记录单签名人显示方式", glngSys, 1255, "0", Array(chk(1)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(2).Value = Val(zlDatabase.GetPara("跨页数据只显示在第一页", glngSys, 1255, "0", Array(chk(2)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(3).Tag = 1
    chk(3).Value = Val(zlDatabase.GetPara("护理文件页码规则", glngSys, 1255, "0", Array(chk(3)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(3).Tag = 0
    '64583:刘鹏飞,2013-09-22,预览、打印时同一页相同日期显示方式:多次;一次
    chk(4).Value = Val(zlDatabase.GetPara("记录单日期显示方式", glngSys, 1255, "0", Array(chk(4)), InStr(mstrPrivs, "护理选项设置") > 0))
    '68739:刘鹏飞,2014-1-2,添加"小结标识颜色"
    picLineColor.BackColor = Val(zlDatabase.GetPara("小结标识颜色", glngSys, 1255, "255", Array(lblLineColor), InStr(mstrPrivs, "护理选项设置") > 0))
    
    Me.Show 1, mfrmMain
    ShowPara = mblnOK
    
End Function

Private Sub cboNodule_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub cboOperSing_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboSinger_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk_Click(Index As Integer)
    Dim strInfo As String
    If Not Index = 3 Then Exit Sub
    If Val(chk(Index).Tag) = 1 Then chk(Index).Tag = 0: Exit Sub
    strInfo = "此参数直接影响着记录单护理文件的页码编号规则，对于以下两种情况文件的页码将按照调整后的规则编号。"
    strInfo = strInfo & vbCrLf & "1、病人记录单文件份数小于等于1，后续建立了新的记录单文件。"
    strInfo = strInfo & vbCrLf & "2、病人所有记录单文件之间设置了合并打印，后续取消了某份被合并的文件。"
    strInfo = strInfo & vbCrLf & "请问您是否继续？"
    If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        chk(Index).Tag = 1
        chk(Index).Value = IIf(chk(Index).Value = 0, 1, 0)
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intStart As Integer
    Dim strTmp As String
    Dim lngColor As Long
    
    '43588:刘鹏飞,2012-09-13,添加记录单审签模式
    Call zlDatabase.SetPara("记录单审签模式", Val(cboSinger.ListIndex), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("小结缺省格式", Val(cboNodule.ListIndex), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("超期录入护理数据天数", Val(txt(0).Text), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("对应多份护理文件", chk(0).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("记录单签名人显示方式", chk(1).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("跨页数据只显示在第一页", chk(2).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("护理文件页码规则", chk(3).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    '58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
    Call zlDatabase.SetPara("护士、签名列显示模式", Val(cboOperSing.ListIndex), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    '64583:刘鹏飞,2013-09-22,预览、打印时同一页相同日期显示方式:多次;一次
    Call zlDatabase.SetPara("记录单日期显示方式", chk(4).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    '68739:刘鹏飞,2014-1-2,添加"小结标识颜色"
    Call zlDatabase.SetPara("小结标识颜色", Val(picLineColor.BackColor), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    
    
    mblnOK = True
    
    Unload Me
End Sub

Private Sub cmdUnVisible_Click()
    picControl.Visible = False
    If picLineColor.Enabled And picLineColor.Visible Then picLineColor.SetFocus
End Sub

Private Sub picLineColor_Click()
    picControl.Top = picLineColor.Top + picLineColor.Height
    picControl.Left = picLineColor.Left
    picControl.Visible = True
    Call SetCOLOR(Val(picLineColor.BackColor))
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt(Index))
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub PicColorCollect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 0 And X < PicColorCollect.ScaleWidth And Y > 0 And Y < PicColorCollect.ScaleHeight Then
        SetCapture PicColorCollect.hWnd
        shpBorder.Visible = True
    Else
        ReleaseCapture
        shpBorder.Visible = False
    End If

    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = Y \ (18 * Screen.TwipsPerPixelY)
    lCol = X \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    shpBorder.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    
    If PicColorCollect.Point(lX, lY) = -1 Then Exit Sub
    picColor.BackColor = PicColorCollect.Point(lX, lY)
    Select Case CStr(Hex(picColor.BackColor))
    Case "0"
        lblColor = "黑色"
    Case "3399"
        lblColor = "褐色"
    Case "3333"
        lblColor = "橄榄色"
    Case "3300"
        lblColor = "深绿"
    Case "663300"
        lblColor = "深青"
    Case "800000"
        lblColor = "深蓝"
    Case "993333"
        lblColor = "靛蓝"
    Case "333333"
        lblColor = "灰色-80%"
    Case "80"
        lblColor = "深红"
    Case "66FF"
        lblColor = "橙色"
    Case "8080"
        lblColor = "深黄"
    Case "8000"
        lblColor = "绿色"
    Case "808000"
        lblColor = "青色"
    Case "FF0000"
        lblColor = "蓝色"
    Case "996666"
        lblColor = "蓝-灰"
    Case "808080"
        lblColor = "灰色-50%"
    Case "FF"
        lblColor = "红色"
    Case "99FF"
        lblColor = "浅橙色"
    Case "CC99"
        lblColor = "酸橙色"
    Case "669933"
        lblColor = "海绿"
    Case "CCCC33"
        lblColor = "水绿色"
    Case "FF6633"
        lblColor = "浅蓝"
    Case "800080"
        lblColor = "紫罗兰"
    Case "999999"
        lblColor = "灰色-40%"
    Case "FF00FF"
        lblColor = "粉红"
    Case "CCFF"
        lblColor = "金色"
    Case "FFFF"
        lblColor = "黄色"
    Case "FF00"
        lblColor = "鲜绿"
    Case "FFFF00"
        lblColor = "青绿"
    Case "FFCC00"
        lblColor = "天蓝"
    Case "663399"
        lblColor = "梅红"
    Case "C0C0C0"
        lblColor = "灰色-25%"
    Case "CC99FF"
        lblColor = "玫瑰红"
    Case "99CCFF"
        lblColor = "茶色"
    Case "99FFFF"
        lblColor = "浅黄"
    Case "CCFFCC"
        lblColor = "浅绿"
    Case "FFFFCC"
        lblColor = "浅青绿"
    Case "FFCC99"
        lblColor = "淡蓝"
    Case "FF99CC"
        lblColor = "淡紫"
    Case "FFFFFF"
        lblColor = "白色"
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
    End Select
End Sub

Private Sub PicColorCollect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = Y \ (18 * Screen.TwipsPerPixelY)
    lCol = X \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    
    '按指定颜色作图
    picLineColor.BackColor = picColor.BackColor
    picControl.Visible = False
    If picLineColor.Enabled And picLineColor.Visible Then picLineColor.SetFocus
End Sub

Private Sub SetCOLOR(vData As OLE_COLOR)
    mvarColor = vData
    Dim lRow As Long, lCol As Long
    shpValue.Visible = True
    Select Case CStr(Hex(vData))
    Case "0"
        lblColor = "黑色"
        lRow = 0
        lCol = 0
    Case "3399"
        lblColor = "褐色"
        lRow = 0
        lCol = 1
    Case "3333"
        lblColor = "橄榄色"
        lRow = 0
        lCol = 2
    Case "3300"
        lblColor = "深绿"
        lRow = 0
        lCol = 3
    Case "663300"
        lblColor = "深青"
        lRow = 0
        lCol = 4
    Case "800000"
        lblColor = "深蓝"
        lRow = 0
        lCol = 5
    Case "993333"
        lblColor = "靛蓝"
        lRow = 0
        lCol = 6
    Case "333333"
        lblColor = "灰色-80%"
        lRow = 0
        lCol = 7
    Case "80"
        lblColor = "深红"
        lRow = 1
        lCol = 0
    Case "66FF"
        lblColor = "橙色"
        lRow = 1
        lCol = 1
    Case "8080"
        lblColor = "深黄"
        lRow = 1
        lCol = 2
    Case "8000"
        lblColor = "绿色"
        lRow = 1
        lCol = 3
    Case "808000"
        lblColor = "青色"
        lRow = 1
        lCol = 4
    Case "FF0000"
        lblColor = "蓝色"
        lRow = 1
        lCol = 5
    Case "996666"
        lblColor = "蓝-灰"
        lRow = 1
        lCol = 6
    Case "808080"
        lblColor = "灰色-50%"
        lRow = 1
        lCol = 7
    Case "FF"
        lblColor = "红色"
        lRow = 2
        lCol = 0
    Case "99FF"
        lblColor = "浅橙色"
        lRow = 2
        lCol = 1
    Case "CC99"
        lblColor = "酸橙色"
        lRow = 2
        lCol = 2
    Case "669933"
        lblColor = "海绿"
        lRow = 2
        lCol = 3
    Case "CCCC33"
        lblColor = "水绿色"
        lRow = 2
        lCol = 4
    Case "FF6633"
        lblColor = "浅蓝"
        lRow = 2
        lCol = 5
    Case "800080"
        lblColor = "紫罗兰"
        lRow = 2
        lCol = 6
    Case "999999"
        lblColor = "灰色-40%"
        lRow = 2
        lCol = 7
    Case "FF00FF"
        lblColor = "粉红"
        lRow = 3
        lCol = 0
    Case "CCFF"
        lblColor = "金色"
        lRow = 3
        lCol = 1
    Case "FFFF"
        lblColor = "黄色"
        lRow = 3
        lCol = 2
    Case "FF00"
        lblColor = "鲜绿"
        lRow = 3
        lCol = 3
    Case "FFFF00"
        lblColor = "青绿"
        lRow = 3
        lCol = 4
    Case "FFCC00"
        lblColor = "天蓝"
        lRow = 3
        lCol = 5
    Case "663399"
        lblColor = "梅红"
        lRow = 3
        lCol = 6
    Case "C0C0C0"
        lblColor = "灰色-25%"
        lRow = 3
        lCol = 7
    Case "CC99FF"
        lblColor = "玫瑰红"
        lRow = 4
        lCol = 0
    Case "99CCFF"
        lblColor = "茶色"
        lRow = 4
        lCol = 1
    Case "99FFFF"
        lblColor = "浅黄"
        lRow = 4
        lCol = 2
    Case "CCFFCC"
        lblColor = "浅绿"
        lRow = 4
        lCol = 3
    Case "FFFFCC"
        lblColor = "浅青绿"
        lRow = 4
        lCol = 4
    Case "FFCC99"
        lblColor = "淡蓝"
        lRow = 4
        lCol = 5
    Case "FF99CC"
        lblColor = "淡紫"
        lRow = 4
        lCol = 6
    Case "FFFFFF"
        lblColor = "白色"
        lRow = 4
        lCol = 7
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
    End Select
    shpBorder.Visible = False
    shpValue.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    shpValue.Visible = True
    If vData = tomAutoColor Or vData = -1 Then
    
    Else
        picColor.BackColor = vData
    End If
End Sub

