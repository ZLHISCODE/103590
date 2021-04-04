VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   Caption         =   "位图编辑器"
   ClientHeight    =   7050
   ClientLeft      =   2190
   ClientTop       =   4665
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   9780
   Begin VB.PictureBox picFinal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   6300
      ScaleHeight     =   915
      ScaleWidth      =   960
      TabIndex        =   3
      Top             =   3735
      Visible         =   0   'False
      Width           =   960
   End
   Begin zlPictureEditor.ucCanvas Canvas 
      Height          =   3885
      Left            =   495
      TabIndex        =   2
      Top             =   360
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   6853
   End
   Begin zlPictureEditor.Progress Progress 
      Height          =   375
      Left            =   4500
      TabIndex        =   0
      Top             =   5220
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   6720
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   4780
            MinWidth        =   4039
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1058
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "fMain.frx":038A
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   7065
      Top             =   1035
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'-- 位图控制
Public WithEvents DIBFilter As cDIBFilter   ' DIB 滤镜对象(24 bpp)
Attribute DIBFilter.VB_VarHelpID = -1
Public WithEvents DIBDither As cDIBDither   ' DIB 抖动对象(1, 4, 8 bpp)
Attribute DIBDither.VB_VarHelpID = -1
Public DIBPal               As New cDIBPal  ' DIB 调色板对象 (1, 4, 8 bpp)
Public DIBSave              As New cDIBSave ' Save 对象 (BMP)  (1, 4, 8, 24 bpp)
Attribute DIBSave.VB_VarHelpID = -1
Public DIBbpp               As Byte         ' 当前颜色深度

'-- 撤销重做控制
Private Const m_UNDO_LEVELS As Long = 25    ' 最大 Undo 数目
Private m_AppID             As Long         ' 程序ID (gfrmMain.hwnd)
Private m_UndoPos           As Long         ' 当前 Undo 位置
Private m_UndoMax           As Long         ' 最大可撤销数目
Private m_Temp              As String       ' 临时文件夹

'-- 对话框
Private m_LastFilter        As Integer      ' 最后使用的滤镜 (滤镜浏览器)
Private m_LastFilename      As String       ' 当前文件
Private m_LastPath          As String       ' 当前路径
Private m_Saved             As Boolean      ' 已保存
Private m_FileExt           As String       ' 当前文件扩展名
Private m_DialogPreview     As Boolean      ' 对话框: 显示预览
Private m_DialogFitMode     As Boolean      ' 对话框: 适合模式
Private m_DialogJPEGquality As Integer      ' 对话框: JPEG 质量 (0-100)

'-- GDI+
Private m_GDIpToken         As Long         ' 用于关闭 GDI+

Private cbp文件 As CommandBarPopup
Private cbp编辑 As CommandBarPopup
Private cbp缩放 As CommandBarPopup
Private cbp颜色 As CommandBarPopup
Private cbp滤镜 As CommandBarPopup
Private cbp视图 As CommandBarPopup
Private cbp帮助 As CommandBarPopup

Private Bar标准 As CommandBar

Private mblnModeless As Boolean

Public Event pOK(ByRef FinalPicture As StdPicture, ByVal lngWidth As Long, ByVal lngHeight As Long)    '保存，返回修改后的临时图片路径（JPEG图片）
Public Event pCancel()                      '取消并退出

'################################################################################################################
'## 功能：  打开并显示表格编辑器窗体
'##
'## 参数：  srcPic      :In     源图片
'##         frmParent   :In     父窗体
'##         blnModeless :In     是否是非模态，默认为非模态
'################################################################################################################
Public Sub ShowMe(ByRef srcPic As StdPicture, Optional ByRef frmParent As Object, Optional ByVal blnModeless As Boolean = True)
    If srcPic = 0 Then
        Unload Me
    Else
        '-- Create DIB
        DoEvents
        Screen.MousePointer = vbHourglass
        Call pvSetDIBPicture(srcPic)
        Screen.MousePointer = vbNormal

        '-- Reset Undo/Redo and save first Undo
        Call pvClearAllDIB
        Call pvSaveUndoDIB
        '-- Save info
        m_LastFilename = "[内部图片]"
        Call RefreshFileInfo
    End If
    mblnModeless = blnModeless
    Me.Show IIf(blnModeless, vbModeless, vbModal), frmParent
End Sub

Private Sub InitMenus()
    Dim i As Long, j As Long
'    '## 窗体位置恢复
'    Me.Left = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "MainLeft", (Screen.Width - 12000) / 2)
'    Me.Top = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "MainTop", (Screen.Height - 9000) / 2)
'    Me.Width = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "MainWidth", 12000)
'    Me.Height = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "MainHeight", 9000)

    '## 菜单初始化
    Dim cbpPopup As CommandBarPopup                     '临时对象
    Dim cbpPopupSub As CommandBarPopup                  '临时对象
    Dim objControl As CommandBarControl                 '工具栏控件
    Dim objCustControl As CommandBarControlCustom       '自定义控件
    Dim Combo As CommandBarComboBox                     '工具栏下拉框控件
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBars.Icons = ImageManager.Icons
    
    CommandBars.ActiveMenuBar.Title = "菜单栏"
    
    Set cbp文件 = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "文件(&F)")
    With cbp文件.CommandBar.Controls
        .Add xtpControlButton, ID_FILE_OPEN, "打开(&O)..."
        .Add xtpControlButton, ID_FILE_SAVE, "保存(&S)"
        .Add xtpControlButton, ID_FILE_SAVEAS, "另存为(&A)..."
        
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "打印(&P)...")
        objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "退出(&Q)")
        objControl.BeginGroup = True
    End With
    
    Set cbp编辑 = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "编辑(&E)")
    With cbp编辑.CommandBar.Controls
        .Add xtpControlButton, ID_EDIT_UNDO, "撤销(&U)"
        .Add xtpControlButton, ID_EDIT_REDO, "重做(&R)"
        
        Set objControl = .Add(xtpControlButton, ID_EDIT_COPY, "复制(&C)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_PASTE, "粘贴(&P)"
        
        Set objControl = .Add(xtpControlButton, ID_EDIT_SIZE, "调整图像尺寸(&S)...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_ORIENT, "调整画布方向(&O)..."
    
        Set objControl = .Add(xtpControlButton, ID_EDIT_SCROLLMODE, "卷动模式(&L)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_CROPMODE, "剪切模式(&O)"
    End With
    
    Set cbp缩放 = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "缩放(&Z)")
    With cbp缩放.CommandBar.Controls
        .Add xtpControlButton, ID_ZOOM_IN, "放大(&I)"
        .Add xtpControlButton, ID_ZOOM_OUT, "缩小(&O)"
        .Add xtpControlButton, ID_ZOOM_11, "实际尺寸(&A)"
        
        Set objControl = .Add(xtpControlButton, ID_ZOOM_FIT, "适合窗口(&F)")
        objControl.BeginGroup = True
    End With
    
    Set cbp颜色 = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "颜色(&C)")
    With cbp颜色.CommandBar.Controls
        .Add xtpControlButton, ID_COLOR_BLACKWHITE, "灰度-黑白"
        .Add xtpControlButton, ID_COLOR_GREYS16, "灰度-16色"
        .Add xtpControlButton, ID_COLOR_GREYS256, "灰度-256色"
        
        Set objControl = .Add(xtpControlButton, ID_COLOR_COLOR2, "彩色-2色")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_COLOR_COLOR16, "彩色-16色"
        .Add xtpControlButton, ID_COLOR_COLOR256, "彩色-256色"
        
        Set objControl = .Add(xtpControlButton, ID_COLOR_TRUECOLOR, "真彩色")
        objControl.BeginGroup = True
    End With
    
    Set cbp滤镜 = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "滤镜(&L)")
    With cbp滤镜.CommandBar.Controls
        .Add xtpControlButton, ID_ADJUST_BRIGHT, "亮度(&B)"
        .Add xtpControlButton, ID_ADJUST_CONTRAST, "对比度(&C)"
        .Add xtpControlButton, ID_ADJUST_SITUATION, "饱和度(&S)"
    
        Set cbpPopup = .Add(xtpControlPopup, 0, "颜色(&C)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_COLOR1, "灰度(&G)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_COLOR2, "负片效果(&N)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_COLOR3, "老照片(&O)"
        
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FILTER_COLOR4, "颜色填充(&C)...")
        objControl.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_COLOR5, "替换 &HS..."
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_COLOR6, "替换 &L..."
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_COLOR7, "曝光过度(&M)..."
    
        Set cbpPopup = .Add(xtpControlPopup, 0, "清晰度(&D)")
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_DEF1, "模糊(&B)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_DEF2, "柔化(&F)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_DEF3, "锐化(&S)"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FILTER_DEF4, "扩散(&D)")
        objControl.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_DEF5, "象素化(&P)"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FILTER_DEF6, "去斑(&K)")
        objControl.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_DEF7, "进一步去斑(&M)"
    
        Set cbpPopup = .Add(xtpControlPopup, 0, "边缘(&E)")
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_EDGES1, "照亮边缘(&C)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_EDGES2, "浮雕效果(&E)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_EDGES3, "墨水轮廓(&O)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_EDGES4, "版画(&R)"
    
        Set cbpPopup = .Add(xtpControlPopup, 0, "特殊(&S)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_SPECIAL1, "噪音(&N)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_SPECIAL2, "扫描线(&S)"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FILTER_SPECIAL3, "扩张(&D)")
        objControl.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILTER_SPECIAL4, "腐蚀(&E)"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FILTER_SPECIAL5, "纹理(&T)...")
        objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, ID_ADJUST_FILTERBROWSER, "所有滤镜(&I)...")
        objControl.BeginGroup = True
    End With

    Set cbp视图 = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "视图(&V)")
    With cbp视图.CommandBar.Controls
        .Add xtpControlButton, ID_VIEW_TOOLBARLIST, "工具栏列表"
        .Add xtpControlButton, ID_VIEW_PANORAMIC, "缩略图(&V)"
        .Add xtpControlButton, ID_VIEW_PROPERTY, "属性(&P)"
    End With
    
    Set cbp帮助 = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "帮助(&H)")
    With cbp帮助.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "帮助主题(&H)")
        objControl.BeginGroup = True
        Set cbpPopupSub = .Add(xtpControlPopup, 0, "&Web上的医业")
        objControl.BeginGroup = True
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_HELP_ONLINE, "医业在线(&H)"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_HELP_CONTACT, "发送反馈(&M)"
        Set objControl = .Add(xtpControlButton, ID_HELP_ABOUT, "关于(&A)...")
        objControl.BeginGroup = True
    End With
    
    '## 工具栏初始化
    
    Set Bar标准 = CommandBars.Add("标准", xtpBarTop)
    With Bar标准.Controls
        Set objControl = .Add(xtpControlButton, ID_FILE_OPEN, "打开")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_FILE_SAVE, "保存"
        .Add xtpControlButton, ID_FILE_PRINT, "打印"
        
        Set objControl = .Add(xtpControlButton, ID_EDIT_UNDO, "撤销")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_REDO, "重做"
        
        Set objControl = .Add(xtpControlButton, ID_EDIT_SCROLLMODE, "卷动模式")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_CROPMODE, "剪切模式"
        
        Set objControl = .Add(xtpControlButton, ID_ZOOM_IN, "放大")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_ZOOM_OUT, "缩小"
        .Add xtpControlButton, ID_ZOOM_11, "实际尺寸"
        
        Set objControl = .Add(xtpControlButton, ID_ZOOM_FIT, "适合窗口")
        objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, ID_EDIT_SIZE, "调节尺寸")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_ORIENT, "调节方向"
    
        Set objControl = .Add(xtpControlButton, ID_FILTER_DEF3, "锐化")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_FILTER_DEF2, "柔化"
        .Add xtpControlButton, ID_FILTER_DEF6, "去斑"
    
        Set objControl = .Add(xtpControlButton, ID_ADJUST_BRIGHT, "亮度")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_ADJUST_CONTRAST, "对比度"
        .Add xtpControlButton, ID_ADJUST_SITUATION, "饱和度"
    
        Set objControl = .Add(xtpControlButton, ID_ADJUST_FILTERBROWSER, "所有滤镜...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_FILTER_SPECIAL5, "纹理..."
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE, "保存(&S)")
        objControl.BeginGroup = True
        objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "退出(&Q)")
        objControl.Style = xtpButtonIconAndCaption
    End With
        
    CommandBars.KeyBindings.Add FCONTROL, Asc("O"), ID_FILE_OPEN
    CommandBars.KeyBindings.Add FCONTROL, Asc("S"), ID_FILE_SAVE
    CommandBars.KeyBindings.Add FCONTROL, Asc("A"), ID_FILE_SAVEAS
    CommandBars.KeyBindings.Add FCONTROL, Asc("P"), ID_FILE_PRINT
    CommandBars.KeyBindings.Add FCONTROL, Asc("Z"), ID_EDIT_UNDO
    CommandBars.KeyBindings.Add FCONTROL, Asc("Y"), ID_EDIT_REDO
    CommandBars.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
    CommandBars.KeyBindings.Add FCONTROL, Asc("V"), ID_EDIT_PASTE
    CommandBars.KeyBindings.Add FCONTROL, Asc("Q"), ID_FILE_EXIT
    CommandBars.KeyBindings.Add FALT, Asc("Q"), ID_FILE_EXIT
    
    CommandBars.KeyBindings.Add 0, VK_F1, ID_HELP_CONTENT
    CommandBars.KeyBindings.Add 0, VK_F2, ID_ADJUST_BRIGHT
    CommandBars.KeyBindings.Add 0, VK_F3, ID_ADJUST_CONTRAST
    CommandBars.KeyBindings.Add 0, VK_F4, ID_ADJUST_SITUATION
    CommandBars.KeyBindings.Add 0, VK_F6, ID_VIEW_PANORAMIC
    CommandBars.KeyBindings.Add 0, VK_F8, ID_VIEW_PROPERTY
    CommandBars.KeyBindings.Add 0, VK_F12, ID_ADJUST_FILTERBROWSER
    
    '显示扩展按钮
    CommandBars.Options.ShowExpandButtonAlways = True
    CommandBars.EnableCustomization (True)
    CommandBars.Options.UseDisabledIcons = True
End Sub

'################################################################################################################
'## 功能：  添加按钮
'################################################################################################################
Public Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, ID As Long, Caption As String, Optional BeginGroup As Boolean = False, Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional Category As String = "Controls") As CommandBarControl
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, ID, Caption)
    
    Control.BeginGroup = BeginGroup
    Control.DescriptionText = DescriptionText
    Control.Style = ButtonStyle
    Control.Category = Category

    Set AddButton = Control
End Function

Private Sub Canvas_DIBProgressStart()
    BeginShowProgress
End Sub

Private Sub CommandBars_Customization(ByVal Options As XtremeCommandBars.ICustomizeOptions)
    Dim Controls As CommandBarControls
    Set Controls = CommandBars.DesignerControls
    
    If (Controls.Count = 0) Then
        AddButton Controls, xtpControlButton, ID_FILE_OPEN, "打开", , "打开", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_SAVE, "保存", , "保存", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_SAVEAS, "另存为", , "另存为", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_PRINT, "打印", , "打印", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_EXIT, "退出", , "退出", xtpButtonAutomatic, "文件"

        AddButton Controls, xtpControlButton, ID_EDIT_UNDO, "撤销", , "撤销", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_REDO, "重做", , "重做", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_COPY, "复制", , "复制", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_PASTE, "粘贴", , "粘贴", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_SIZE, "调整图像尺寸...", , "调整图像尺寸...", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_ORIENT, "调整画布方向...", , "调整画布方向...", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_SCROLLMODE, "卷动模式", , "卷动模式", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_CROPMODE, "剪切模式", , "剪切模式", xtpButtonAutomatic, "编辑"

        AddButton Controls, xtpControlButton, ID_ZOOM_IN, "放大", , "放大", xtpButtonAutomatic, "缩放"
        AddButton Controls, xtpControlButton, ID_ZOOM_OUT, "缩小", , "缩小", xtpButtonAutomatic, "缩放"
        AddButton Controls, xtpControlButton, ID_ZOOM_11, "实际尺寸", , "实际尺寸", xtpButtonAutomatic, "缩放"
        AddButton Controls, xtpControlButton, ID_ZOOM_FIT, "适合窗口", , "适合窗口", xtpButtonAutomatic, "缩放"

        AddButton Controls, xtpControlButton, ID_COLOR_BLACKWHITE, "灰度-黑白", , "灰度-黑白", xtpButtonAutomatic, "颜色"
        AddButton Controls, xtpControlButton, ID_COLOR_GREYS16, "灰度-16色", , "灰度-16色", xtpButtonAutomatic, "颜色"
        AddButton Controls, xtpControlButton, ID_COLOR_GREYS256, "灰度-256色", , "灰度-256色", xtpButtonAutomatic, "颜色"
        AddButton Controls, xtpControlButton, ID_COLOR_COLOR2, "彩色-2色", , "彩色-2色", xtpButtonAutomatic, "颜色"
        AddButton Controls, xtpControlButton, ID_COLOR_COLOR16, "彩色-16色", , "彩色-16色", xtpButtonAutomatic, "颜色"
        AddButton Controls, xtpControlButton, ID_COLOR_COLOR256, "彩色-256色", , "彩色-256色", xtpButtonAutomatic, "颜色"
        AddButton Controls, xtpControlButton, ID_COLOR_TRUECOLOR, "真彩色", , "真彩色", xtpButtonAutomatic, "颜色"

        AddButton Controls, xtpControlButton, ID_ADJUST_BRIGHT, "亮度", , "亮度", xtpButtonAutomatic, "调节"
        AddButton Controls, xtpControlButton, ID_ADJUST_CONTRAST, "对比度", , "对比度", xtpButtonAutomatic, "调节"
        AddButton Controls, xtpControlButton, ID_ADJUST_SITUATION, "饱和度", , "饱和度", xtpButtonAutomatic, "调节"
        AddButton Controls, xtpControlButton, ID_ADJUST_FILTERBROWSER, "所有滤镜...", , "所有滤镜...", xtpButtonAutomatic, "调节"

        AddButton Controls, xtpControlButton, ID_FILTER_COLOR1, "灰度", , "灰度", xtpButtonAutomatic, "滤镜－颜色"
        AddButton Controls, xtpControlButton, ID_FILTER_COLOR2, "负片效果", , "负片效果", xtpButtonAutomatic, "滤镜－颜色"
        AddButton Controls, xtpControlButton, ID_FILTER_COLOR3, "老照片", , "老照片", xtpButtonAutomatic, "滤镜－颜色"
        AddButton Controls, xtpControlButton, ID_FILTER_COLOR4, "颜色填充...", , "颜色填充...", xtpButtonAutomatic, "滤镜－颜色"
        AddButton Controls, xtpControlButton, ID_FILTER_COLOR5, "替换 HS...", , "替换 HS...", xtpButtonAutomatic, "滤镜－颜色"
        AddButton Controls, xtpControlButton, ID_FILTER_COLOR6, "替换 L...", , "替换 L...", xtpButtonAutomatic, "滤镜－颜色"
        AddButton Controls, xtpControlButton, ID_FILTER_COLOR7, "曝光过度...", , "曝光过度...", xtpButtonAutomatic, "滤镜－颜色"

        AddButton Controls, xtpControlButton, ID_FILTER_DEF1, "模糊", , "模糊", xtpButtonAutomatic, "滤镜－清晰度"
        AddButton Controls, xtpControlButton, ID_FILTER_DEF2, "柔化", , "柔化", xtpButtonAutomatic, "滤镜－清晰度"
        AddButton Controls, xtpControlButton, ID_FILTER_DEF3, "锐化", , "锐化", xtpButtonAutomatic, "滤镜－清晰度"
        AddButton Controls, xtpControlButton, ID_FILTER_DEF4, "扩散", , "扩散", xtpButtonAutomatic, "滤镜－清晰度"
        AddButton Controls, xtpControlButton, ID_FILTER_DEF5, "象素化", , "象素化", xtpButtonAutomatic, "滤镜－清晰度"
        AddButton Controls, xtpControlButton, ID_FILTER_DEF6, "去斑", , "去斑", xtpButtonAutomatic, "滤镜－清晰度"
        AddButton Controls, xtpControlButton, ID_FILTER_DEF7, "进一步去斑", , "进一步去斑", xtpButtonAutomatic, "滤镜－清晰度"

        AddButton Controls, xtpControlButton, ID_FILTER_EDGES1, "照亮边缘", , "照亮边缘", xtpButtonAutomatic, "滤镜－边缘"
        AddButton Controls, xtpControlButton, ID_FILTER_EDGES2, "浮雕效果", , "浮雕效果", xtpButtonAutomatic, "滤镜－边缘"
        AddButton Controls, xtpControlButton, ID_FILTER_EDGES3, "墨水轮廓", , "墨水轮廓", xtpButtonAutomatic, "滤镜－边缘"
        AddButton Controls, xtpControlButton, ID_FILTER_EDGES4, "版画", , "版画", xtpButtonAutomatic, "滤镜－边缘"

        AddButton Controls, xtpControlButton, ID_FILTER_SPECIAL1, "噪音", , "噪音", xtpButtonAutomatic, "滤镜－特殊"
        AddButton Controls, xtpControlButton, ID_FILTER_SPECIAL2, "扫描线", , "扫描线", xtpButtonAutomatic, "滤镜－特殊"
        AddButton Controls, xtpControlButton, ID_FILTER_SPECIAL3, "扩张", , "扩张", xtpButtonAutomatic, "滤镜－特殊"
        AddButton Controls, xtpControlButton, ID_FILTER_SPECIAL4, "腐蚀", , "腐蚀", xtpButtonAutomatic, "滤镜－特殊"
        AddButton Controls, xtpControlButton, ID_FILTER_SPECIAL5, "纹理...", , "纹理...", xtpButtonAutomatic, "滤镜－特殊"

        AddButton Controls, xtpControlButton, ID_VIEW_TOOLBARLIST, "工具栏列表", , "工具栏列表", xtpButtonAutomatic, "视图"
        AddButton Controls, xtpControlButton, ID_VIEW_PANORAMIC, "缩略图", , "缩略图", xtpButtonAutomatic, "视图"
        AddButton Controls, xtpControlButton, ID_VIEW_PROPERTY, "属性", , "属性", xtpButtonAutomatic, "视图"

        AddButton Controls, xtpControlButton, ID_HELP_CONTENT, "帮助主题", , "帮助主题", xtpButtonAutomatic, "帮助"
        AddButton Controls, xtpControlButton, ID_HELP_ONLINE, "医业在线", , "医业在线", xtpButtonAutomatic, "帮助"
        AddButton Controls, xtpControlButton, ID_HELP_CONTACT, "发送反馈", , "发送反馈", xtpButtonAutomatic, "帮助"
        AddButton Controls, xtpControlButton, ID_HELP_ABOUT, "关于...", , "关于...", xtpButtonAutomatic, "帮助"
    End If
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_FILE_OPEN
        '打开
        DoFileMenu 0
    Case ID_FILE_SAVE
        '保存
        DoFileMenu 1
    Case ID_FILE_SAVEAS
        '另存为
        DoFileMenu 2
    Case ID_FILE_PRINT
        '打印
        DoFileMenu 3
    Case ID_FILE_EXIT
        '退出
        DoFileMenu 4
    Case ID_EDIT_UNDO
        '撤销
        DoEditMenu 0
    Case ID_EDIT_REDO
        '重做
        DoEditMenu 1
    Case ID_EDIT_COPY
        '复制
        DoEditMenu 2
    Case ID_EDIT_PASTE
        '粘贴
        DoEditMenu 3
    Case ID_EDIT_SIZE
        '调整尺寸
        DoEditMenu 4
    Case ID_EDIT_ORIENT
        '调整方向
        DoEditMenu 5
    Case ID_EDIT_SCROLLMODE
        '卷动模式
        DoEditMenu 6
    Case ID_EDIT_CROPMODE
        '剪切模式
        DoEditMenu 7
    Case ID_ZOOM_IN
        '放大
        Bar标准.FindControl(, ID_ZOOM_FIT).Checked = False
        DoZoomMenu 0
    Case ID_ZOOM_OUT
        '缩小
        Bar标准.FindControl(, ID_ZOOM_FIT).Checked = False
        DoZoomMenu 1
    Case ID_ZOOM_11
        '1:1
        Bar标准.FindControl(, ID_ZOOM_FIT).Checked = False
        DoZoomMenu 2
    Case ID_ZOOM_FIT
        '适合
        DoZoomMenu 3
    Case ID_COLOR_BLACKWHITE
        '灰度-黑白
        DoColorMenu 0
    Case ID_COLOR_GREYS16
        '灰度-16色
        DoColorMenu 1
    Case ID_COLOR_GREYS256
        '灰度-256色
        DoColorMenu 2
    Case ID_COLOR_COLOR2
        '彩色-2色
        DoColorMenu 3
    Case ID_COLOR_COLOR16
        '彩色-16色
        DoColorMenu 4
    Case ID_COLOR_COLOR256
        '彩色-256色
        DoColorMenu 5
    Case ID_COLOR_TRUECOLOR
        '真彩色
        DoColorMenu 6
    Case ID_ADJUST_BRIGHT
        '亮度
        DoAdjustMenu 0
    Case ID_ADJUST_CONTRAST
        '对比度
        DoAdjustMenu 1
    Case ID_ADJUST_SITUATION
        '饱和度
        DoAdjustMenu 2
    Case ID_ADJUST_FILTERBROWSER
        '滤镜浏览器
        DoAdjustMenu 3
    Case ID_FILTER_COLOR1
        '颜色－灰度
        DoFilterColorMenu 0
    Case ID_FILTER_COLOR2
        '颜色－负片效果
        DoFilterColorMenu 1
    Case ID_FILTER_COLOR3
        '颜色－水粉画
        DoFilterColorMenu 2
    Case ID_FILTER_COLOR4
        '颜色－颜色填充
        DoFilterColorMenu 3
    Case ID_FILTER_COLOR5
        '颜色－替换 HS...
        DoFilterColorMenu 4
    Case ID_FILTER_COLOR6
        '颜色－替换 L...
        DoFilterColorMenu 5
    Case ID_FILTER_COLOR7
        '颜色－位移
        DoFilterColorMenu 6
    Case ID_FILTER_DEF1
        '清晰度－模糊
        DoFilterDefMenu 0
    Case ID_FILTER_DEF2
        '清晰度－柔化
        DoFilterDefMenu 1
    Case ID_FILTER_DEF3
        '清晰度－锐化
        DoFilterDefMenu 2
    Case ID_FILTER_DEF4
        '清晰度－扩散
        DoFilterDefMenu 3
    Case ID_FILTER_DEF5
        '清晰度－象素化
        DoFilterDefMenu 4
    Case ID_FILTER_DEF6
        '清晰度－去癍
        DoFilterDefMenu 5
    Case ID_FILTER_DEF7
        '清晰度－进一步去癍
        DoFilterDefMenu 6
    Case ID_FILTER_EDGES1
        '边缘－轮廓
        DoFilterEdgesMenu 0
    Case ID_FILTER_EDGES2
        '边缘－浮雕
        DoFilterEdgesMenu 1
    Case ID_FILTER_EDGES3
        '边缘－草图
        DoFilterEdgesMenu 2
    Case ID_FILTER_EDGES4
        '边缘－醒目
        DoFilterEdgesMenu 3
    Case ID_FILTER_SPECIAL1
        '特殊－噪音
        DoFilterSpecialMenu 0
    Case ID_FILTER_SPECIAL2
        '特殊－扫描线
        DoFilterSpecialMenu 1
    Case ID_FILTER_SPECIAL3
        '特殊－扩张
        DoFilterSpecialMenu 2
    Case ID_FILTER_SPECIAL4
        '特殊－腐蚀
        DoFilterSpecialMenu 3
    Case ID_FILTER_SPECIAL5
        '特殊－材质...
        DoFilterSpecialMenu 4
    Case ID_VIEW_PANORAMIC
        '缩略图
        DoViewMenu 1
    Case ID_VIEW_PROPERTY
        '属性
        DoViewMenu 2
    Case ID_HELP_CONTENT
        '帮助主题
        ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
    Case ID_HELP_CONTACT
        '发送反馈
        Call zlMailTo(Me.hwnd)
    Case ID_HELP_ONLINE
        '在线医业
        Call zlHomePage(Me.hwnd)
    Case ID_HELP_ABOUT
        '关于...
        ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
    End Select
End Sub

Private Sub CommandBars_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub CommandBars_Resize()
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    CommandBars.GetClientRect Left, Top, Right, Bottom
    If Right >= Left And Bottom >= Top Then
        Canvas.Move Left, Top, Right - Left, Bottom - Top
    Else
        Canvas.Move 0, 0, 0, 0
    End If
End Sub

Private Sub CommandBars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim bEnbBPP As Boolean
    Dim bEnbDIB As Boolean
    Dim lIdx    As Long
    
    bEnbBPP = (DIBbpp = 24)             '允许真彩色操作
    bEnbDIB = (Canvas.DIB.hDIB <> 0)    '允许DIB操作
    
    Select Case Control.ID
    Case ID_FILE_OPEN
        '打开
    Case ID_FILE_SAVE
        '保存
        Control.Enabled = bEnbDIB
    Case ID_FILE_SAVEAS
        '另存为
        Control.Enabled = bEnbDIB
    Case ID_FILE_PRINT
        '打印
        Control.Enabled = bEnbDIB
    Case ID_FILE_EXIT
        '退出
    Case ID_EDIT_UNDO
        '撤销
        Control.Enabled = (m_UndoPos > 1)
    Case ID_EDIT_REDO
        '重做
        Control.Enabled = (m_UndoPos <> m_UndoMax)
    Case ID_EDIT_COPY
        '复制
        Control.Enabled = (Canvas.DIB.hDIB <> 0)
    Case ID_EDIT_PASTE
        '粘贴
        Control.Enabled = (Clipboard.GetFormat(vbCFBitmap))
    Case ID_EDIT_SIZE
        '调整尺寸
        Control.Enabled = bEnbDIB
    Case ID_EDIT_ORIENT
        '调整方向
        Control.Enabled = bEnbDIB
    Case ID_EDIT_SCROLLMODE
        '卷动模式
        Control.Enabled = (bEnbDIB)
        Control.Checked = (Canvas.WorkMode = [cnvScrollMode])
    Case ID_EDIT_CROPMODE
        '剪切模式
        Control.Enabled = (bEnbDIB)
        Control.Checked = (Canvas.WorkMode = [cnvCropMode])
    Case ID_ZOOM_IN
        '放大
    Case ID_ZOOM_OUT
        '缩小
    Case ID_ZOOM_11
        '1:1
    Case ID_ZOOM_FIT
        '适合
        Control.Checked = (Canvas.FitMode = True)
    Case ID_COLOR_BLACKWHITE
        '灰度-黑白
        Control.Enabled = bEnbDIB
    Case ID_COLOR_GREYS16
        '灰度-16色
        Control.Enabled = bEnbDIB
    Case ID_COLOR_GREYS256
        '灰度-256色
        Control.Enabled = bEnbDIB
    Case ID_COLOR_COLOR2
        '彩色-2色
        Control.Enabled = bEnbDIB
    Case ID_COLOR_COLOR16
        '彩色-16色
        Control.Enabled = bEnbDIB
    Case ID_COLOR_COLOR256
        '彩色-256色
        Control.Enabled = bEnbDIB
    Case ID_COLOR_TRUECOLOR
        '真彩色
        Control.Enabled = bEnbDIB
    Case ID_ADJUST_BRIGHT
        '亮度
        Control.Enabled = bEnbBPP
    Case ID_ADJUST_CONTRAST
        '对比度
        Control.Enabled = bEnbBPP
    Case ID_ADJUST_SITUATION
        '饱和度
        Control.Enabled = bEnbBPP
    Case ID_ADJUST_FILTERBROWSER
        '滤镜浏览器
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR1
        '颜色－灰度
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR2
        '颜色－负片效果
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR3
        '颜色－水粉画
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR4
        '颜色－颜色填充
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR5
        '颜色－替换 HS...
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR6
        '颜色－替换 L...
        Control.Enabled = bEnbBPP
    Case ID_FILTER_COLOR7
        '颜色－位移
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF1
        '清晰度－模糊
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF2
        '清晰度－柔化
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF3
        '清晰度－锐化
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF4
        '清晰度－扩散
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF5
        '清晰度－象素化
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF6
        '清晰度－去癍
        Control.Enabled = bEnbBPP
    Case ID_FILTER_DEF7
        '清晰度－进一步去癍
        Control.Enabled = bEnbBPP
    Case ID_FILTER_EDGES1
        '边缘－轮廓
        Control.Enabled = bEnbBPP
    Case ID_FILTER_EDGES2
        '边缘－浮雕
        Control.Enabled = bEnbBPP
    Case ID_FILTER_EDGES3
        '边缘－草图
        Control.Enabled = bEnbBPP
    Case ID_FILTER_EDGES4
        '边缘－醒目
        Control.Enabled = bEnbBPP
    Case ID_FILTER_SPECIAL1
        '特殊－噪音
        Control.Enabled = bEnbBPP
    Case ID_FILTER_SPECIAL2
        '特殊－扫描线
        Control.Enabled = bEnbBPP
    Case ID_FILTER_SPECIAL3
        '特殊－扩张
        Control.Enabled = bEnbBPP
    Case ID_FILTER_SPECIAL4
        '特殊－腐蚀
        Control.Enabled = bEnbBPP
    Case ID_FILTER_SPECIAL5
        '特殊－材质...
        Control.Enabled = bEnbBPP
    Case ID_VIEW_PANORAMIC
        '缩略图
        Control.Checked = gfPanView.Visible
    Case ID_VIEW_PROPERTY
        '属性
        Control.Enabled = bEnbDIB
    Case ID_HELP_CONTENT
        '帮助主题
    Case ID_HELP_CONTACT
        '发送反馈
    Case ID_HELP_ONLINE
        '在线医业
    Case ID_HELP_ABOUT
        '关于...
    End Select
End Sub

Private Sub DIBDither_ProgressStart()
    BeginShowProgress
End Sub

Private Sub DIBFilter_ProgressStart()
    BeginShowProgress
End Sub

Private Sub Form_Activate()
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub

'========================================================================================
' 主程序
'========================================================================================
Private Sub Form_Load()
    Dim GpInput As GdiplusStartupInput
    '-- 调入 GDI+ Dll
    GpInput.GdiplusVersion = 1
    If (mGDIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
        Call MsgBox("调入 GDI+ 出错，无法进行位图编辑！请检查 GDI+ DLL 是否存在或者损坏！", vbInformation + vbOKOnly)
        Call Unload(Me)
        Exit Sub
    End If
    '-- 菜单初始化
    Call InitMenus
    '-- 恢复设置
    Call mSettings.LoadMainSettings
    
    '-- Initial zoom = 100%
    stbThis.Panels(3).Text = "100%"
    
    '-- Initialize 'evented' objects
    Set DIBFilter = New cDIBFilter
    Set DIBDither = New cDIBDither
    
    '-- Get App. ID and <Temp> path (Undo/Redo temp. files)
    m_AppID = Me.hwnd
    m_Temp = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))

'    '-- Hook wheel for zooming
'    Call mHook.HookWheel(Me.hwnd)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim sRet As VbMsgBoxResult
    
    If (Canvas.DIB.hDIB <> 0 And Not m_Saved And m_UndoPos > 1) Then
    
        sRet = MsgBox("是否在退出之前保存图片？", vbYesNoCancel Or vbInformation)
        Select Case sRet
            Case vbYes    '-- Save
                Call DoFileMenu(1)
                Cancel = 0
            Case vbNo     '-- Don't save
                Cancel = 0
            Case vbCancel '-- Cancel
                Cancel = 1
        End Select
    End If
    If (Cancel = 0) Then Call Unload(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Canvas.DIB.Destroy
    '-- Save settings
    Call mSettings.SaveMainSettings
    
    '-- Delete temp. files
    Call pvClearAllDIB
    
    ' Unload the GDI+ Dll
    Call mGDIpEx.GdiplusShutdown(m_GDIpToken)

    '-- Free objects
    Set DIBFilter = Nothing
    Set DIBDither = Nothing
    Set DIBPal = Nothing
    Set DIBSave = Nothing
    
    RaiseEvent pCancel
End Sub

'========================================================================================
' Processing
'========================================================================================

Public Sub Canvas_DIBProgress(ByVal p As Long)
    Progress.Value = CDbl(p) / Progress.Max
End Sub

Public Sub Canvas_DIBProgressEnd()

    '-- Progress end
    Progress.Value = 0
    Progress.Cls
    Progress.Visible = False
    Call gfPanView.Repaint
    '-- DIB processed (-> 24bpp: Size changed, orientation changed)
    DIBbpp = 24
    Call pvSetPalMode(DIBbpp)
    stbThis.Panels(2).Text = Canvas.DIB.Width & "×" & Canvas.DIB.Height & "×" & DIBbpp & "bpp"
    
    '-- Save Undo
    Call pvSaveUndoDIB
End Sub

Public Sub DIBFilter_Progress(ByVal p As Long)
    Progress.Value = CDbl(p) / Progress.Max
End Sub

Public Sub DIBFilter_ProgressEnd()

    '-- Progress end
    Progress.Value = 0
    Progress.Cls
    Progress.Visible = False
    Call Canvas.Repaint
    Call gfPanView.Repaint
    '-- If not previewing (Filter browse box), save Undo
    If (gfFilter.Previewing = False And gfTexturize.Previewing = False) Then Call pvSaveUndoDIB
End Sub

Public Sub DIBDither_Progress(ByVal p As Long)
    Progress.Value = CDbl(p) / Progress.Max
End Sub

Public Sub DIBDither_ProgressEnd()

    '-- Progress end
    Progress.Value = 0
    Progress.Cls
    Progress.Visible = False
    Call Canvas.Repaint
    Call gfPanView.Repaint
    '-- Save Undo
    Call pvSaveUndoDIB
End Sub

Private Sub RefreshFileInfo()
    '-- Compact path to Textfile panel width
    If (m_LastFilename <> "[未命名]") Then
        Dim strTemp As String
        strTemp = Mid(m_LastFilename, InStrRev(m_LastFilename, "\") + 1)
        stbThis.Panels(1).Text = "文件名: " & strTemp ' CompactPath(Me.hDC, m_LastFilename, Info.TextFileWidth)
      Else
        stbThis.Panels(1).Text = "文件名: [未命名]"
    End If
End Sub

Private Sub DoFileMenu(Index As Integer)
    '文件菜单的执行事件
    Dim fDlg     As New fDialogEx
    Dim sRet     As String
    Dim bSuccess As Boolean
    Select Case Index
        Case 0 '-- Open...
            '-- Show Open Dialog
            sRet = GetFileName(m_LastPath, "所有支持的文件格式|*.bmp;*.gif;*.jpg;*.png;*.tif|位图格式文件 (*.bmp)|*.bmp|GIF 格式文件 (*.gif)|*.gif|JPEG 格式文件 (*.jpg)|*.jpg|PNG 格式文件 (*.png)|*.png|TIFF 格式文件 (*.tif)|*.tif", 0, "打开...", True, fDlg)
            If (sRet <> vbNullString) Then
                '-- Get last path
                m_LastPath = sRet
                '-- Create DIB
                DoEvents
                Screen.MousePointer = vbHourglass
                Call pvSetDIBPicture(pvGetStdPicture(sRet, bSuccess))
                Screen.MousePointer = vbNormal
                
                If (bSuccess) Then
                    '-- Reset Undo/Redo and save first Undo
                    Call pvClearAllDIB
                    Call pvSaveUndoDIB
                    '-- Save info
                    m_LastFilename = sRet
                    Call RefreshFileInfo
                End If
            End If
        Case 1 '-- Save
'            If (m_LastFilename = "[未命名]" Or (FileFound(pvExtToBMP(m_LastFilename)) And pvExtToBMP(m_LastFilename) <> m_LastFilename)) Then
'                '-- Save as...
'                Call Unload(fDlg)
'                Set fDlg = Nothing
'                Call DoFileMenu(2)
'              Else
'                '-- Save as BMP
'                DoEvents
'                Call DIBSave.Save_BMP(pvExtToBMP(m_LastFilename), Canvas.DIB, DIBPal, DIBDither, DIBbpp)
'                '-- Saved flag
'                m_Saved = True
'                '-- Save info
'                m_LastFilename = pvExtToBMP(m_LastFilename)
'                Call RefreshFileInfo
'            End If
            Dim strFileName As String
            strFileName = m_Temp & "\R" & m_AppID & ".jpg"
            Call pvCorrectExt(strFileName)
            Call mGDIpEx.SaveDIB(gfrmMain.Canvas.DIB, strFileName, [ImageJPEG], 100)         '100%的图片质量
            If FileFound(strFileName) Then
                Set picFinal.Picture = LoadPicture(strFileName)
                RaiseEvent pOK(picFinal.Picture, picFinal.Width, picFinal.Height)
                m_Saved = True
                Err = 0: On Error Resume Next: Kill strFileName
            Else
                MsgBox "保存失败！", vbOKOnly + vbInformation, "zlPictureEditor"
            End If
        Case 2 '-- Save as...
            '-- Show Open Dialog
            sRet = GetFileName(m_LastFilename, "Bitmap (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif", 0, "另存为...", False, fDlg)
            If (sRet <> vbNullString) Then
                '-- Missing ext.?
                Call pvCorrectExt(sRet)
                '-- Save...
                DoEvents
                Select Case m_FileExt
                    Case "*.bmp" '-- BMP
                        Call DIBSave.Save_BMP(sRet, Canvas.DIB, DIBPal, DIBDither, DIBbpp)
                    Case "*.gif" '-- GIF
                        Call mGDIpEx.SaveDIB(gfrmMain.Canvas.DIB, sRet, [ImageGIF])
                    Case "*.jpg" '-- JPEG
                        Call mGDIpEx.SaveDIB(gfrmMain.Canvas.DIB, sRet, [ImageJPEG], fDlg.txtJPEGQuality)
                    Case "*.png" '-- PNG
                        Call mGDIpEx.SaveDIB(gfrmMain.Canvas.DIB, sRet, [ImagePNG])
                    Case "*.tif" '-- TIFF
                        Call mGDIpEx.SaveDIB(gfrmMain.Canvas.DIB, sRet, [ImageTIFF])
                End Select
                '-- Saved flag
                m_Saved = True
                '-- Save info
                If (m_FileExt = "*.bmp") Then
                    m_LastFilename = sRet
                End If
                Call RefreshFileInfo
            End If
        Case 3 '-- Print...
            If (Printers.Count) Then
                Call gfPrint.Show(vbModal, Me)
              Else
                Call MsgBox("对不起，没有安装打印机！", vbExclamation)
            End If
        Case 4 '-- Exit
            Call Unload(Me)
    End Select
    
    Call Unload(fDlg)
    Set fDlg = Nothing
End Sub

Private Sub mnuEditTop_Click()
End Sub

Private Sub DoEditMenu(Index As Integer)
    Dim sRet As VbMsgBoxResult
    Select Case Index
        Case 0 '-- Undo
            Call Undo
        Case 1 '-- Redo
            Call Redo
        Case 2 '-- Copy
            If (Canvas.DIB.hDIB <> 0) Then
                Call Canvas.DIB.CopyToClipboard
            End If
        Case 3 '-- Paste
            If (Clipboard.GetFormat(vbCFBitmap)) Then
                '-- Something there ?
                If (Canvas.DIB.hDIB <> 0 And Not m_Saved And m_UndoPos > 1) Then
                    '-- Ask for save
                    sRet = MsgBox("是否在粘贴前保存改变？", vbYesNoCancel Or vbInformation)
                    Select Case sRet
                        Case vbYes    '-- Save
                            Call DoFileMenu(1)
                        Case vbNo     '-- Ignore
                        Case vbCancel '-- Exit
                            Exit Sub
                    End Select
                End If
                '-- Initialize DIB
                Call pvSetDIBPicture(Clipboard.GetData(vbCFBitmap))
                '-- Reset Undo/Redo and save first Undo
                Call pvClearAllDIB
                Call pvSaveUndoDIB
                '-- [未命名] image
                m_LastFilename = "[未命名]"
            End If
        Case 4 '-- Resize
            Call gfResize.Show(vbModal, Me)
        Case 5 '-- Orientation
            Call gfOrientation.Show(vbModal, Me)
        Case 6 '-- Scroll mode
            Canvas.WorkMode = [cnvScrollMode]
            Call Canvas.Repaint
        Case 7 '-- Crop mode
            Canvas.WorkMode = [cnvCropMode]
    End Select
End Sub

Public Sub DoZoomMenu(Index As Integer)
    Select Case Index
        Case 0 '-- Zoom +
            Canvas.Zoom = Canvas.Zoom + IIf(Canvas.Zoom < 25, 1, 0)
            Canvas.FitMode = False
        Case 1 '-- Zoom -
            Canvas.Zoom = Canvas.Zoom - IIf(Canvas.Zoom > 1, 1, 0)
            Canvas.FitMode = False
        Case 2 '-- 1 : 1
            Canvas.Zoom = 1
            Canvas.FitMode = False
        Case 3 '-- Best fit
            Canvas.FitMode = True
    End Select
    Call Canvas.Resize
    stbThis.Panels(3).Text = Format(Canvas.Zoom, "0%")
End Sub

Private Sub DoColorMenu(Index As Integer)
    Dim sDIB As New cDIB
    Dim bfW As Long, bfH As Long
    Dim bfx As Long, bfy As Long
    
    Select Case Index
        Case 0  '-- Black and White (Stucki)
            DIBbpp = 1
            Call DIBPal.CreateBlackAndWhite
            Call DIBDither.DitherToBlackAndWhite(Canvas.DIB, 16)
        Case 1  '-- 16 greys
            DIBbpp = 4
            Call DIBPal.CreateGreys_016
            Call DIBDither.DitherToGreyPalette(Canvas.DIB, DIBPal, True)
        Case 2  '-- 256 greys
            DIBbpp = 8
            Call DIBPal.CreateGreys_256
            Call DIBDither.DitherToGreyPalette(Canvas.DIB, DIBPal, False)
        Case 3  '-- 2 colors
            DIBbpp = 1
            Call DIBPal.CreateBlackAndWhite
            Call DIBDither.DitherToBlackAndWhite(Canvas.DIB, 16)
        Case 4  '-- 16 colors
            DIBbpp = 4
            If (DIBPal.IsGreyScale) Then
                Call DIBPal.CreateGreys_016
                Call DIBDither.DitherToGreyPalette(Canvas.DIB, DIBPal, True)
              Else
                '-- Strecth to fit 150x150 (This will speed up all this)
                Call Canvas.DIB.GetBestFitInfo(150, 150, bfx, bfy, bfW, bfH)
                Call sDIB.Create(bfW, bfH)
                Call sDIB.LoadDIBBlt(Canvas.DIB)
                '-- Create optimal palette and dither.
                '   I don't know why these weight coeffs. work well...
                '   wChannel = f(Lchannel)
                '   wR = 1/(3-0.222) = 0.360
                '   wG = 1/(3-0.707) = 0.436
                '   wB = 1/(3-0.071) = 0.341
                Screen.MousePointer = vbHourglass
                Call DIBPal.CreateOptimal(sDIB, 16, 8, 0.36, 0.436, 0.341)
                Screen.MousePointer = vbNormal
                Call DIBDither.DitherToColorPalette(Canvas.DIB, DIBPal, True)
            End If
        Case 5  '-- 256 colors
            DIBbpp = 8
                If (DIBPal.IsGreyScale) Then
                Call DIBPal.CreateGreys_256
                Call DIBDither.DitherToGreyPalette(Canvas.DIB, DIBPal, False)
              Else
                '-- Strecth to fit 150x150 (This will speed up all this)
                Call Canvas.DIB.GetBestFitInfo(150, 150, bfx, bfy, bfW, bfH)
                Call sDIB.Create(bfW, bfH)
                Call sDIB.LoadDIBBlt(Canvas.DIB)
                '-- Create optimal palette and dither
                Screen.MousePointer = vbHourglass
                Call DIBPal.CreateOptimal(sDIB, 256, 8, 1, 1, 1)
                Screen.MousePointer = vbNormal
                Call DIBDither.DitherToColorPalette(Canvas.DIB, DIBPal, True)
            End If
        Case 6  '-- True color (24bpp)
            DIBbpp = 24
            Call DIBPal.Clear
            Call DIBDither.DitherToTrueColor(Canvas.DIB)
    End Select
    '-- Refresh
    Call Canvas.Repaint
    '-- Select current mode
    Call pvSetPalMode(DIBbpp)
    '-- Update info
    stbThis.Panels(2).Text = Canvas.DIB.Width & "×" & Canvas.DIB.Height & "×" & DIBbpp & "bpp"
End Sub

Private Sub DoAdjustMenu(Index As Integer)
    Select Case Index
    Case 0 '-- Brightness...
        Call gfFilter.Initialize(fltBrightness)
        Call gfFilter.Show(vbModal, Me)
    Case 1 '-- Contrast...
        Call gfFilter.Initialize(fltContrast)
        Call gfFilter.Show(vbModal, Me)
    Case 2 '-- Saturation...
        Call gfFilter.Initialize([fltSaturation])
        Call gfFilter.Show(vbModal, Me)
    Case 3 '-- Filter browser...
        Call gfFilter.Initialize(m_LastFilter)
        Call gfFilter.Show(vbModal, Me)
    End Select
    Call Canvas.Repaint
End Sub

Private Sub DoFilterColorMenu(Index As Integer)
    Select Case Index
    Case 0 '-- Greys
        Call DIBFilter.Greys(Canvas.DIB)
    Case 1 '-- Negative
        Call DIBFilter.Negative(Canvas.DIB)
    Case 2 '-- Sepia
        Call DIBFilter.Colorize(Canvas.DIB, 0.5, 0.25)
    Case 3 '-- Colorize...
        Call gfFilter.Initialize([fltColorize])
        Call gfFilter.Show(vbModal, Me)
    Case 4 '-- Replace HS...
        Call gfFilter.Initialize([fltReplaceHS])
        Call gfFilter.Show(vbModal, Me)
    Case 5 '-- Replace L...
        Call gfFilter.Initialize([fltReplaceL])
        Call gfFilter.Show(vbModal, Me)
    Case 6 '-- Shift...
        Call gfFilter.Initialize(fltShift)
        Call gfFilter.Show(vbModal, Me)
    End Select
    
    Call Canvas.Repaint
End Sub

Private Sub DoFilterDefMenu(Index As Integer)
    Select Case Index
    Case 0 '-- Blur
        Call DIBFilter.Blur(Canvas.DIB)
    Case 1 '-- Soften
        Call DIBFilter.Soften(Canvas.DIB)
    Case 2 '-- Sharpen
        Call DIBFilter.Sharpen(Canvas.DIB)
    Case 3 '-- Diffuse
        Call DIBFilter.Diffuse(Canvas.DIB)
    Case 4 '-- Pixelize
        Call DIBFilter.Pixelize(Canvas.DIB)
    Case 5 '-- Despeckle
        Call DIBFilter.Despeckle(Canvas.DIB)
    Case 6 '-- Despeckle more
        Call DIBFilter.DespeckleMore(Canvas.DIB)
    End Select
    
    Call Canvas.Repaint
End Sub

Private Sub DoFilterEdgesMenu(Index As Integer)
    Select Case Index
    Case 0 '-- Contour
        Call DIBFilter.Contour(Canvas.DIB)
    Case 1 '-- Emboss
        Call DIBFilter.Emboss(Canvas.DIB)
    Case 2 '-- Outline
        Call DIBFilter.Outline(Canvas.DIB)
    Case 3 '-- Relieve
        Call DIBFilter.Relieve(Canvas.DIB)
    End Select
    
    Call Canvas.Repaint
End Sub

Private Sub DoFilterSpecialMenu(Index As Integer)
    Select Case Index
    Case 0 '-- Noise
        Call DIBFilter.Noise(Canvas.DIB)
    Case 1 '-- Scanlines
        Call DIBFilter.Scanlines(Canvas.DIB)
    Case 2 '-- Dilate (Max.R.F. - 4N)
        Call DIBFilter.RankFilterMaximum(Canvas.DIB)
    Case 3 '-- Erode (Min.R.F. - 4N)
        Call DIBFilter.RankFilterMinimum(Canvas.DIB)
    Case 4 '-- Texturize...
        Call gfTexturize.Show(vbModal, Me)
    End Select
    
    Call Canvas.Repaint
End Sub

Private Sub DoViewMenu(Index As Integer)
    Select Case Index
    Case 1 '-- Panoramic view
        If gfPanView.Visible = False Then
            Call gfPanView.Show(IIf(mblnModeless, vbModeless, vbModal), Me)
          Else
            Call gfPanView.Hide
        End If
    Case 2 '-- Properties...
        Call gfProperties.Show(vbModal, Me)
    End Select
End Sub

'========================================================================================
' 画布 键盘事件、卷动、剪切
'========================================================================================

Private Sub Canvas_Resize()
    Call gfPanView.Repaint
End Sub

Private Sub Canvas_Scroll()
    Call gfPanView.Repaint
End Sub

Private Sub Canvas_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim scrHMax As Long, scrVMax As Long
  Dim scrHPos As Long, scrVPos As Long

  Dim bScroll As Boolean
  Dim lInc    As Long
    
    With Canvas
        Select Case KeyCode
            Case vbKeyAdd      '{NumPad +}
                Call DoZoomMenu(0)
            Case vbKeySubtract '{NumPad -}
                Call DoZoomMenu(1)
            Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
                Call .GetScrollInfo(scrHMax, scrVMax, scrHPos, scrVPos)
                bScroll = True
        End Select
                    
        If (bScroll) Then
            lInc = 10 * Canvas.Zoom
            Select Case KeyCode
                Case vbKeyUp    '{Cursor Up}
                    If (scrVPos > 0) Then
                        Call .SetScrollInfo(scrHPos, scrVPos - lInc)
                      Else
                        Call .SetScrollInfo(scrHPos, 0)
                    End If
                Case vbKeyDown  '{Cursor Down}
                    If (scrVPos < scrVMax) Then
                        Call .SetScrollInfo(scrHPos, scrVPos + lInc)
                      Else
                        Call .SetScrollInfo(scrHPos, scrVMax)
                    End If
                Case vbKeyLeft  '{Cursor Left}
                    If (scrHPos > 0) Then
                        Call .SetScrollInfo(scrHPos - lInc, scrVPos)
                      Else
                        Call .SetScrollInfo(0, scrVPos)
                    End If
                Case vbKeyRight '{Cursor Right}
                    If (scrHPos < scrHMax) Then
                        Call .SetScrollInfo(scrHPos + lInc, scrVPos)
                      Else
                        Call .SetScrollInfo(scrHMax, scrVPos)
                    End If
            End Select
            Call gfPanView.Repaint
        End If
        
        Call Canvas.Repaint
    End With
End Sub

Private Sub Canvas_Crop()
    '-- Change to True color mode
    DIBbpp = 24
    Call pvSetPalMode(DIBbpp)
    
    '-- Update Info and Progress
    With Canvas.DIB
        stbThis.Panels(2).Text = .Width & "×" & .Height & "×" & DIBbpp & "bpp"
        Progress.Max = .Height
    End With
    
    '-- Refresh Panoramic view
    Call gfPanView.Repaint
    
    '-- Save Undo
    Call pvSaveUndoDIB
End Sub

'========================================================================================
' DIB/调色板 初始化
'========================================================================================

Private Function pvGetStdPicture(ByVal sFilename As String, bSuccess As Boolean) As StdPicture
    On Error Resume Next
    If (pvGetExt(sFilename) = "png" Or pvGetExt(sFilename) = "tif") Then
        '-- Use GDI+ loading
        Set pvGetStdPicture = mGDIpEx.LoadPictureEx(sFilename)
      Else
        '-- Use VB LoadPicture
        Set pvGetStdPicture = LoadPicture(sFilename)
    End If
    
    '-- Is there an image ?
    bSuccess = Not (pvGetStdPicture Is Nothing)
    
    If (bSuccess = False) Then
        '-- Nothing loaded
        Call MsgBox("调入图片时发生意外错误！", vbExclamation)
    End If
    
    On Error GoTo 0
End Function
    
Private Sub pvSetDIBPicture(Image As StdPicture)
  Static lstW As Long
  Static lstH As Long

    If (Not Picture Is Nothing) Then

        '-- Save last DIB dimensions
        lstW = Canvas.DIB.Width
        lstH = Canvas.DIB.Height
        
        '-- Clear palette
        Call DIBPal.Clear
        
        '-- Create 32bpp DIB section from std. picture.
        '   Case source <=8bpp, palette saved in DIBPal, palette indexes in DIBDither.
        '   Return value: source color depth / 0 = Err.
        DIBbpp = Canvas.DIB.CreateFromStdPicture(Image, DIBPal, DIBDither)
        
        '-- Select current depth mode
        Call pvSetPalMode(DIBbpp)
        
        '-- Remove Crop rectangle and resize canvas
        Call Canvas.RemoveCropRectangle
        With Canvas.DIB
            If (lstW <> .Width Or lstH <> .Height) Then
                Call Canvas.Resize
              Else
                Call Canvas.Repaint
            End If
        End With
        
        '-- Refresh panoramic view
        Call gfPanView.Repaint
        
        '-- Set progress bar max value
        Progress.Max = Canvas.DIB.Height
        
        '-- Show image info: Size + bpp
        stbThis.Panels(2).Text = Canvas.DIB.Width & "×" & Canvas.DIB.Height & "×" & DIBbpp & "bpp"
        stbThis.Panels(3).Text = Format(Canvas.Zoom, "0%")
    End If
End Sub

Private Sub pvSetPalMode(ByVal bpp As Long)
  Dim lIdxNew As Long
  Dim lIdxOld As Long
    
    Select Case bpp
        Case 1  '-- 2 colors / Black and White
            lIdxNew = IIf(DIBPal.IsGreyScale, 0, 4)
        Case 4  '-- 16 colors / 16 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 1, 5)
        Case 8  '-- 256 colors / 256 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 2, 6)
        Case 24 '-- True color
            lIdxNew = 8
        Case Else
            Exit Sub
    End Select
'
'    For lIdxOld = 0 To 8
'        mnuColors(lIdxOld).Checked = False
'    Next lIdxOld
'    mnuColors(lIdxNew).Checked = True
    
End Sub

'========================================================================================
' 重做/撤销 控制
'========================================================================================

Public Sub Undo()
    Dim sPath As String
    If (m_UndoPos > 1) Then
        '-- Get path
        sPath = m_Temp & "\b" & m_AppID & Format(m_UndoPos - 2, "000") & ".dat"
        '-- Load Undo DIB
        Call pvSetDIBPicture(LoadPicture(sPath))
        '-- Refresh Panoramic view
        Call gfPanView.Repaint
    
        If (m_UndoPos > 0) Then
            m_UndoPos = m_UndoPos - 1
        End If
    End If
End Sub

Public Sub Redo()
    Dim sPath As String
    If (m_UndoPos < m_UndoMax) Then
        '-- Get path
        sPath = m_Temp & "\b" & m_AppID & Format(m_UndoPos, "000") & ".dat"
        '-- Load Redo DIB
        Call pvSetDIBPicture(LoadPicture(sPath))
        '-- Refresh Panoramic view
        Call gfPanView.Repaint
    
        m_UndoPos = m_UndoPos + 1
        If (m_UndoPos > m_UndoMax) Then
            m_UndoMax = m_UndoPos
        End If
    End If
End Sub

Private Sub pvClearAllDIB()
    
    '-- Delete all temp. files
    On Error Resume Next
       Kill m_Temp & "\b" & m_AppID & "*.dat"
    On Error GoTo 0
    
    '-- Reset 'counters'
    m_UndoPos = 0
    m_UndoMax = 0
End Sub

Private Sub pvSaveUndoDIB()
    Dim lIdx  As Long
    Dim sPath As String
    
    '-- Get path
    sPath = m_Temp & "\b" & m_AppID & Format(m_UndoPos, "000") & ".dat"
    '-- Save DIB
    With gfrmMain
        Call .DIBSave.Save_BMP(sPath, .Canvas.DIB, .DIBPal, .DIBDither, .DIBbpp)
    End With
    '-- Saved flag
    m_Saved = False
    If (m_UndoMax - m_UndoPos > 0) Then
        On Error Resume Next
        For lIdx = m_UndoPos + 1 To m_UndoMax
            Kill m_Temp & "\b" & m_AppID & Format(lIdx, "000") & ".dat"
        Next lIdx
        On Error GoTo 0
    End If

    If (m_UndoPos < m_UNDO_LEVELS) Then
        m_UndoPos = m_UndoPos + 1
        m_UndoMax = m_UndoPos
      Else
        Call pvRotateUndoFiles
    End If
End Sub

Private Sub pvRotateUndoFiles()

  Dim bOldName As String
  Dim bNewName As String
  Dim lIdx     As Long

    On Error Resume Next
    '-- Kill first
    Kill m_Temp & "\b" & m_AppID & "000.dat"
    '-- 'Rotate' the others (Move up 1)
    For lIdx = 1 To m_UNDO_LEVELS
        bOldName = m_Temp & "\b" & m_AppID & Format(lIdx - 0, "000") & ".dat"
        bNewName = m_Temp & "\b" & m_AppID & Format(lIdx - 1, "000") & ".dat"
        Name bOldName As bNewName
    Next lIdx
    On Error GoTo 0
End Sub

Private Function pvExtToBMP(ByVal sFilename As String) As String
    pvExtToBMP = Left$(sFilename, Len(sFilename) - 3) & "bmp"
End Function

Private Function pvGetExt(ByVal sFilename As String) As String
    pvGetExt = Right$(sFilename, 3)
End Function

Private Function pvCorrectExt(sFilename As String)
    If (Right$(sFilename, 4) <> Right$(m_FileExt, 4)) Then
        sFilename = sFilename & Right$(m_FileExt, 4)
    End If
End Function

'========================================================================================
' 全局属性 (设置)
'========================================================================================

Public Property Let LastFilterID(ByVal FilterID As fltIDCts)
    m_LastFilter = FilterID
End Property

Public Property Get LastFilename() As String
    LastFilename = m_LastFilename
End Property

Public Property Let LastFilename(ByVal sLastFilename As String)
    m_LastFilename = sLastFilename
End Property

Public Property Get LastPath() As String
    LastPath = m_LastPath
End Property

Public Property Let LastPath(ByVal sLastPath As String)
    m_LastPath = sLastPath
End Property

Public Property Get FileExt() As String
    FileExt = m_FileExt
End Property

Public Property Let FileExt(ByVal sFileExt As String)
    m_FileExt = sFileExt
End Property

Public Property Get DialogPreview() As Boolean
    DialogPreview = m_DialogPreview
End Property

Public Property Let DialogPreview(ByVal bShow As Boolean)
    m_DialogPreview = bShow
End Property

Public Property Get DialogFitMode() As Boolean
    DialogFitMode = m_DialogFitMode
End Property

Public Property Let DialogFitMode(ByVal bEnable As Boolean)
    m_DialogFitMode = bEnable
End Property

Public Property Get DialogJPEGquality() As Integer
    DialogJPEGquality = m_DialogJPEGquality
End Property

Public Property Let DialogJPEGquality(ByVal iValue As Integer)
    m_DialogJPEGquality = iValue
End Property

Private Sub BeginShowProgress()
    '进度条定位
    On Error Resume Next
    With Progress
        .Left = stbThis.Panels(4).Left + Screen.TwipsPerPixelX * 2
        .Top = stbThis.Top + (stbThis.Height - .Height) / 2 + Screen.TwipsPerPixelY * 2
        .Width = stbThis.Panels(4).Width - Screen.TwipsPerPixelX * 4
        .Height = stbThis.Height - Screen.TwipsPerPixelY * 4
        .Visible = True: Me.Refresh
    End With
End Sub

