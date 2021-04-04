VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmMarkMapMan 
   Caption         =   "病历标记图管理"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   Icon            =   "frmFigureMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picVBar_S 
      BackColor       =   &H00808080&
      Height          =   4080
      Left            =   3495
      MouseIcon       =   "frmFigureMan.frx":058A
      MousePointer    =   99  'Custom
      ScaleHeight     =   4080
      ScaleWidth      =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   990
      Width           =   40
   End
   Begin MSComctlLib.ImageList imlTool 
      Left            =   855
      Top             =   225
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFigureMan.frx":06DC
            Key             =   ""
            Object.Tag             =   "301"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFigureMan.frx":0C76
            Key             =   ""
            Object.Tag             =   "302"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFigureMan.frx":1210
            Key             =   ""
            Object.Tag             =   "303"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFigureMan.frx":17AA
            Key             =   ""
            Object.Tag             =   "304"
         EndProperty
      EndProperty
   End
   Begin zlRichEPR.ucCanvas Canvas 
      Height          =   1140
      Left            =   4005
      TabIndex        =   2
      Top             =   1260
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   2011
   End
   Begin MSComctlLib.ImageList imglvw 
      Left            =   2850
      Top             =   4695
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFigureMan.frx":1D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFigureMan.frx":85A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   4230
      Left            =   15
      TabIndex        =   0
      Top             =   930
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   7461
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imglvw"
      SmallIcons      =   "imglvw"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   5850
      Top             =   585
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   $"frmFigureMan.frx":EE08
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5565
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3043
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1235
            MinWidth        =   1235
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   210
      Top             =   210
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmMarkMapMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ID_FitSize = 301                      '适合尺寸
Private Const ID_ActualSize = 302                   '实际尺寸
Private Const ID_ZoomIn = 303                       '放大
Private Const ID_ZoomOut = 304                      '缩小

'窗体级变量
Private mlngScaleLeft As Long, mlngScaleTop As Long, mlngScaleRight As Long, mlngScaleBottom As Long  '客户区域的大小
Private mstrPrivs As String

'临时变量
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem

'################################################################################################################
'-- 位图控制
Private WithEvents DIBFilter As cDIBFilter      ' DIB 滤镜对象(24 bpp)
Attribute DIBFilter.VB_VarHelpID = -1
Private WithEvents DIBDither As cDIBDither      ' DIB 抖动对象(1, 4, 8 bpp)
Attribute DIBDither.VB_VarHelpID = -1
Private DIBPal               As New cDIBPal     ' DIB 调色板对象 (1, 4, 8 bpp)
Private DIBSave              As New cDIBSave    ' Save 对象 (BMP)  (1, 4, 8, 24 bpp)
Private DIBbpp               As Byte            ' 当前颜色深度
Private WithEvents cPicEditor As cPictureEditor     ' 图片编辑对象
Attribute cPicEditor.VB_VarHelpID = -1
Private m_LastFilename As String                    ' 最后打开的图片位置
Private m_Temp As String                            ' 临时文件路径
Private m_AppID As Long
'-- GDI+
Private m_GDIpToken         As Long         ' 用于关闭 GDI+
'扫描函数声明
Private Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hWnd As Long, ByVal wPixTypes As Integer) As Integer
Private Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hWnd As Long) As Integer

Private Sub Canvas_Crop()
    DIBbpp = 24
    Call pvSetPalMode(DIBbpp)
    With Canvas.DIB
        stbThis.Panels(3).Text = "图片大小：" & Canvas.DIB.Width & "×" & Canvas.DIB.Height & "×" & DIBbpp & "位(Bpp)"
    End With
End Sub

Private Sub Canvas_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = vbRightButton Then
        Dim Popup As CommandBar
        Dim Control As CommandBarControl
        
        Set Popup = Me.cbsThis.Add("Popup", xtpBarPopup)
        With Popup.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "图片(&I)…"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, ID_FitSize, "适合尺寸(&F)"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, ID_ActualSize, "实际尺寸(&S)")
            Set cbrControl = .Add(xtpControlButton, ID_ZoomIn, "放大(&Z)")
            Set cbrControl = .Add(xtpControlButton, ID_ZoomOut, "缩小(&O)")
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub DoZoomMenu(Index As Integer)
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
    stbThis.Panels(4).Text = Format(Canvas.Zoom, "0%")
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strItemKey As String
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview:  Call zlRptPrint(0)
    Case conMenu_File_Print:    Call zlRptPrint(1)
    Case conMenu_File_Excel:    Call zlRptPrint(3)
    Case conMenu_File_Exit:     Unload Me
    Case ID_FitSize
        DoZoomMenu 3
    Case ID_ActualSize
        DoZoomMenu 2
    Case ID_ZoomIn
        DoZoomMenu 0
    Case ID_ZoomOut
        DoZoomMenu 1
    Case conMenu_Edit_NewItem
        strItemKey = frmMarkMapEdit.ShowMe(Me, True)
        If strItemKey <> "" Then Call zlRefLists(strItemKey)
    Case conMenu_Edit_Modify
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        
        Dim strFileName1 As String, strTmp1 As String, oDIB As New cDIB, bSuccess1 As Boolean
        
        strFileName1 = m_Temp & "\R" & m_AppID & ".jpg"
        Call mGdIpEx.SaveDIB(Me.Canvas.DIB, strFileName1, [ImageJPEG], 90)         '100%的图片质量，在修改时不损失质量
        Call oDIB.CreateFromStdPicture(pvGetStdPicture(strFileName1, bSuccess1), DIBPal, DIBDither)
        
        strItemKey = frmMarkMapEdit.ShowMe(Me, False, Mid(Me.lvwList.SelectedItem.Key, 2), oDIB)
        If strItemKey <> "" Then Call zlRefLists(strItemKey): Me.Canvas.Resize
    Case conMenu_Edit_Delete
        With Me.lvwList
            If .SelectedItem Is Nothing Then Exit Sub
            If MsgBox("真的删除该标记图吗？" & vbCrLf & "——" & .SelectedItem.Text, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "zl_病历标记图形_delete('" & Mid(.SelectedItem.Key, 2) & "')"
            Err = 0: On Error GoTo errHand
            Call SQLTest(App.ProductName, Me.Caption, gstrSQL): gcnOracle.Execute gstrSQL, , adCmdStoredProc: Call SQLTest
            Call .ListItems.Remove(.SelectedItem.Key)
            If Not .SelectedItem Is Nothing Then
                Call lvwList_ItemClick(.SelectedItem)
            Else
                stbThis.Panels(3).Text = ""
                Set Canvas.DIB = New cDIB
                Canvas.Resize
            End If
            Me.stbThis.Panels(2).Text = "剩余" & .ListItems.Count & "标记图"
            If .Visible And .Enabled Then .SetFocus
        End With
        Exit Sub
errHand:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
        Exit Sub
        
    Case conMenu_Edit_MarkMap
        Dim strFileName As String, bSuccess As Boolean, strTmp As String
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        dlgThis.InitDir = m_LastFilename
        dlgThis.CancelError = True
        On Error GoTo LL
        dlgThis.ShowOpen
        strFileName = dlgThis.Filename
        If gobjFSO.FileExists(strFileName) Then
            If MsgBox("注意：选择新图片将覆盖当前图片。是否继续？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        If Trim(strFileName) <> "" Then
            '-- Create DIB
            DoEvents
            Call pvSetDIBPicture(pvGetStdPicture(strFileName, bSuccess))
            
            If (bSuccess) Then
                m_LastFilename = strFileName
                stbThis.Panels(3).Text = "图片大小：" & Canvas.DIB.Width & "×" & Canvas.DIB.Height & "×" & DIBbpp & "位(Bpp)"
                stbThis.Panels(4).Text = Format(Canvas.Zoom, "0%")
            End If
        End If
        strFileName = m_Temp & "\R" & m_AppID & ".jpg"
        Call mGdIpEx.SaveDIB(Me.Canvas.DIB, strFileName, [ImageJPEG], 90)         '90%的图片质量，部分压缩
        
        Dim arySql() As String, lngSql As Long
        If gobjFSO.FileExists(strFileName) Then
            If zlBlobSql(0, Mid(Me.lvwList.SelectedItem.Key, 2), strFileName, arySql()) = False Then
                MsgBox "标记图形保存失败", vbExclamation, gstrSysName
                Exit Sub
            End If
            gobjFSO.DeleteFile strFileName  '删除临时文件
        End If
        
        '执行保存
        Err = 0: On Error GoTo ErrMap
        gcnOracle.BeginTrans
        For lngSql = LBound(arySql) To UBound(arySql)
            Call SQLTest(App.ProductName, Me.Caption, arySql(lngSql))
            gcnOracle.Execute arySql(lngSql), , adCmdStoredProc
            Call SQLTest
        Next
        gcnOracle.CommitTrans
        Exit Sub
ErrMap:
        gcnOracle.RollbackTrans
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
        Exit Sub
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh
        If Me.lvwList.SelectedItem Is Nothing Then
            strItemKey = ""
        Else
            strItemKey = Mid(Me.lvwList.SelectedItem.Key, 2)
        End If
        Call zlRefLists(strItemKey)
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    End Select
LL:
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    Call Me.cbsThis.GetClientRect(mlngScaleLeft, mlngScaleTop, mlngScaleRight, mlngScaleBottom)
    Call Form_Resize
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup
            Control.Visible = Not (InStr(1, mstrPrivs, "增删改") = 0 And InStr(1, mstrPrivs, "图片更改") = 0)
        End Select
    Else
        Err = 0: On Error Resume Next
        Select Case Control.ID
        Case ID_FitSize
            Control.Enabled = (Canvas.DIB.hDIB <> 0)
            Control.Checked = (Canvas.FitMode = True)
        Case ID_ActualSize
            Control.Enabled = (Canvas.DIB.hDIB <> 0)
        Case ID_ZoomIn
            Control.Enabled = (Canvas.DIB.hDIB <> 0)
        Case ID_ZoomOut
            Control.Enabled = (Canvas.DIB.hDIB <> 0)
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            Control.Enabled = Not (Me.lvwList.ListItems.Count = 0)
        Case conMenu_Edit_NewItem
            Control.Visible = Not (InStr(1, mstrPrivs, "增删改") = 0)
        Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_MarkMap
            Control.Visible = Not (InStr(1, mstrPrivs, "增删改") = 0)
            Control.Enabled = Not (Me.lvwList.SelectedItem Is Nothing)
        Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
        Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
        Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
        End Select
    End If
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    '图片操作相关
    m_LastFilename = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "LastFilename", App.Path)
    
    Dim GpInput As GdiplusStartupInput
    '-- 调入 GDI+ Dll
    GpInput.GdiplusVersion = 1
    If (mGdIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
        Call MsgBox("调入 GDI+ 出错，无法进行图片插入！请检查 GDI+ DLL 是否存在或者损坏！", vbInformation + vbOKOnly)
        Call Unload(Me)
        Exit Sub
    End If
    
    m_Temp = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    m_AppID = Me.hWnd
    Set DIBFilter = New cDIBFilter
    Set DIBDither = New cDIBDither
    Set cPicEditor = New cPictureEditor
    
    Canvas.FitMode = True
    
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "图片(&I)…"): cbrControl.BeginGroup = True
        
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        Set cbrControl = .Add(xtpControlButton, ID_FitSize, "适合尺寸(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, ID_ActualSize, "实际尺寸(&S)")
        Set cbrControl = .Add(xtpControlButton, ID_ZoomIn, "放大(&Z)")
        Set cbrControl = .Add(xtpControlButton, ID_ZoomOut, "缩小(&O)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add FCONTROL, Asc("I"), conMenu_Edit_MarkMap
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, vbKeySubtract, ID_ZoomOut
        .Add 0, vbKeyAdd, ID_ZoomIn
        .Add FCONTROL, vbKeyF, ID_FitSize
        .Add FCONTROL, vbKeyS, ID_ActualSize
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "图片")
        cbrControl.BeginGroup = True
                
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '数据元素形态设置
    With Me.lvwList.ColumnHeaders
        .Clear
        .Add , "_名称", "名称", 2000
        .Add , "_编码", "编码", 650
        .Add , "_简码", "简码", 1000
    End With
    With Me.lvwList
        .ColumnHeaders("_编码").Position = 1
        .SortKey = .ColumnHeaders("_编码").Index - 1
        .SortOrder = lvwAscending
    End With
    
    '-----------------------------------------------------
    '加入附加图标
    Me.cbsThis.AddImageList Me.imlTool
    '界面恢复
    Call Me.cbsThis.GetClientRect(mlngScaleLeft, mlngScaleTop, mlngScaleRight, mlngScaleBottom)
    Call RestoreWinState(Me, App.ProductName)
    Me.picVBar_S.BackColor = Me.BackColor
    
    '-----------------------------------------------------
    '数据装入
    Call zlRefLists
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.picVBar_S
        .Top = mlngScaleTop: .Height = mlngScaleBottom - mlngScaleTop
        If .Left < 2000 + mlngScaleLeft Then .Left = 2000 + mlngScaleLeft
        If .Left > mlngScaleRight - mlngScaleLeft - 2000 Then .Left = mlngScaleRight - mlngScaleLeft - 2000
    End With
    With Me.lvwList
        .Left = mlngScaleLeft: .Width = Me.picVBar_S.Left - .Left
        .Top = mlngScaleTop: .Height = mlngScaleBottom - .Top
    End With
    With Me.Canvas
        .Top = Me.lvwList.Top: .Height = mlngScaleBottom - .Top
        .Left = Me.picVBar_S.Left + Me.picVBar_S.Width: .Width = mlngScaleRight - .Left
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "LastFilename", m_LastFilename

    Me.Canvas.DIB.Destroy
    LockWindowUpdate 0
    UpdateWindow Me.hWnd
    ' Unload the GDI+ Dll
    Call mGdIpEx.GdiplusShutdown(m_GDIpToken)

    '-- Free objects
    Set DIBFilter = Nothing
    Set DIBDither = Nothing
    Set DIBPal = Nothing
    Set DIBSave = Nothing
    Set cPicEditor = Nothing
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgMark_DblClick()
    Set cbrControl = Me.cbsThis.FindControl(xtpControlButton, conMenu_Edit_MarkMap)
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwList.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwList.SortOrder = IIf(Me.lvwList.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwList.SortKey = ColumnHeader.Index - 1
        Me.lvwList.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwList_DblClick()
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strTemp As String, bSuccess As Boolean
    
    If lvwList.Tag = Item Then Exit Sub
    Screen.MousePointer = vbHourglass
    stbThis.Panels(3).Text = ""
    Set Canvas.DIB = New cDIB
    strTemp = zlBlobRead(0, Mid(Item.Key, 2))
    If Len(strTemp) > 0 Then
        Call pvSetDIBPicture(pvGetStdPicture(strTemp, bSuccess))
        If bSuccess Then
            stbThis.Panels(3).Text = "图片大小：" & Canvas.DIB.Width & "×" & Canvas.DIB.Height & "×" & DIBbpp & "位(Bpp)"
        Else
            If Err <> 0 Then MsgBox "图片可能被损坏！", vbExclamation, gstrSysName
        End If
        Kill strTemp
    End If
    Canvas.Resize
    lvwList.Tag = Item
    Screen.MousePointer = vbNormal
End Sub

Private Sub lvwList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call lvwList_DblClick
End Sub

Private Sub lvwList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '弹出菜单定义
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button <> vbRightButton Or Me.cbsThis.ActiveMenuBar.Controls(2).Visible <> True Then Exit Sub
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        If cbrControl.Visible Then
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
            cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        End If
    Next
    Set cbrControl = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_File_Print, "打印")
    cbrControl.BeginGroup = True
    
    cbrPopupBar.ShowPopup
End Sub

Private Sub picVBar_S_DblClick()
    If lvwList.ListItems.Count > 0 Then
        picVBar_S.Left = lvwList.ListItems(1).Width + Screen.TwipsPerPixelX * 4
        Call Form_Resize
    End If
End Sub

Private Sub picVBar_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Me.picVBar_S.Left = Me.picVBar_S.Left + X
End Sub

Private Sub picVBar_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call Form_Resize
End Sub

Public Sub zlRefLists(Optional strKeyCode As String)
    '---------------------------------------------
    '填写列表
    '---------------------------------------------
    Err = 0: On Error GoTo errHand
    
    gstrSQL = "Select 编码,名称,简码 From 病历标记图形 Order By 编码"
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.lvwList.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwList.ListItems.Add(, "_" & !编码, !名称)
            objItem.SubItems(Me.lvwList.ColumnHeaders("_编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwList.ColumnHeaders("_简码").Index - 1) = IIf(IsNull(!简码), "", !简码)
            objItem.Icon = 1: objItem.SmallIcon = objItem.Icon
            If !编码 = strKeyCode Then objItem.Selected = True
            .MoveNext
        Loop
    End With
    If Me.lvwList.ListItems.Count > 0 Then
        If Me.lvwList.SelectedItem Is Nothing Then Me.lvwList.ListItems(1).Selected = True
        Me.lvwList.SelectedItem.EnsureVisible
        Call lvwList_ItemClick(Me.lvwList.SelectedItem)
        Me.stbThis.Panels(2).Text = "共有" & Me.lvwList.ListItems.Count & "标记图"
    Else
        stbThis.Panels(3).Text = ""
        Set Canvas.DIB = New cDIB
        Canvas.Resize
        Me.stbThis.Panels(2).Text = ""
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:记录表打印
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    If Me.lvwList.ListItems.Count = 0 Then Exit Sub
    
    Err = 0: On Error Resume Next
    Set objPrint.Body.objData = Me.lvwList
    objPrint.Title.Text = "病历标记图清单"
    objPrint.BelowAppItems.Add "打印时间:" & Now
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrViewLvw objPrint, bytMode
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

'################################################################################################################
'## 功能：  画布相关函数
'################################################################################################################
Private Function pvGetStdPicture(ByVal sFileName As String, bSuccess As Boolean) As StdPicture
    On Error Resume Next
    If (pvGetExt(sFileName) = "png" Or pvGetExt(sFileName) = "tif") Then
        '-- Use GDI+ loading
        Set pvGetStdPicture = mGdIpEx.LoadPictureEx(sFileName)
      Else
        '-- Use VB LoadPicture
        Set pvGetStdPicture = LoadPicture(sFileName)
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
        
        '-- Show image info: Size + bpp
'        stbThis.Panels(3).Text = "图片大小：" & Canvas.DIB.Width & "×" & Canvas.DIB.Height & "×" & DIBbpp & "位(Bpp)"
        stbThis.Panels(4).Text = Format(Canvas.Zoom, "0%")
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
End Sub

Private Function pvGetExt(ByVal sFileName As String) As String
    pvGetExt = Mid(sFileName, Len(sFileName) - 2)
End Function

