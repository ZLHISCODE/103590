VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPartogram 
   Caption         =   "产程图查阅"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10725
   Icon            =   "frmPartogram.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10725
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3615
      ScaleWidth      =   5655
      TabIndex        =   1
      Top             =   480
      Width           =   5655
      Begin VB.PictureBox picTmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   720
         ScaleHeight     =   300
         ScaleWidth      =   1005
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1000
         Begin VB.ComboBox cboBaby 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   0
            Width           =   1000
         End
      End
      Begin MSComCtl2.FlatScrollBar vsb 
         Height          =   1155
         Left            =   5040
         TabIndex        =   3
         Top             =   1320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2037
         _Version        =   393216
         Appearance      =   0
         Max             =   100
         Orientation     =   1572864
      End
      Begin MSComCtl2.FlatScrollBar hsb 
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   2640
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         Max             =   100
         Orientation     =   1572865
      End
      Begin VB.PictureBox picMain 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   0
         ScaleHeight     =   2655
         ScaleWidth      =   5055
         TabIndex        =   2
         Top             =   0
         Width           =   5055
         Begin VB.PictureBox PicDraw 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2000
            Left            =   600
            ScaleHeight     =   1995
            ScaleWidth      =   1995
            TabIndex        =   7
            Top             =   480
            Width           =   2000
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6435
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPartogram.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16007
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
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPartogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************
'病人基本信息
'***************************************************************
Private Type type_Patient
    lng文件ID As Long
    lng病人ID As Long
    lng主页ID As Long
    lng科室ID As Long
    lng编辑 As Long
    lng份数 As Long
End Type
Private T_Info As type_Patient

Private mblnChildForm As Boolean
Private mblnShowState As Boolean
Private mblnShowOk As Boolean
Private mstrPrivs As String
Private msinVStep As Single      '滚动条的步长
Private msinHStep As Single      '滚动条的步长
Private mblnInit As Boolean
Private mintPage As Integer
Private mbytFontSize  As Byte

'获取当前选中的婴儿号
Public Property Get FileNumIndex() As Long
    FileNumIndex = T_Info.lng份数
End Property

Public Property Let FileNumIndex(lngIndex As Long)
    T_Info.lng份数 = lngIndex
End Property

Public Property Get FileID() As Long
    FileID = T_Info.lng文件ID
End Property

Public Property Get PatiID() As Long
    PatiID = T_Info.lng病人ID
End Property

Public Property Get PageID() As Long
    PageID = T_Info.lng主页ID
End Property

Public Property Get PartogramParam() As String
    PartogramParam = T_Info.lng文件ID & ";" & T_Info.lng病人ID & ";" & T_Info.lng主页ID & ";" & T_Info.lng科室ID & ";" & T_Info.lng编辑 & ";" & T_Info.lng份数
End Property

Public Property Get ScrollBarY() As FlatScrollBar
    Set ScrollBarY = vsb
End Property

Public Property Get ScrollBarX() As FlatScrollBar
    Set ScrollBarX = hsb
End Property


Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-19 15:16
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = mbytFontSize
    
    Me.FontSize = bytFontSize
    Me.FontName = "宋体"
    
    Set CtlFont = cbsThis.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsThis.Options.Font = CtlFont
    stbThis.Font.Size = bytFontSize
    cboBaby.FontSize = bytFontSize
    picTmp.Height = cboBaby.Height
    cbsThis.RecalcLayout
End Sub

Public Function ShowEdit(ByVal frmMain As Object, strParam As String, Optional ByVal bytMode As Byte = 1, Optional strPrivs As String, Optional ByVal bytSize As Byte = 0) As Boolean
'******************************************************************************************************************
'功能：完成产程图展示;独立展示
'参数：frmMain：父窗体;strParam:病人信息格式串(文件ID;病人ID;主页ID;科室ID;是否允许编辑);bytmode(是否以模态窗体显示);strPrivs;权限
'返回：数据成功返回True,否则返回false
'******************************************************************************************************************
    Dim blnShowSatate As Boolean
    Dim varParam() As String
    
    mstrPrivs = strPrivs
    mblnChildForm = False
    mblnShowOk = False
    blnShowSatate = mblnShowState
    mblnInit = False
    
    If UBound(Split(strParam & ";", ";")) < 5 Then
        MsgBox "请检查传入的参数格式!" & vbCrLf & _
            "格式为：[文件ID;病人ID;主页ID;科室ID;是否允许编辑]", vbInformation, gstrSysName
        Exit Function
    End If
    
    varParam = Split(strParam, ";")
    If blnShowSatate Then
        If Val(varParam(0)) = T_Info.lng文件ID Then
            Call ShowWindow(Me.Hwnd, SW_RESTORE)
            Call BringWindowToTop(Me.Hwnd)
            Exit Function
        End If
    End If
    
    T_Info.lng文件ID = Val(varParam(0))
    T_Info.lng病人ID = Val(varParam(1))
    T_Info.lng主页ID = Val(varParam(2))
    T_Info.lng科室ID = Val(varParam(3))
    
    If InStr(1, ";" & strPrivs & ";", ";产程图作图;") = 0 Then
        T_Info.lng编辑 = 0
    Else
        T_Info.lng编辑 = 1
    End If
    If UBound(varParam) > 3 Then T_Info.lng编辑 = Val(varParam(4))
    T_Info.lng份数 = 1
    If UBound(varParam) > 4 Then T_Info.lng份数 = Val(varParam(5))
    If T_Info.lng份数 < 0 Then T_Info.lng份数 = 1
    '进行数据展示
    If Not InitBodyPartogram Then
        Unload Me
        Exit Function
    End If
    stbThis.Visible = Not mblnChildForm
    cbsThis.ActiveMenuBar.Visible = Not mblnChildForm
    mblnShowState = True
    If blnShowSatate = False Then
        Call SetFontSize(bytSize)
        Me.WindowState = 2
        Set mfrmPartogram = Me
        Hook mfrmPartogram.Hwnd
        If bytMode = 1 Then
            Me.Show 1, frmMain
        Else
            Me.Show , frmMain
        End If
        strParam = T_Info.lng文件ID & ";" & T_Info.lng病人ID & ";" & T_Info.lng主页ID & ";" & T_Info.lng科室ID & ";" & T_Info.lng编辑 & ";" & T_Info.lng份数
        ShowEdit = mblnShowOk
    End If
End Function

Public Function zlRefresh(ByVal frmMain As Object, strParam As String, Optional strPrivs As String, Optional blnChildForm As Boolean = True) As Boolean
'******************************************************************************************************************
'功能：完成产程图数据刷新;控件加载
'参数：frmMain：父窗体;strParam:病人信息格式串(文件ID;病人ID;主页ID;科室ID;是否允许编辑);strPrivs;权限
'返回：数据成功返回True,否则返回false
'******************************************************************************************************************
    Dim blnShowSatate As Boolean
    Dim varParam() As String
    
    mstrPrivs = strPrivs
    mblnChildForm = blnChildForm
    mblnShowState = False
    mblnShowOk = False
    mblnInit = False
    
    If UBound(Split(strParam & ";", ";")) < 5 Then
        MsgBox "请检查传入的参数格式!" & vbCrLf & _
            "格式为：[文件ID;病人ID;主页ID;科室ID;是否允许编辑]", vbInformation, gstrSysName
        Exit Function
    End If
    
    '去掉窗体标题栏
    If mblnChildForm = True Then Call FormSetCaption(Me, False, False)

    stbThis.Visible = Not mblnChildForm
    cbsThis.ActiveMenuBar.Visible = Not mblnChildForm
    
    varParam = Split(strParam, ";")
    
    T_Info.lng文件ID = Val(varParam(0))
    T_Info.lng病人ID = Val(varParam(1))
    T_Info.lng主页ID = Val(varParam(2))
    T_Info.lng科室ID = Val(varParam(3))
    
    If InStr(1, ";" & strPrivs & ";", ";产程图作图;") = 0 Then
        T_Info.lng编辑 = 0
    Else
        T_Info.lng编辑 = 1
    End If
    If UBound(varParam) > 3 Then T_Info.lng编辑 = Val(varParam(4))
    T_Info.lng份数 = 1
    If UBound(varParam) > 4 Then T_Info.lng份数 = Val(varParam(5))
    
    '进行数据展示
    If Not InitBodyPartogram Then
        Exit Function
    End If
    zlRefresh = True
End Function

Public Function InitBodyPartogram() As Boolean
 '***************************************************
 '功能：初始化数据，以及进行数据展示
 '***************************************************
    Dim rs As New ADODB.Recordset
    Dim lngCount As Long, i As Integer
    On Error GoTo errHand
    
    mintPage = 1
    gblnMoved = False
    gstrSQL = "Select 数据转出 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "判断数据是否转出", T_Info.lng病人ID, T_Info.lng主页ID)
    gblnMoved = (NVL(rs!数据转出, 0) <> 0)
    
    gstrSQL = "select 项目名称 from 护理记录项目 where 保留项目=1"
    Call zlDatabase.OpenRecordset(rs, gstrSQL, "护理记录项目")
    With rs
        Do While Not .EOF
            If InStr(1, "[ZLSOFTLPF]宫口扩大[ZLSOFTLPF]先露高低[ZLSOFTLPF]生产[ZLSOFTLPF]处理[ZLSOFTLPF]", "[ZLSOFTLPF]" & NVL(rs!项目名称) & "[ZLSOFTLPF]") <> 0 Then lngCount = lngCount + 1
        .MoveNext
        Loop
    End With
    If lngCount < 4 Then
        MsgBox "产程固定项目丢失，请在护理项目管理进行检查！" & vbCrLf & _
            "项目[宫口扩大、先露高低、生产、处理]", vbInformation, gstrSysName
        Exit Function
    End If
    '提取文件分数
    lngCount = GetFileCount(T_Info.lng文件ID, T_Info.lng病人ID, T_Info.lng主页ID)
    If T_Info.lng份数 < 1 Or T_Info.lng份数 > lngCount Then T_Info.lng份数 = 1
    With cboBaby
        .Clear
        .Tag = 0
        For i = 1 To lngCount
            .AddItem i: .ItemData(.NewIndex) = i
            If i = T_Info.lng份数 Then
                .ListIndex = i - 1
            End If
        Next i
        If .ListIndex = -1 Then .ListIndex = 0
    End With
    
    InitBodyPartogram = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cboBaby_Click()
'--------------------------------------------------
'开始进行产程数据展示(调用模块 mdlPrint)
'---------------------------------------------------
    Dim strParam As String
    If T_Info.lng文件ID <= 0 Then Exit Sub
    If Val(cboBaby.Tag) = Val(cboBaby.ItemData(cboBaby.ListIndex)) Then Exit Sub
    T_Info.lng份数 = Val(cboBaby.ItemData(cboBaby.ListIndex))
    cboBaby.Tag = T_Info.lng份数
    strParam = T_Info.lng文件ID & ";" & T_Info.lng病人ID & ";" & T_Info.lng主页ID & ";" & T_Info.lng科室ID & ";" & T_Info.lng份数 & ";" & mintPage
    Call PreViewOrPrintPartogram(strParam, PicDraw, Me)
    mblnInit = True
End Sub

Private Sub Form_Load()
    If Not mblnChildForm Then
        Call RestoreWinState(Me, App.ProductName)
    End If
    Call InitMenuBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnShowState = False
    If Not mblnChildForm Then
        If Not (mfrmPartogram Is Nothing) Then UnHook mfrmPartogram.Hwnd
        Call SaveWinState(Me, App.ProductName)
    End If
End Sub

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrCustom  As CommandBarControlCustom
    
    On Error GoTo errHand
    PicDraw.AutoRedraw = True
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&E)")
                
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
               
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        cbrControl.BeginGroup = True
    End With


    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Billing, "数据编辑(&E)")
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    With cbrMenuBar.CommandBar.Controls
                
'       Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
'
'       cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
'       cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
                
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."):
        cbrControl.BeginGroup = True
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("标准", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0

    With cbrToolBar.Controls
        .Add xtpControlLabel, conMenu_View_Option, "婴儿"
        Set cbrCustom = .Add(xtpControlCustom, conMenu_View_Option, "")
        cbrCustom.Flags = xtpFlagAlignLeft
        picTmp.Visible = True
        cbrCustom.Handle = picTmp.Hwnd
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Billing, "数据"): cbrControl.ToolTipText = "产程数据编辑": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrevPage, "上页"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "上页"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextPage, "下页"):   cbrControl.ToolTipText = "下页"
    End With

    '定位工具栏
    '------------------------------------------------------------------------------------------------------------------

    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
     '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("E"), conMenu_Edit_Billing
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With
    
    InitMenuBar = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey As String
    Dim lngLoop As Long
    Dim strSQL() As String, strParam As String
    Dim blnTran As Boolean
    Dim lngIndex As Long
    Dim cbrControl As CommandBarControl
    Dim lngKey As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
        
    Select Case Control.ID
        Case conMenu_File_PrintSet   '打印设置
            
            On Error Resume Next
            Call frmPrintSet.ShowMe(Me, 1)
            
        Case conMenu_File_Preview  '打印预览
            
            Call PrintData(1)
            
        Case conMenu_File_Print  '打印
        
            Call PrintData(2)
            
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_Edit_Billing '数据编辑
            If frmPartogramEditor.ShowMe(Me, T_Info.lng文件ID, T_Info.lng病人ID, T_Info.lng主页ID, T_Info.lng科室ID, 0, mstrPrivs, (T_Info.lng编辑 = 1), IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))) = True Then
                '进行数据刷新
                If Not InitBodyPartogram Then Exit Sub
                strParam = T_Info.lng文件ID & ";" & T_Info.lng病人ID & ";" & T_Info.lng主页ID & ";" & T_Info.lng科室ID & ";" & T_Info.lng份数 & ";" & mintPage
                Call PreViewOrPrintPartogram(strParam, PicDraw, Me)
                mblnShowOk = Not mblnChildForm
            End If
        Case conMenu_Edit_PrevPage '上页
            If mintPage > 1 Then
                mintPage = mintPage - 1
                strParam = T_Info.lng文件ID & ";" & T_Info.lng病人ID & ";" & T_Info.lng主页ID & ";" & T_Info.lng科室ID & ";" & T_Info.lng份数 & ";" & mintPage
                Call PreViewOrPrintPartogram(strParam, PicDraw, Me)
            End If
        Case conMenu_Edit_NextPage '下页
            If mintPage < mintMaxPage Then
                mintPage = mintPage + 1
                strParam = T_Info.lng文件ID & ";" & T_Info.lng病人ID & ";" & T_Info.lng主页ID & ";" & T_Info.lng科室ID & ";" & T_Info.lng份数 & ";" & mintPage
                Call PreViewOrPrintPartogram(strParam, PicDraw, Me)
            End If
        Case conMenu_Help_Help
        
            Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_About
            
            Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
            
        Case conMenu_Help_Web_Home
            
            Call zlHomePage(Me.Hwnd)
            
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.Hwnd)
            
        Case conMenu_Help_Web_Mail
            
            Call zlMailTo(Me.Hwnd)
        
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup
        End Select
    End If
    
    Err = 0
    On Error Resume Next
    
    Select Case Control.ID

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Billing '数据编辑
        Control.Enabled = (T_Info.lng编辑 = 1)
        
    Case conMenu_View_StatusBar
    
        Control.Checked = Me.stbThis.Visible
    Case conMenu_Edit_PrevPage '上一页
        Control.Enabled = (mintPage > 1)
    Case conMenu_Edit_NextPage '下一页
        Control.Enabled = (mintPage < mintMaxPage)
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '客户区域的大小

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With picBack
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Top = lngTop
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub picBack_Resize()
    With vsb
        .Left = picBack.Width - .Width
        .Top = 0
        .Height = IIf(picBack.Height - hsb.Height < 0, 0, picBack.Height - hsb.Height)
    End With
    
    With hsb
        .Left = 0
        .Top = picBack.Height - .Height
        .Width = IIf(picBack.Width - vsb.Width < 0, 0, picBack.Width - vsb.Width)
    End With
    
    With picMain
        .Left = 0
        .Top = 0
        .Height = IIf(hsb.Visible = True, hsb.Top, picBack.Height)
        .Width = IIf(vsb.Visible = True, vsb.Left, picBack.Width)
    End With
End Sub

Private Function CalcScrollBarSize() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回： 调用成功返回TRUE；否则FALSE
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    '只根据没显示出来的那部分来计算步长
    msinHStep = (PicDraw.Width - picMain.Width) / 100
    msinVStep = (PicDraw.Height - picMain.Height) / 100

    
    hsb.Max = 0 - Int(0 - ((PicDraw.Width - picMain.Width) / 300)) - 1
    vsb.Max = 0 - Int(0 - ((PicDraw.Height - picMain.Height) / 300)) - 1
    hsb.Enabled = (hsb.Max > 0)
    hsb.Visible = hsb.Enabled
    vsb.Enabled = (vsb.Max > 0)
    vsb.Visible = vsb.Enabled
    
    With vsb
        .Height = picMain.Height - IIf(hsb.Visible = True, hsb.Height, 0)
    End With
    
    With hsb
        .Width = picMain.Width - IIf(vsb.Visible = True, vsb.Width, 0)
    End With
    picMain.Move 0, 0, IIf(vsb.Visible = True, vsb.Left, picBack.Width), IIf(hsb.Visible = True, hsb.Top, picBack.Height)
    
    '恒定为100,只是步长发生变化
    If hsb.Enabled Then
        hsb.Max = 100
        hsb.LargeChange = 100 / Int((Round((PicDraw.Width - picMain.Width) / picMain.Width, 2) + 1))
        hsb.SmallChange = hsb.LargeChange / 2
    End If
    
    If vsb.Enabled Then
        vsb.Max = 100
        vsb.LargeChange = 100 / Int((Round((PicDraw.Height - picMain.Height) / picMain.Height, 2) + 1))
        vsb.SmallChange = vsb.LargeChange / 2
    End If
    
    CalcScrollBarSize = True
    
End Function

Private Sub picMain_Paint()
    CalcScrollBarSize
End Sub

Private Sub picMain_Resize()
    picMain.BackColor = PicDraw.BackColor
    With PicDraw
        .Left = 0
        .Top = 0
'        .Width = IIf(.Width < picMain.Width, picMain.Width, .Width)
'        .Height = IIf(.Height < picMain.Height, picMain.Height, .Height)
    End With
    CalcScrollBarSize
End Sub

Private Sub vsb_Change()
    PicDraw.Top = -1 * vsb.Value * msinVStep
End Sub

Private Sub hsb_Change()
    PicDraw.Left = -1 * hsb.Value * msinHStep
End Sub

Public Sub PrintData(ByVal bytMode As Byte, Optional ByVal strPrintDevice As String = "")
    Dim bytOp As Byte
    Dim lngFileIndex As Long, lngFilePage As Long
    Dim blnPrint As Boolean
    
    If strPrintDevice = "" Then
        lngFileIndex = T_Info.lng份数
        lngFilePage = mintPage
        '打印选项窗体
        bytOp = frmPartogramPrintSet.PrintSet(Me, bytMode, cboBaby.ListCount, lngFileIndex, lngFilePage)
    Else
        lngFileIndex = -1
        lngFilePage = -1
    End If
    If bytOp = 0 Then Exit Sub
    '开始进行预览打印
    blnPrint = (bytOp = 2)
    
    Call ShowPrintPartogram(Me, T_Info.lng文件ID, T_Info.lng病人ID, T_Info.lng主页ID, T_Info.lng科室ID, lngFileIndex, lngFilePage, blnPrint, strPrintDevice)
End Sub
