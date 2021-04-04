VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~3.OCX"
Begin VB.Form frm对码诊断 
   Caption         =   "体检诊断对码"
   ClientHeight    =   7830
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11940
   Icon            =   "frm对照诊断.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   4860
      Left            =   3660
      ScaleHeight     =   4860
      ScaleWidth      =   4860
      TabIndex        =   2
      Top             =   1470
      Width           =   4860
      Begin zlPiesFlat.VsfGrid vsf 
         Height          =   2130
         Left            =   390
         TabIndex        =   3
         Top             =   255
         Width           =   3540
         _extentx        =   6244
         _extenty        =   3757
      End
      Begin VB.Frame fra 
         Height          =   1500
         Left            =   495
         TabIndex        =   4
         Top             =   2820
         Width           =   8085
         Begin VB.Frame fra2 
            Height          =   75
            Left            =   30
            TabIndex        =   10
            Top             =   540
            Width           =   8010
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1065
            TabIndex        =   9
            Top             =   225
            Width           =   2250
         End
         Begin VB.CommandButton cmdMenu 
            Height          =   270
            Left            =   120
            Picture         =   "frm对照诊断.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   240
            Width           =   285
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   1155
            TabIndex        =   7
            Top             =   720
            Width           =   1245
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   3450
            TabIndex        =   6
            Top             =   735
            Width           =   3840
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   2
            Left            =   1155
            TabIndex        =   5
            Top             =   1080
            Width           =   1245
         End
         Begin VB.Label lblFind 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&2.编码"
            Height          =   180
            Left            =   480
            TabIndex        =   14
            Top             =   285
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&N.干保编码"
            Height          =   180
            Index           =   1
            Left            =   180
            TabIndex        =   13
            Top             =   780
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&M.干保名称"
            Height          =   180
            Index           =   0
            Left            =   2475
            TabIndex        =   12
            Top             =   795
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&P.疾病编码"
            Height          =   180
            Index           =   2
            Left            =   180
            TabIndex        =   11
            Top             =   1140
            Width           =   900
         End
      End
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3345
      Left            =   645
      TabIndex        =   0
      Top             =   1470
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   5900
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3705
      Top             =   1605
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
            Picture         =   "frm对照诊断.frx":6AD8
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm对照诊断.frx":7072
            Key             =   "Root"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7470
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm对照诊断.frx":D8D4
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16007
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   8580
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm对照诊断.frx":E168
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm对照诊断.frx":E388
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm对照诊断.frx":E5A8
            Key             =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
      DesignerControls=   "frm对照诊断.frx":ED22
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   600
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frm对码诊断"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mlngLoop As Long
Private mstrKey As String
Private mfrmMain As Object
Private mvarParam As Variant
Private mblnEditMode As Boolean
Private mstrSvrFind As String
Private mlngRow As Long
Private mblnShowAll As Boolean
Private mblnShowOK As Boolean

Private WithEvents mobjPopMenu As clsPopMenu                '自定义弹出菜单对象
Attribute mobjPopMenu.VB_VarHelpID = -1

Private Enum mCol
    干保编码 = 5
    干保名称 = 6
    疾病编码 = 7
End Enum

Private Function InitMenuBar() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：初始化菜单、工具栏
    '------------------------------------------------------------------------------------------------------------------
    Dim cbrMenuBar As Object
    Dim obj As CommandBarControl
    Dim cbrControl As Object
    Dim cbrToolBar As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = True
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
                
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        
    End With

        
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "显示所有下级(&A)")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "显示已对码项(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
        cbrControl.BeginGroup = True
        
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."): cbrControl.BeginGroup = True
    End With
    
     '快键绑定
    With cbsThis.KeyBindings
        
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    

    '工具栏定义
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
End Function

Private Function InitClient() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：初始化窗格
    '------------------------------------------------------------------------------------------------------------------
    Dim panTab As Pane
    
    Set panTab = dkpMan.CreatePane(1, 200, 500, DockLeftOf, Nothing)
    panTab.Title = ""
    panTab.Options = PaneNoCaption
    
    Set panTab = dkpMan.CreatePane(2, 500, 200, DockRightOf, Nothing)
    panTab.Title = ""
    panTab.Options = PaneNoCaption
    
    dkpMan.SetCommandBars cbsThis
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
        
End Function

Private Function RefreshStateInfo() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能： 刷新状态栏的提示信息。
    '返回： True
    '------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String
    
    If tvw.SelectedItem Is Nothing Then
        strInfo = ""
    Else
        strInfo = "分类‘" & tvw.SelectedItem.Text & "’"
        If Val(vsf.RowData(1)) > 0 Then
            strInfo = strInfo & "下共有 " & vsf.Rows & " 个项目。"
        Else
            strInfo = strInfo & "下无项目。"
        End If
        
    End If
    
    stbThis.Panels(2).Text = strInfo
    
    RefreshStateInfo = True
    
End Function

Private Function ApplyEditColor() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能： 设置可编辑列的颜色，以示区别。
    '返回： True
    '------------------------------------------------------------------------------------------------------------------
    vsf.Cell(flexcpBackColor, 1, mCol.干保编码, vsf.Rows - 1, mCol.疾病编码) = &HFFEBD7
    ApplyEditColor = True
    
End Function

Private Function zlMenuClick(ByVal strMenuItem As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能： 实现基本的操作功能
    '参数：
    '       strMenuItem          功能名称
    '返回： 成功返回True;否则返回False
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Select Case strMenuItem
    Case "分类数据"
        
        tvw.Nodes.Clear
        vsf.Rows = 2
        vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
        vsf.RowData(1) = 0
        tvw.Nodes.Add , , "_0", "所有分类", "Root", "Root"
        gstrSQL = "SELECT 序号 AS ID,DECODE(上级序号,NULL,0,上级序号) AS 上级ID,名称,2 AS 排序,编码 " & _
                "FROM 体检诊断建议 " & _
                "WHERE NVL(末级,0)=0 " & _
                "START WITH 上级序号 IS NULL " & _
                "CONNECT BY PRIOR 序号=上级序号  order by 排序,编码 "
        Call OpenRecordSet(rs)
        
        Do Until rs.EOF
            If IsNull(rs("上级id")) Then
                tvw.Nodes.Add "_0", tvwChild, "_" & rs("id"), "【" & rs("编码") & "】" & rs("名称"), "Class", "Class"
            Else
                tvw.Nodes.Add "_" & rs("上级id"), tvwChild, "_" & rs("id"), "【" & rs("编码") & "】" & rs("名称"), "Class", "Class"
            End If
            rs.MoveNext
        Loop
        
        
    Case "明细数据"
        
        vsf.Rows = 2
        vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
        vsf.RowData(1) = 0
    
        If tvw.SelectedItem Is Nothing Then Exit Function
            
        gstrSQL = " SELECT E.序号 AS ID," & _
                          "E.编码," & _
                          "E.名称," & _
                          "E.简码," & _
                          "E.诊断建议," & _
                          "F.名称 AS 所属分类 "
        
        If mblnShowAll Then
            
            gstrSQL = gstrSQL & " From (Select 序号 As ID,名称 From 体检诊断建议 Where Nvl(末级,0)=0 Connect by Prior 序号=上级序号"
            
            If Val(Mid(tvw.SelectedItem.Key, 2)) = 0 Then
                gstrSQL = gstrSQL & "  Start With 上级序号 IS NULL) F,"
            Else
                gstrSQL = gstrSQL & "  Start With 序号 = " & Val(Mid(tvw.SelectedItem.Key, 2)) & ") F,"
            End If
        Else
            gstrSQL = gstrSQL & " From (Select 序号 As ID,名称 From 体检诊断建议 Where 序号=" & Val(Mid(tvw.SelectedItem.Key, 2)) & ") F,"
        End If
        
        gstrSQL = gstrSQL & _
                        "体检诊断建议 E " & _
                    "Where F.ID=E.上级序号 " & _
                        "And E.末级 = 1 "
                    
        If mblnShowOK Then
            gstrSQL = "Select A.*,C.干保编码,C.干保名称,C.疾病编码 From (" & gstrSQL & ") A,体检诊断建议_干保 C Where A.ID=C.结论id"
        Else
            gstrSQL = "Select A.*,C.干保编码,C.干保名称,C.疾病编码 From (" & gstrSQL & ") A,体检诊断建议_干保 C Where A.ID=C.结论id(+)"
        End If
        
        
        Call OpenRecordSet(rs, Me.Caption)
        If rs.BOF = False Then
            
            Call FillGrid(vsf, rs)
            
        End If
        
    End Select
    
    zlMenuClick = True
    
    Exit Function
    
errHand:

    ShowSimpleMsg Err.Description

End Function

Private Function CheckValid() As Boolean
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    Dim strCode As String

    lngKey = Val(vsf.RowData(vsf.Row))
    strCode = Trim(txt(1).Text)

    '检查唯一性
    gstrSQL = "Select 1 From 体检诊断建议_干保 Where 结论id<>" & lngKey & " And 干保编码='" & strCode & "'"
    rs.Open gstrSQL, gcnOracle
    If rs.BOF = False Then

        ShowSimpleMsg "此码[" & strCode & "]已经对应，不能一码对应多个项目！"

        vsf.Row = vsf.Row
        vsf.Col = mCol.干保编码
        vsf.ShowCell vsf.Row, vsf.Col

        DoEvents
        LocationObj txt(1)

        Exit Function

    End If
    
    CheckValid = True
    
End Function

Private Function SaveData() As Boolean
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim lngKey As Long
    Dim strCode As String
    Dim blnTran As Boolean
    
    On Error GoTo errHand
    
    lngKey = Val(vsf.RowData(vsf.Row))
    strCode = Trim(vsf.TextMatrix(vsf.Row, mCol.干保编码))
    
    If lngKey > 0 Then
        
        blnTran = True
        gcnOracle.BeginTrans
        
        strSQL = "Delete From 体检诊断建议_干保 Where 结论id=" & lngKey
        gcnOracle.Execute strSQL

        If strCode <> "" Then
            
            strSQL = "Insert Into 体检诊断建议_干保(结论id,干保编码,干保名称,疾病编码) Values (" & lngKey & ",'" & strCode & "','" & Trim(vsf.TextMatrix(vsf.Row, mCol.干保名称)) & "','" & Trim(vsf.TextMatrix(vsf.Row, mCol.疾病编码)) & "')"
            gcnOracle.Execute strSQL
  
        End If
        
        gcnOracle.CommitTrans
        blnTran = False
        
    End If
    
    SaveData = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg Err.Description
    
    If blnTran Then gcnOracle.RollbackTrans
End Function

Private Function InitData() As Boolean
    
    With vsf
        
        .Cols = 0
        .NewColumn "", 255, 4
                
        .NewColumn "名称", 1500, 1
        .NewColumn "编码", 900, 1
        .NewColumn "简码", 900, 1
        .NewColumn "所属分类", 1800, 1
        .NewColumn "干保编码", 900, 1, , 1, GetMaxLength("体检诊断建议_干保", "干保编码")
        .NewColumn "干保名称", 1500, 1, , 1, GetMaxLength("体检诊断建议_干保", "干保名称")
        .NewColumn "疾病编码", 1500, 1, , 1, GetMaxLength("体检诊断建议_干保", "疾病编码")
        
        .NewColumn "", 15, 1
        
        .ExtendLastCol = True
        .FixedCols = 1
        .Body.GridColor = &HC1C1C1
        .Body.GridColorFixed = &HC1C1C1
        .Body.GridLines = flexGridFlat
        .Body.BackColorFixed = .Body.BackColorBkg
        
        .Body.Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = &H8000000F
        .AppendRow = True
        
        If mblnEditMode = False Then
            .EditMode(mCol.干保编码) = 0
            .EditMode(mCol.干保名称) = 0
            .EditMode(mCol.疾病编码) = 0
        End If
        
    End With
    
    txt(1).MaxLength = GetMaxLength("体检诊断建议_干保", "干保编码")
    txt(0).MaxLength = GetMaxLength("体检诊断建议_干保", "干保名称")
    txt(2).MaxLength = GetMaxLength("体检诊断建议_干保", "疾病编码")
    
    txt(0).Enabled = mblnEditMode
    txt(1).Enabled = mblnEditMode
    txt(2).Enabled = mblnEditMode
    txt(0).BackColor = IIf(mblnEditMode, &H80000005, &H8000000F)
    txt(1).BackColor = IIf(mblnEditMode, &H80000005, &H8000000F)
    txt(2).BackColor = IIf(mblnEditMode, &H80000005, &H8000000F)
    
    InitData = True
    
End Function



Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As Object

    On Error GoTo errHand
    
    Select Case Control.ID
        Case conMenu_View_ToolBar_Button
        
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text
        
            For Each cbrControl In cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            
            cbsThis.RecalcLayout
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
                
                
        If Not (tvw.SelectedItem Is Nothing) Then
            mstrKey = ""
            Call tvw_NodeClick(tvw.SelectedItem)
        End If
    
        Case conMenu_View_Expend_CurExpend
'
            mblnShowAll = Not mblnShowAll
            If Not (tvw.SelectedItem Is Nothing) Then
                mstrKey = ""
                Call tvw_NodeClick(tvw.SelectedItem)
            End If
        
        Case conMenu_View_Expend_AllExpend
            
            mblnShowOK = Not mblnShowOK
            
            If Not (tvw.SelectedItem Is Nothing) Then
                mstrKey = ""
                Call tvw_NodeClick(tvw.SelectedItem)
            End If
            
        Case conMenu_View_Refresh
            
            Call RefreshData
                        
        Case conMenu_Help_Help
        
            Call ShowHelp(Me.hWnd, Me.Name)
        
        Case conMenu_Help_About
            
            frmAbout.Show 1, Me
            
        Case conMenu_File_Exit
        
            Unload Me
            Exit Sub
            
    End Select
    
    
    cbsThis.RecalcLayout
    
    Exit Sub
    
errHand:
    
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button
        Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_StatusBar
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Expend_CurExpend
    
        Control.Checked = mblnShowAll
    Case conMenu_View_Expend_AllExpend
        
        Control.Checked = mblnShowOK
        
    End Select
    
    
End Sub

Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenuByCursor
    
    txtFind.Text = ""
    
    LocationObj txtFind
    
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    
    Select Case Item.ID
    Case 1
        Item.Handle = tvw.hWnd
    Case 2
       Item.Handle = picContainer.hWnd
    End Select
    
End Sub

Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    If InitData = False Then
        Unload Me
        Exit Sub
    End If
    
    DoEvents
    
    Call RefreshData
        
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    Call InitMenuBar
    Call InitClient
    
    mblnShowAll = True
    mblnShowOK = False
    
    Call RestoreWinState(Me, App.ProductName)
    
    mblnEditMode = (InStr(gstrPrive, ";数据对码;") > 0)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub RefreshData()
    
    Dim strTvwKey As String
    Dim strVsfKey As String
    
    If Not (tvw.SelectedItem Is Nothing) Then strTvwKey = tvw.SelectedItem.Key
    strVsfKey = Val(vsf.RowData(vsf.Row))
    
    '装载分类数据
    Call zlMenuClick("分类数据")
    
    On Error Resume Next
    
    tvw.Nodes(strTvwKey).Selected = True
    tvw.Nodes(strTvwKey).EnsureVisible
    
    On Error GoTo 0
    
    If tvw.SelectedItem Is Nothing Then
        If tvw.Nodes.Count > 0 Then
            tvw.Nodes(1).Selected = True
            tvw.Nodes(1).EnsureVisible
            tvw.Nodes(1).Expanded = True
        End If
    End If
    
    If Not (tvw.SelectedItem Is Nothing) Then
        '装载明细数据
        Call zlMenuClick("明细数据")
                        
        If Val(strVsfKey) > 0 Then
            For mlngLoop = 1 To vsf.Rows - 1
                If Val(vsf.RowData(mlngLoop)) = Val(strVsfKey) Then
                    vsf.Row = mlngLoop
                    vsf.ShowCell vsf.Row, vsf.Col
                    Exit For
                End If
            Next
        End If
        Call RefreshStateInfo
        Call ApplyEditColor
    End If
    
End Sub

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    Dim strChar As String
    Dim intIndex As Integer
    
    strChar = "123456789ABCDEFGHIJKLMNOPQUVRSTWXYZ"
    
    For mlngLoop = 0 To vsf.Cols - 1
        
        If Trim(vsf.TextMatrix(0, mlngLoop)) <> "" Then
            
            intIndex = intIndex + 1
            
            mobjPopMenu.Add intIndex, "&" & Mid(strChar, intIndex, 1) & "." & Trim(vsf.TextMatrix(0, mlngLoop))
            
        End If
        
    Next

End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)

    lblFind.Caption = Caption
    
    txtFind.Left = lblFind.Left + lblFind.Width + 60
    
   
End Sub


Private Sub picContainer_Resize()
    On Error Resume Next
    
    With vsf
        .Left = 0
        .Top = 0
        .Width = picContainer.Width - .Left
        .Height = picContainer.Height - fra.Height + 60 - .Top
    End With
    
    With fra
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height - 60
        .Width = vsf.Width
    End With
    
    fra2.Left = 0
    fra2.Width = fra.Width
End Sub


Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)

    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    
    Call zlMenuClick("明细数据")
    Call RefreshStateInfo
    
    vsf.AppendRow = True
    
    Call ApplyEditColor

End Sub

Private Sub txt_GotFocus(Index As Integer)
    TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCol As String
    Dim lngCol As Long
    
    If KeyAscii = vbKeyReturn Then
        
        If CheckValid = False Then
            Exit Sub
        End If
        
        If Index = 1 Then vsf.TextMatrix(vsf.Row, mCol.干保编码) = txt(Index)
        If Index = 0 Then vsf.TextMatrix(vsf.Row, mCol.干保名称) = txt(Index)
        If Index = 2 Then vsf.TextMatrix(vsf.Row, mCol.疾病编码) = txt(Index)
        
        If SaveData Then
            If Index = 2 Then
                txtFind.SetFocus
            Else
                SendKeys "{TAB}"
            End If
        End If
        
    End If
    
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    
    Cancel = Not StrIsValid(txt(Index), txt(Index).MaxLength)
    
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strCol As String
    Dim lngCol As Long
    
    Dim lngLoop As Long
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    Dim lngRow As Long
    
    If KeyAscii = vbKeyReturn Then
        
        If Trim(txtFind.Text) <> "" Then
            
            strCol = Mid(lblFind.Caption, 4)
            lngCol = GetCol(vsf, strCol)
            
            If lngCol < 0 Then Exit Sub
            
            If mstrSvrFind <> txtFind.Text Then
                
                mstrSvrFind = txtFind.Text
                
                For lngLoop = 1 To vsf.Rows - 1
                    If InStr(UCase(vsf.TextMatrix(lngLoop, lngCol)), UCase(mstrSvrFind)) > 0 Then
                        mlngRow = lngLoop
                        Exit For
                    End If
                Next
                If lngLoop = vsf.Rows Then mlngRow = -1
            Else
                
                For lngLoop = mlngRow + 1 To vsf.Rows - 1
                    If InStr(UCase(vsf.TextMatrix(lngLoop, lngCol)), UCase(mstrSvrFind)) > 0 Then
                        mlngRow = lngLoop
                        Exit For
                    End If
                Next
                
                If lngLoop = vsf.Rows Then mlngRow = -1
            End If
            
            If mlngRow = -1 Then
                ShowSimpleMsg "已经查找完，如再查找将重新搜索一次！"
                mlngRow = 0
                DoEvents
            Else
                vsf.Row = mlngRow
                vsf.ShowCell vsf.Row, vsf.Col
                
                txt(1).Text = vsf.TextMatrix(vsf.Row, mCol.干保编码)
                txt(0).Text = vsf.TextMatrix(vsf.Row, mCol.干保名称)
                txt(2).Text = vsf.TextMatrix(vsf.Row, mCol.疾病编码)
                
                SendKeys "{TAB}"
            End If
            
        End If
        
        txtFind.SetFocus
        TxtSelAll txtFind
   
    End If
End Sub

Private Sub txtFind_GotFocus()
    TxtSelAll txtFind
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    If mblnEditMode Then Call SaveData
        
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngCol As Long

    If OldRow <> NewRow Then

        lngCol = GetCol(vsf, "干保编码")

        On Error Resume Next

        If OldRow + 1 > vsf.FixedRows Then
            vsf.Cell(flexcpBackColor, OldRow, vsf.FixedCols, OldRow, lngCol - 1) = vsf.Body.BackColor
            vsf.Cell(flexcpBackColor, OldRow, lngCol + 3, OldRow, vsf.Cols - 1) = vsf.Body.BackColor

            vsf.Cell(flexcpForeColor, OldRow, vsf.FixedCols, OldRow, lngCol - 1) = vsf.Body.ForeColor
            vsf.Cell(flexcpForeColor, OldRow, lngCol + 3, OldRow, vsf.Cols - 1) = vsf.Body.ForeColor
        End If

        If NewRow + 1 > vsf.FixedRows Then
            vsf.Cell(flexcpBackColor, NewRow, vsf.FixedCols, NewRow, lngCol - 1) = vsf.Body.BackColorSel
            vsf.Cell(flexcpBackColor, NewRow, lngCol + 3, NewRow, vsf.Cols - 1) = vsf.Body.BackColorSel

            vsf.Cell(flexcpForeColor, NewRow, vsf.FixedCols, NewRow, lngCol - 1) = &H80000005
            vsf.Cell(flexcpForeColor, NewRow, lngCol + 3, NewRow, vsf.Cols - 1) = &H80000005

        End If

    End If
    
    If vsf.Col < mCol.干保编码 Then vsf.Col = mCol.干保编码
    If vsf.Col > mCol.疾病编码 Then vsf.Col = mCol.疾病编码
    
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_GotFocus()
    mlngRow = -1
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    Dim strCode As String
    
    If Col = mCol.干保编码 Then
        lngKey = Val(vsf.RowData(vsf.Row))
        strCode = Trim(vsf.EditText)
    
        '检查唯一性
        gstrSQL = "Select 1 From 体检诊断建议_干保 Where 结论id<>" & lngKey & " And 干保编码='" & strCode & "'"
        rs.Open gstrSQL, gcnOracle
        If rs.BOF = False Then
    
            ShowSimpleMsg "此码[" & strCode & "]已经对应，不能一码对应多个项目！"
    
            Cancel = True
    
        End If
    End If
    
End Sub


