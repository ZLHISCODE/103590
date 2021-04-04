VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPatholSpecialExamined_WorkPrint 
   Caption         =   "特检批量处理"
   ClientHeight    =   8280
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   11685
   Icon            =   "frmPatholSpecialExamined_WorkPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11685
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7215
      ScaleWidth      =   11415
      TabIndex        =   2
      Top             =   480
      Width           =   11415
      Begin VB.OptionButton optAll 
         Caption         =   "所 有"
         Height          =   180
         Left            =   6480
         TabIndex        =   21
         Top             =   1080
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optXiMu2 
         Caption         =   "多药耐药"
         Height          =   180
         Left            =   5160
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optXiMu1 
         Caption         =   "鉴 别"
         Height          =   180
         Left            =   3960
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Frame framFilter 
         Height          =   735
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   11415
         Begin VB.CheckBox chkMoney 
            Caption         =   "计费"
            Height          =   255
            Left            =   9120
            TabIndex        =   12
            Top             =   280
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.TextBox txtStartPatholNum 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5880
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtEndPatholNum 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7600
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdFilter 
            Caption         =   "查 询(&Q)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   9960
            TabIndex        =   9
            Top             =   200
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtpStartRequisition 
            Height          =   300
            Left            =   1080
            TabIndex        =   13
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd hh:mm"
            Format          =   97058819
            CurrentDate     =   40679.0594097222
         End
         Begin MSComCtl2.DTPicker dtpEndRequisition 
            Height          =   300
            Left            =   3120
            TabIndex        =   14
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   97058819
            CurrentDate     =   40679.0594097222
         End
         Begin VB.Label Label2 
            Caption         =   "到"
            Height          =   255
            Left            =   7360
            TabIndex        =   18
            Top             =   285
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "到"
            Height          =   255
            Left            =   2920
            TabIndex        =   17
            Top             =   285
            Width           =   255
         End
         Begin VB.Label Label7 
            Caption         =   "申请时间："
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   280
            Width           =   975
         End
         Begin VB.Label labPatholNum 
            Caption         =   "病理号："
            Height          =   255
            Left            =   5160
            TabIndex        =   15
            Top             =   285
            Width           =   720
         End
      End
      Begin VB.Frame framSpeExam 
         Caption         =   "制片批量处理"
         Height          =   5055
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   11295
         Begin zl9PACSWork.ucFlexGrid ufgData 
            Height          =   4695
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   8281
            IsKeepRows      =   0   'False
            BackColor       =   12648447
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            Editable        =   0
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
      End
      Begin VB.CheckBox chkYSQ 
         Caption         =   "已申请"
         Height          =   255
         Left            =   4560
         TabIndex        =   5
         ToolTipText     =   "显示制片状态为“未处理”的制片记录。"
         Top             =   6600
         Width           =   855
      End
      Begin VB.CheckBox chkYJS 
         Caption         =   "已接受"
         Height          =   180
         Left            =   5520
         TabIndex        =   4
         ToolTipText     =   "显示制片状态为“已接受”的制片记录。"
         Top             =   6630
         Width           =   855
      End
      Begin VB.CheckBox chkYWC 
         Caption         =   "已完成"
         Height          =   180
         Left            =   6480
         TabIndex        =   3
         ToolTipText     =   "显示制片状态为“未完成”的制片记录。"
         Top             =   6630
         Width           =   855
      End
      Begin XtremeSuiteControls.TabControl tabFilter 
         Height          =   375
         Left            =   -120
         TabIndex        =   22
         Top             =   960
         Width           =   11445
         _Version        =   589884
         _ExtentX        =   20188
         _ExtentY        =   661
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picTag 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7920
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholSpecialExamined_WorkPrint.frx":179A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13732
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholSpecialExamined_WorkPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1

Private mlngPatholAdivceId As Long
Private mufgGrid As ucFlexGrid
Private mstrPrivs As String

Private mblnAutoAcceptOfAfterPrint As Boolean '打印后自动接受


Public blnIsOk As Boolean


Private mlngFilterTabIndex As Long

Private Enum TMenuType
    mtLab = 1
    mtLabView = 10
    mtLabPrint = 11
    
    mtWork = 2
    mtWorkView = 20
    mtWorkPrint = 21
    
    mtAccept = 3
    mtComplete = 4
    mtchkYSQ = 5
    mtchkYJS = 6
    mtchkYWC = 7
End Enum

Public Sub ShowWorkPrint(ufgGrid As ucFlexGrid, ByVal lngPatholAdivceId As Long, _
    ByVal lngCurSpeExamType As Long, ByVal strPrivs As String, owner As Form)
'显示工作清单打印窗口
    Set mufgGrid = ufgGrid

    mlngPatholAdivceId = lngPatholAdivceId
    mstrPrivs = strPrivs
    blnIsOk = False

    '配置当前特检类型
'    Call ConfigSpeExamType(lngCurSpeExamType)
    
    Call ConfigSpeExamPopedom
    
'    '载入特检数据
'    If lngPatholAdivceId > 0 Then
'        Call LoadSpeExamData(lngPatholAdivceId, lngCurSpeExamType)
'    End If
    
    '刷新数量
    Call RefreshSilcesCount
    
    Call Me.Show(1, owner)
    
End Sub


'Private Sub ConfigSpeExamType(ByVal strCurSpeExamType As String)
''配置当前特检类型
'    Dim i As Long
'
'    For i = 0 To tabFilter.ItemCount - 1
'        If tabFilter(i).Tag Like "*" & strCurSpeExamType & "*" Then
'            tabFilter(i).Selected = True
'            Exit Sub
'        End If
'    Next i
'End Sub


Private Sub ConfigSpeExamPopedom()
'配置特检权限，隐藏没有权限的标签
    Dim blnIsPopedom As Boolean
    
    blnIsPopedom = CheckPopedom(mstrPrivs, "免疫组化")
    tabFilter(0).Visible = blnIsPopedom
    
    blnIsPopedom = CheckPopedom(mstrPrivs, "特殊染色")
    tabFilter(1).Visible = blnIsPopedom
    
    blnIsPopedom = CheckPopedom(mstrPrivs, "分子病理")
    tabFilter(2).Visible = blnIsPopedom
End Sub


Private Sub InitSpeExamWorkList()
'初始化特检工作清单显示列表
'    ufgData.DataGrid.MergeCells = flexMergeRestrictRows
    Dim strTemp As String
    
        '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    
    '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("特检清单列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        ufgData.ColNames = gstrSpeExamWorkCols
    Else
        ufgData.ColNames = strTemp
    End If

    ufgData.DefaultColNames = gstrSpeExamWorkCols
    ufgData.ColConvertFormat = gstrSpeExamWorkConvertFormat
    ufgData.IsShowPopupMenu = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo ErrorHand
    
    Select Case control.ID
        Case TMenuType.mtLabView                    '标签预览
            Call Menu_File_LabView(control)
            
        Case TMenuType.mtLabPrint                   '标签打印
            Call Menu_File_LabPrint(control)
        
        Case TMenuType.mtWorkView                   '标签预览
            Call Menu_File_WorkView(control)
        
        Case TMenuType.mtWorkPrint                  '清单打印
            Call Menu_File_WorkPrint(control)
        
        Case TMenuType.mtAccept                     '特检接受
            Call Menu_Edit_Accept
        
        Case TMenuType.mtComplete                   '特检完成
            Call Menu_Edit_Complate
        
        Case conMenu_File_Exit                      '退出
            Call Menu_File_Exit
            
'---------------------------查看----------------
        Case conMenu_View_ToolBar_Button            '工具栏
            Call Menu_View_ToolBar_Button_click(control)

        Case conMenu_View_ToolBar_Text              '按钮文字
            Call Menu_View_ToolBar_Text_click(control)

        Case conMenu_View_StatusBar                 '状态栏
            Call Menu_View_StatusBar_click(control)
            
'--------------------------帮助-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click

        Case conMenu_Help_Web_Forum
            Call Menu_Help_Web_Forum_click

        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click

        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click

        Case conMenu_Help_About
            Call Menu_Help_About_click
    End Select
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_Exit()
On Error Resume Next
    blnIsOk = False
    Call Unload(Me)
End Sub

Private Sub Menu_Help_About_click()
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub Menu_Help_Web_Mail_click()
    zlMailTo hWnd
End Sub

Private Sub Menu_Help_Web_Home_click()
    zlHomePage hWnd
End Sub

Private Sub Menu_Help_Web_Forum_click()
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next
    
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHand
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).Style
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If
        
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = intStyle
        Next
    Next
    
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
    
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_Help_Help_click()
    '功能：调用帮助主题
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible = True Then Bottom = stbThis.Height
End Sub

Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
On Error Resume Next
    picBack.Left = Left
    picBack.Top = Top
    picBack.Width = Right - Left
    picBack.Height = Bottom - Top
End Sub

Private Sub chkYJS_Click()
On Error GoTo errHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkYSQ_Click()
On Error GoTo errHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkYWC_Click()
On Error GoTo errHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optAll_Click()
On Error GoTo errHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optXiMu1_Click()
On Error GoTo errHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optXiMu2_Click()
On Error GoTo errHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tabFilter_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errHandle
    Call ConfigSpeexamDetail(Item.Tag)
    
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSpeExamData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Exit Sub
End Sub


Private Sub ConfigSpeexamDetail(ByVal strSpeExamTag As String)
'配置特检细目

    optXiMu1.Visible = True
    optXiMu2.Visible = True
    optAll.Visible = True
            
    Select Case Val(strSpeExamTag)
        Case 0
            optXiMu1.Caption = "鉴 别"
            optXiMu2.Caption = "多药耐药"
            
        Case 1
            optXiMu1.Visible = False
            optXiMu2.Visible = False
            optAll.Visible = False
            
        Case 2
            optXiMu1.Caption = "荧光分子"
            optXiMu2.Caption = "普通分子"

    End Select
End Sub



Private Sub ufgData_OnColFormartChange()
 '保存列表参数
     zlDatabase.SetPara "特检清单列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub


Private Sub RefreshSilcesCount()
'刷新制片记录数量
    Dim i As Long
    Dim lngFinishCount As Long
    Dim lngNeedCount As Long

    lngFinishCount = 0
    lngNeedCount = 0


    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.RowHidden(i) Then
            If Not ufgData.IsNullRow(i) Then

                If ufgData.Text(i, gstrSlices_当前状态) <> "已完成" Then
                    lngNeedCount = lngNeedCount + 1
                Else
                    lngFinishCount = lngFinishCount + 1
                End If
            End If
        End If
    Next i

    stbThis.Panels(2).Text = "已完成项目数：" & lngFinishCount & "    需检查项目数：" & lngNeedCount
    
End Sub



Private Sub LoadSpeExamData(ByVal lngPatholAdivceId As Long, Optional ByVal lngSpeExamType As Long = -1)
'载入特检数据
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select a.Id,c.检查类型,c.病理号,c.病理医嘱ID, e.姓名, a.材块ID, b.序号, b.标本名称, a.特检类型, a.抗体id,  d.抗体名称, a.制作类型,a.当前状态,a.清单状态 " & _
                " from 病理特检信息 a, 病理取材信息 b, 病理检查信息 c, 病理抗体信息 d, 病人医嘱记录 e " & _
                " Where a.材块id = b.材块id And b.病理医嘱ID = c.病理医嘱ID And c.医嘱ID = e.ID And a.抗体id = d.抗体id " & _
                " and c.病理医嘱ID=[1] " & IIf(lngSpeExamType >= 0, " and 特检类型=[2]", "") & " and a.当前状态 <> 2 and a.清单状态=0" & _
                " order by a.特检类型,b.病理号,a.当前状态,b.序号,a.Id "
    'If mblnMoved Then strSql = GetMovedDataSql(strSql)
                
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdivceId, lngSpeExamType)
    
    Call ufgData.RefreshData
End Sub


Private Sub GetSpeExamData()
'根据过滤条件查询特检数据
    Dim strSql As String
    Dim strPatholNumQuery As String
    Dim rsData As ADODB.Recordset
    
    
    strPatholNumQuery = ""
    If Trim(txtStartPatholNum.Text) <> "" And Trim(txtEndPatholNum.Text) <> "" Then
        strPatholNumQuery = " and (REGEXP_SUBSTR(upper(c.病理号), '[[:alpha:]]+') >=REGEXP_SUBSTR(upper([3]),'[[:alpha:]]+') and to_number(REGEXP_SUBSTR(upper(c.病理号), '[[:digit:]]+')) >=to_number(REGEXP_SUBSTR(upper([3]),  '[[:digit:]]+'))) "
        strPatholNumQuery = strPatholNumQuery & " and  (REGEXP_SUBSTR(upper(c.病理号), '[[:alpha:]]+') <=REGEXP_SUBSTR(upper([4]),'[[:alpha:]]+') and to_number(REGEXP_SUBSTR(upper(c.病理号),  '[[:digit:]]+')) <=to_number(REGEXP_SUBSTR(upper([4]), '[[:digit:]]+'))) "
    ElseIf Trim(txtStartPatholNum.Text) <> "" Then
        strPatholNumQuery = " and upper(c.病理号)=upper([3]) "
    ElseIf Trim(txtStartPatholNum.Text) <> "" Then
        strPatholNumQuery = " and upper(c.病理号) =upper([4]) "
    End If
    
    
    strSql = "select * from (select /*+ Rule*/ a.Id,c.检查类型,c.病理号,c.病理医嘱ID, e.姓名, a.材块ID,b.序号, b.标本名称, a.特检类型,a.特检细目, a.抗体id,  d.抗体名称, a.制作类型,a.当前状态,a.清单状态,f.补费状态, " & _
                " (select count(*) from 病人医嘱附费 X,门诊费用记录 Y where X.记录性质=Y.记录性质 and X.no = Y.no and Y.记录状态=0 and X.医嘱Id=c.医嘱ID) as 附费, f.申请时间, a.完成时间" & _
                " from 病理特检信息 a, 病理取材信息 b, 病理检查信息 c, 病理抗体信息 d, 病人医嘱记录 e ,病理申请信息 f " & _
                " Where a.材块id = b.材块id And b.病理医嘱ID = c.病理医嘱ID And c.医嘱ID = e.ID And a.抗体id = d.抗体id and a.申请ID=f.申请ID and f.申请时间 between [1] and [2] " & _
                IIf(strPatholNumQuery <> "", strPatholNumQuery, "") & ")" & _
                IIf(chkMoney.value <> 0, " where 补费状态<>1 and 附费<=0 ", "") & _
                " order by 特检类型,病理号, 特检细目,当前状态,序号,ID "
    'If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
                
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                            CDate(dtpStartRequisition.value), _
                                            CDate(dtpEndRequisition.value), _
                                            txtStartPatholNum.Text, _
                                            txtEndPatholNum.Text)
                                            
                                                                    
    Call FilterSpeExamData
End Sub


Private Sub FilterSpeExamData()
'过滤查询出来的制片数据
    Dim strFilters As String
    Dim strStudyTypeFilter As String
    
    strFilters = ""
    
    
    strStudyTypeFilter = ""
    
    '特检细目：0-无，1-鉴别，2-多药耐药，3-荧光分子，4-普通分子
    Select Case Val(tabFilter.Selected.Tag)
        Case 0
            'optXiMu1表示是否选择免疫鉴别
            If optXiMu1.value Then
                strStudyTypeFilter = "特检类型=0 and 特检细目=1"
            ElseIf optXiMu2.value Then
                strStudyTypeFilter = "特检类型=0 and 特检细目=2"
            Else
                strStudyTypeFilter = "特检类型=0"
            End If
            
            
        Case 1
            strStudyTypeFilter = "特检类型=1"
            
        Case 2
            'optXiMu1表示是否选择荧光分子
            If optXiMu1.value Then
                strStudyTypeFilter = "特检类型=2 and 特检细目=3"
            ElseIf optXiMu2.value Then
                strStudyTypeFilter = "特检类型=2 and 特检细目=4"
            Else
                strStudyTypeFilter = "特检类型=2"
            End If
    End Select
    
        
    If chkYSQ.value <> 0 Then
        strFilters = "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " 当前状态=0)"
    End If
    
    If chkYJS.value <> 0 Then
        If strFilters <> "" Then strFilters = strFilters & " or "
        strFilters = strFilters & "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " 当前状态=1)"
    End If
    
    If chkYWC.value <> 0 Then
        If strFilters <> "" Then strFilters = strFilters & " or "
        strFilters = strFilters & "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " 当前状态=2)"
    End If
    
    '如果三种状态都不勾选，则显示当前特检类型下所有记录
    If chkYSQ.value = 0 And chkYJS.value = 0 And chkYWC.value = 0 Then
        strFilters = "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " 当前状态=0)" & " or " & _
                     "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " 当前状态=1)" & " or " & _
                     "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " 当前状态=2)"
    End If
    
    If ufgData.AdoData Is Nothing Then Exit Sub
    
    ufgData.AdoData.Filter = strFilters
    ufgData.RefreshData
    
    Call RefreshSilcesCount
End Sub


Private Sub picBack_Resize()
'调整窗口布局
    On Error Resume Next
    
    framFilter.Left = 120
    framFilter.Top = 0
    framFilter.Width = picBack.Width - 240
    
    tabFilter.Left = 120
    tabFilter.Top = framFilter.Top + framFilter.Height
    tabFilter.Width = picBack.Width - 240
    
    optXiMu1.Left = dtpEndRequisition.Left + 120
    optXiMu1.Top = tabFilter.Top + 90
    
    optXiMu2.Left = optXiMu1.Left + optXiMu1.Width + 120
    optXiMu2.Top = optXiMu1.Top
    
    optAll.Left = optXiMu2.Left + optXiMu2.Width + 120
    optAll.Top = optXiMu1.Top

    chkYSQ.Left = optAll.Left + optAll.Width + 720
    chkYSQ.Top = optAll.Top - 20
    
    chkYJS.Left = chkYSQ.Left + chkYSQ.Width + 240
    chkYJS.Top = chkYSQ.Top + 40
    
    chkYWC.Left = chkYJS.Left + chkYJS.Width + 240
    chkYWC.Top = chkYJS.Top
    
    framSpeExam.Left = 120
    framSpeExam.Top = tabFilter.Top + tabFilter.Height
    framSpeExam.Width = picBack.Width - 240
    framSpeExam.Height = picBack.Height - framFilter.Height - tabFilter.Height - 60
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framSpeExam.Width - 240
    ufgData.Height = framSpeExam.Height - 300
End Sub



Private Sub SpeExamBatAccept()
'特检批量接受
    Dim i As Long
    Dim curPatholAdviceID As String
    Dim strSql As String
    Dim blnUpdateCallWind As Boolean
    
    blnUpdateCallWind = False
    
    For i = 1 To ufgData.GridRows - 1
        '如果选中的检查，才进行接收
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            If curPatholAdviceID <> ufgData.Text(i, gstrSpeExamWork_病理医嘱ID) Then
                curPatholAdviceID = ufgData.Text(i, gstrSpeExamWork_病理医嘱ID)
                
                strSql = "Zl_病理特检_接受(" & Val(curPatholAdviceID) & "," & Val(tabFilter.Selected.Tag) & ",'" & UserInfo.姓名 & "')"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
            
            '更新当前列表状态
            If ufgData.Text(i, gstrSpeExamWork_当前状态) = "已申请" Then
                ufgData.Text(i, gstrSpeExamWork_当前状态) = "已接受"
            End If
            
            '更新调用界面列表状态
            If Val(curPatholAdviceID) = mlngPatholAdivceId Then
                blnUpdateCallWind = True
            End If
        End If
    Next i
    
    If blnUpdateCallWind And Not (mufgGrid Is Nothing) Then
        For i = 1 To mufgGrid.GridRows - 1
            If mufgGrid.Text(i, gstrSpeExam_当前状态) = "已申请" Then
                Call mufgGrid.SyncText(i, gstrSpeExam_当前状态, "已接受", True)
                Call mufgGrid.SyncText(i, gstrSpeExam_特检医师, UserInfo.姓名, True)
            End If
        Next i
    End If
End Sub




Private Sub SpeExamBatSure()
'特检批量接受
    Dim i As Long
    Dim curPatholAdviceID As String
    Dim strSql As String
    Dim blnUpdateCallWind As Boolean
    Dim dtServicesTime As Date
    
    dtServicesTime = zlDatabase.Currentdate
    blnUpdateCallWind = False
    
    For i = 1 To ufgData.GridRows - 1
        '如果选中的检查，才进行接收
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            If curPatholAdviceID <> ufgData.Text(i, gstrSpeExamWork_病理医嘱ID) Then
                curPatholAdviceID = ufgData.Text(i, gstrSpeExamWork_病理医嘱ID)
                
                strSql = "Zl_病理特检_确认(" & Val(curPatholAdviceID) & "," & Val(tabFilter.Selected.Tag) & "," & To_Date(dtServicesTime) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
            
            '更新当前列表状态
            If ufgData.Text(i, gstrSpeExamWork_当前状态) = "已接受" Then
                ufgData.Text(i, gstrSpeExamWork_当前状态) = "已完成"
            End If
            
            '更新调用界面列表状态
            If Val(curPatholAdviceID) = mlngPatholAdivceId Then
                blnUpdateCallWind = True
            End If
        End If
    Next i
    
    If blnUpdateCallWind And Not (mufgGrid Is Nothing) Then
        For i = 1 To mufgGrid.GridRows - 1
            If mufgGrid.Text(i, gstrSpeExam_当前状态) = "已接受" Then
                Call mufgGrid.SyncText(i, gstrSpeExam_当前状态, "已完成", True)
                Call mufgGrid.SyncText(i, gstrSpeExam_特检医师, UserInfo.姓名, True)
            End If
        Next i
    End If
End Sub



'Private Sub cbxSpeExamType_Click()
'On Error GoTo errHandle
'    Dim blnIsAllowFilter As Boolean
'
'    blnIsAllowFilter = False
'    Select Case Val(tabFilter.Selected.Tag)
'        Case 0
'            blnIsAllowFilter = CheckPopedom(mstrPrivs, "免疫组化")
'
'        Case 1
'            blnIsAllowFilter = CheckPopedom(mstrPrivs, "特殊染色")
'
'        Case 2
'            blnIsAllowFilter = CheckPopedom(mstrPrivs, "分子病理")
'
'    End Select
'
'    cmdFilter.Enabled = blnIsAllowFilter
'
'    If Not blnIsAllowFilter Then
'        Call MsgBoxD(Me, "不具备查询该特检类型数据的权限。", vbOKOnly, Me.Caption)
'    End If
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub




Private Function CheckAllowSureOrAccept(Optional ByVal blnIsSure As Boolean = True) As Boolean
'判断是否需要进行核收
    Dim i As Long
    
    CheckAllowSureOrAccept = False
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetRowCheck(i) = True And (ufgData.Text(i, gstrSlices_当前状态) = IIf(blnIsSure, "已接受", "已申请")) Then
            CheckAllowSureOrAccept = True
            Exit Function
        End If
    Next i
End Function


Private Sub Menu_Edit_Accept()
'特检接受
On Error GoTo errHandle
    If Not CheckAllowSureOrAccept(False) Then
        Call MsgBoxD(Me, "没有需要进行接受的特检项目。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    Call SpeExamBatAccept
    
    blnIsOk = True
    
    Call MsgBoxD(Me, "已完成对选中检查的接受处理。", vbOKOnly, Me.Caption)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Edit_Complate()
'特检确认
On Error GoTo errHandle
    If Not CheckAllowSureOrAccept(True) Then
        Call MsgBoxD(Me, "没有需要进行完成的特检项目。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    Call SpeExamBatSure
    
    blnIsOk = True
    
    Call MsgBoxD(Me, "已完成对所选检查的完成处理。", vbOKOnly, Me.Caption)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdFilter_Click()
On Error GoTo errHandle
    Call GetSpeExamData
    
    Call RefreshSilcesCount
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub PrintSpeExamLabel(ByVal cbrControl As CommandBarControl)
'打印特检项目标签
    Dim i As Long
    Dim j As Long
    Dim strValue(5) As String
    Dim bytStyle As Byte
    
    j = 0
    strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            If zlCommFun.ActualLen(strValue(j)) > 2000 Then
                j = j + 1
                strValue(j) = ""
            End If

            If strValue(j) <> "" Then strValue(j) = strValue(j) & ","

            strValue(j) = strValue(j) & ufgData.KeyValue(i)
        End If
    Next i
    
    If cbrControl.ID = TMenuType.mtLabView Then
        bytStyle = 1
    Else
        bytStyle = 2
    End If
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_11", Me, "项目ID1=" & strValue(0), "项目ID2=" & strValue(1), "项目ID3=" & strValue(2), "项目ID4=" & strValue(3), "项目ID5=" & strValue(4), "项目ID6=" & strValue(5), bytStyle)
End Sub

Private Sub Menu_File_LabView(ByVal cbrControl As CommandBarControl)
'标签预览
On Error GoTo errHandle
    Call PrintSpeExamLabel(cbrControl)
    
    blnIsOk = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_LabPrint(ByVal cbrControl As CommandBarControl)
'标签打印
On Error GoTo errHandle
    Call PrintSpeExamLabel(cbrControl)
    
    blnIsOk = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub PrintWorkList(ByVal cbrControl As CommandBarControl)
'打印特检工作列表
    Dim i As Long
    Dim j As Long
    Dim strValue(5) As String
    Dim bytStyle As Byte
    
    j = 0
    strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            If zlCommFun.ActualLen(strValue(j)) > 2000 Then
                j = j + 1
                strValue(j) = ""
            End If

            If strValue(j) <> "" Then strValue(j) = strValue(j) & ","

            strValue(j) = strValue(j) & ufgData.KeyValue(i)
        End If
    Next i
    
    If cbrControl.ID = TMenuType.mtWorkView Then
        bytStyle = 1
    Else
        bytStyle = 2
    End If
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_10", Me, "项目ID=" & strValue(0), "项目ID1=" & strValue(1), "项目ID2=" & strValue(2), "项目ID3=" & strValue(3), "项目ID4=" & strValue(4), "项目ID5=" & strValue(5), bytStyle)
    
End Sub

Private Sub Menu_File_WorkView(ByVal cbrControl As CommandBarControl)
'预览特检工作清单
On Error GoTo errHandle
    
    Call PrintWorkList(cbrControl)
    
    blnIsOk = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_WorkPrint(ByVal cbrControl As CommandBarControl)
'打印特检工作清单
On Error GoTo errHandle
    
    Call PrintWorkList(cbrControl)
    
    blnIsOk = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Initialize()
    Set zlReport = New zl9Report.clsReport
    mblnAutoAcceptOfAfterPrint = False
End Sub


Private Sub InitFilterPage()
    With tabFilter
        .RemoveAll
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .RemoveAll
        



        .InsertItem 0, "免疫组化", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "0-免疫组化"
'        .Item(tabFilter.ItemCount - 1).Visible = true
'        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        .InsertItem 1, "特殊染色", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "1-特殊染色"
        
        .InsertItem 2, "分子病理", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "2-分子病理"

    End With
    
    tabFilter.Item(mlngFilterTabIndex).Selected = True
End Sub


Private Sub LoadFilterParameter()
    mlngFilterTabIndex = Val(zlDatabase.GetPara("特检批量过滤页面", glngSys, glngModul, 0))
    chkYSQ.value = Val(zlDatabase.GetPara("特检批量已申请", glngSys, glngModul, 1))
    chkYJS.value = Val(zlDatabase.GetPara("特检批量已接受", glngSys, glngModul, 0))
    chkYWC.value = Val(zlDatabase.GetPara("特检批量已完成", glngSys, glngModul, 0))
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Dim curDate As Date
    
    Call InitCommandBars
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call LoadFilterParameter
    
    Call InitFilterPage
    
    Call InitSpeExamWorkList
    
    curDate = zlDatabase.Currentdate
    dtpStartRequisition.value = Format(curDate, "yyyy-mm-dd 00:00")
    dtpEndRequisition.value = Format(curDate, "yyyy-mm-dd 23:59")
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    
    '设置菜单栏和工具栏风格
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True                                '显示按钮提示
        .AlwaysShowFullMenus = False                            '不常用的菜单项先隐藏
        .UseFadedIcons = False                                  '图标显示为褪色效果
        .IconsWithShadow = True                                 '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True                                '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True                                      '工具栏显示为大图标
        .SetIconSize True, 24, 24                               '设置大图标的尺寸
        .SetIconSize False, 16, 16                              '设置小图标的尺寸
    End With
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                       '设置控件显示风格
        .EnableCustomization False                              '是否允许自定义设置
        Set .Icons = zlCommFun.GetPubIcons                      '设置关联的图标控件
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '菜单定义
'Begin------------------------编辑菜单--------------------------------------默认可见
    cbrMain.ActiveMenuBar.Title = "菜单"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, TMenuType.mtLab, "标签"): cbrControl.IconId = 9023
            With cbrControl.CommandBar '二级菜单
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtLabView, "预览(0)"): cbrPopControl.IconId = 102
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtLabPrint, "打印(1)"): cbrPopControl.IconId = 103
            End With
        Set cbrControl = .Add(xtpControlPopup, TMenuType.mtWork, "清单"): cbrControl.IconId = 3031
            With cbrControl.CommandBar '二级菜单
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtWorkView, "预览(0)"): cbrPopControl.IconId = 102
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtWorkPrint, "打印(1)"): cbrPopControl.IconId = 103
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&Q)")
        cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAccept, "特价接受(&R)"): cbrControl.IconId = 747
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtComplete, "特检完成(&S)"): cbrControl.IconId = 3200
    End With
    
    'Begin----------------------查看菜单--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(V)")
    With cbrMenuBar.CommandBar
        Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(T)")
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar '二级菜单
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(0)"): cbrPopControl.Checked = True
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(1)"): cbrPopControl.Checked = True
            End With
        Set cbrControl = .Controls.Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(S)"): cbrControl.Checked = True
    End With

    'Begin----------------------帮助菜单--------------------------------------默认可见
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(H)")
    With cbrMenuBar.CommandBar
        Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_Help, "帮助主题(M)")
        Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB上的中联(W)")
            With cbrControl.CommandBar
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(0)")
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Home, "中联主页(1)")
                Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(2)")
            End With
        Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_About, "关于…(A)")
    End With
    '---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlPopup, TMenuType.mtLab, "标签"): cbrControl.IconId = 9023
            With cbrControl.CommandBar '二级菜单
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtLabView, "预览(0)"): cbrPopControl.IconId = 102
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtLabPrint, "打印(1)"): cbrPopControl.IconId = 103
            End With
        Set cbrControl = .Add(xtpControlPopup, TMenuType.mtWork, "清单"): cbrControl.IconId = 3031
            With cbrControl.CommandBar '二级菜单
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtWorkView, "预览(0)"): cbrPopControl.IconId = 102
                Set cbrPopControl = .Controls.Add(xtpControlButton, TMenuType.mtWorkPrint, "打印(1)"): cbrPopControl.IconId = 103
            End With
            
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAccept, "特检接受"): cbrControl.IconId = 747
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtComplete, "特检完成"): cbrControl.IconId = 3200
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub SaveFilterParameter()
    Call zlDatabase.SetPara("特检批量过滤页面", tabFilter.Selected.Index, glngSys, glngModul)
    Call zlDatabase.SetPara("特检批量已申请", chkYSQ.value, glngSys, glngModul)
    Call zlDatabase.SetPara("特检批量已接受", chkYJS.value, glngSys, glngModul)
    Call zlDatabase.SetPara("特检批量已完成", chkYWC.value, glngSys, glngModul)
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    Call SaveFilterParameter
    
    Set zlReport = Nothing
End Sub




Private Sub UpdateWorkListPrintState()
'在打印后，更新工作清单的打印状态
    Dim strSql As String
    Dim i As Long
    Dim strPrintIds As String
        
    strPrintIds = ""
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            strPrintIds = strPrintIds & "," & ufgData.KeyValue(i)

            strSql = "Zl_病理特检_清单打印(" & ufgData.KeyValue(i) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

            ufgData.Text(i, gstrSpeExamWork_清单状态) = "已打印"
        End If
    Next i

    '更新当前检查的特检状态
    If Trim(strPrintIds) <> "" And Not (mufgGrid Is Nothing) Then
        strPrintIds = strPrintIds & ","

        For i = 1 To mufgGrid.GridRows - 1
            If UCase(strPrintIds) Like "*," & UCase(mufgGrid.KeyValue(i)) & ",*" Then

                Call mufgGrid.SyncText(i, gstrSpeExam_清单状态, "已打印", True)
            End If
        Next i
    End If
End Sub






Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle

   If ufgData.DataGrid.Rows > 1 Then Call GetSpeExamData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub zlReport_AfterPrint(ByVal ReportNum As String)
'清单已打印
On Error GoTo errHandle
    Call UpdateWorkListPrintState
    
    If mblnAutoAcceptOfAfterPrint Then
        Call SpeExamBatAccept
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
