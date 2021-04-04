VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPatholSlices_WorkPrint 
   Caption         =   "制片批量处理"
   ClientHeight    =   7470
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   12240
   Icon            =   "frmPatholSlices_WorkPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   12240
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picTag 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11520
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   600
      ScaleHeight     =   6735
      ScaleWidth      =   11175
      TabIndex        =   0
      Top             =   360
      Width           =   11175
      Begin VB.Frame framFilter 
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   11055
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
            Left            =   8160
            TabIndex        =   9
            Top             =   240
            Width           =   1455
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
            Left            =   6360
            TabIndex        =   8
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
            Left            =   9720
            TabIndex        =   7
            Top             =   200
            Width           =   1215
         End
         Begin VB.OptionButton optMaterialTime 
            Caption         =   "取材"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   165
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optRequisitionTime 
            Caption         =   "申请"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   390
            Width           =   735
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   3720
            TabIndex        =   10
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   99155971
            CurrentDate     =   40679.726087963
         End
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   300
            Left            =   1560
            TabIndex        =   11
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   99155971
            CurrentDate     =   40679.0594097222
         End
         Begin VB.Label Label3 
            Caption         =   "到"
            Height          =   255
            Left            =   7920
            TabIndex        =   15
            Top             =   285
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "按         时间："
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   280
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "到"
            Height          =   255
            Left            =   3480
            TabIndex        =   13
            Top             =   285
            Width           =   255
         End
         Begin VB.Label labPatholNum 
            Caption         =   "病理号："
            Height          =   255
            Left            =   5640
            TabIndex        =   12
            Top             =   285
            Width           =   720
         End
      End
      Begin VB.CheckBox chkYWC 
         Caption         =   "已完成"
         Height          =   180
         Left            =   10080
         TabIndex        =   3
         ToolTipText     =   "显示制片状态为“未完成”的制片记录。"
         Top             =   5910
         Width           =   855
      End
      Begin VB.CheckBox chkYJS 
         Caption         =   "已接受"
         Height          =   180
         Left            =   9120
         TabIndex        =   2
         ToolTipText     =   "显示制片状态为“已接受”的制片记录。"
         Top             =   5910
         Width           =   855
      End
      Begin VB.CheckBox chkWCL 
         Caption         =   "未处理"
         Height          =   255
         Left            =   8160
         TabIndex        =   1
         ToolTipText     =   "显示制片状态为“未处理”的制片记录。"
         Top             =   5880
         Width           =   855
      End
      Begin XtremeSuiteControls.TabControl tabFilter 
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   840
         Width           =   11085
         _Version        =   589884
         _ExtentX        =   19553
         _ExtentY        =   661
         _StockProps     =   64
      End
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   4215
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7435
         DefaultCols     =   ""
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   7110
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatholSlices_WorkPrint.frx":179A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14711
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
Attribute VB_Name = "frmPatholSlices_WorkPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1

Private mlngPatholAdviceId As Long
Private mufgParGrid As ucFlexGrid

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
    mtchkWCL = 5
    mtchkYJS = 6
    mtchkYWC = 7
End Enum

Private Sub RefreshSilcesCount()
'刷新制片记录数量
    Dim i As Long
    Dim lngRecord As Long
    Dim lngTotal As Long
    Dim lngSlices As Long
    
    On Error GoTo errH
    
    lngTotal = 0
    lngSlices = 0
    
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.RowHidden(i) Then
            If Not ufgData.IsNullRow(i) Then

                lngTotal = lngTotal + Val(ufgData.Text(i, gstrSlices_制片数))

                If ufgData.Text(i, gstrSlices_当前状态) <> "已完成" Then
                    lngSlices = lngSlices + Val(ufgData.Text(i, gstrSlices_制片数))
                End If
            End If
        End If
    Next i
    
    stbThis.Panels(2).Text = "制片总数：" & lngTotal & "    需制片数：" & lngSlices
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, Me.Caption)
End Sub



Public Sub ShowWorkPrint(ufgGrid As ucFlexGrid, ByVal lngPatholAdviceId As Long, owner As Form)
'显示工作清单打印窗口
    Set mufgParGrid = ufgGrid

    mlngPatholAdviceId = lngPatholAdviceId
    blnIsOk = False
        
'    '载入当前检查制片数据
'    If Trim(lngPatholAdviceId) > 0 Then
'        Call LoadSpecifySlicesData
'    End If
    
    Call RefreshSilcesCount
    
    Call Me.Show(1, owner)
    
End Sub


Private Sub GetSlicesData()
    Dim strSql As String
    Dim strPatholNumQuery As String


    strPatholNumQuery = ""
    If Trim(txtStartPatholNum.Text) <> "" And Trim(txtEndPatholNum.Text) <> "" Then
        strPatholNumQuery = " and (REGEXP_SUBSTR(upper(c.病理号), '[[:alpha:]]+') >=REGEXP_SUBSTR(upper([3]),'[[:alpha:]]+') and to_number(REGEXP_SUBSTR(upper(c.病理号), '[[:digit:]]+')) >=to_number(REGEXP_SUBSTR(upper([3]),  '[[:digit:]]+'))) "
        strPatholNumQuery = strPatholNumQuery & " and  (REGEXP_SUBSTR(upper(c.病理号), '[[:alpha:]]+') <=REGEXP_SUBSTR(upper([4]),'[[:alpha:]]+') and to_number(REGEXP_SUBSTR(upper(c.病理号),  '[[:digit:]]+')) <=to_number(REGEXP_SUBSTR(upper([4]), '[[:digit:]]+'))) "
    ElseIf Trim(txtStartPatholNum.Text) <> "" Then
        strPatholNumQuery = " and upper(c.病理号)=upper([3]) "
    ElseIf Trim(txtStartPatholNum.Text) <> "" Then
        strPatholNumQuery = " and upper(c.病理号) =upper([4]) "
    End If
    
    
    strSql = "select a.Id,c.病理号,c.病理医嘱ID, e.姓名, g.名称 as 号别名称, a.材块ID,b.序号,b.取材位置, d.标本名称, d.标本类型, a.制片类型, a.制片方式,a.制片数,b.取材时间, a.当前状态,a.清单状态 " & _
                " from 病理制片信息 a, 病理取材信息 b, 病理检查信息 c, 病理标本信息 d, 病人医嘱记录 e ,病理号码规则 g" & _
                IIf(optRequisitionTime.value, ",病理申请信息 f ", "") & _
                " Where a.材块id = b.材块id And b.病理医嘱ID = c.病理医嘱ID and b.确认状态=1 And c.医嘱ID = e.ID And b.标本id = d.标本id and c.号码规则ID=g.ID" & _
                IIf(optMaterialTime.value, " and b.取材时间 between [1] and [2]", "") & _
                IIf(optRequisitionTime.value, " and a.申请ID=f.申请ID and f.申请时间 between [1] and [2]", "") & _
                IIf(strPatholNumQuery <> "", strPatholNumQuery, "") & _
                " order by c.病理号,a.当前状态,b.序号,a.Id "
    'If mblnMoved Then strSql = GetMovedDataSql(strSql)
                

    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                            CDate(dtpStart.value), _
                                            CDate(dtpEnd.value), _
                                            txtStartPatholNum.Text, _
                                            txtEndPatholNum.Text)
                                            
                                                                    
    Call FilterSlicesData
End Sub



Private Sub FilterSlicesData()
'过滤查询出来的制片数据
    Dim strFilters As String
    Dim strStudyTypeFilter As String
    
    strFilters = ""
    
    
    strStudyTypeFilter = ""
    Select Case tabFilter.Selected.tag
        Case "所有"
            strStudyTypeFilter = ""
        Case Else
            strStudyTypeFilter = "号别名称=" & "'" & tabFilter.Selected.tag & "'"
    End Select
    
        
    If chkWCL.value <> 0 Then
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
    If chkWCL.value = 0 And chkYJS.value = 0 And chkYWC.value = 0 Then
        strFilters = "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " 当前状态=0)" & " or " & _
                     "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " 当前状态=1)" & " or " & _
                     "(" & IIf(strStudyTypeFilter = "", "", strStudyTypeFilter & " and ") & " 当前状态=2)"
    End If
    
    If ufgData.AdoData Is Nothing Then Exit Sub
    
    ufgData.AdoData.Filter = strFilters
    ufgData.RefreshData
    
    Call RefreshSilcesCount
End Sub



Private Sub LoadSpecifySlicesData()
'载入指定的病理制片数据
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select a.ID, a.Id,c.病理号,c.病理医嘱ID, f.名称 as 号别名称 ,e.姓名, c.号码规则ID, a.材块ID,b.序号,b.取材位置, d.标本名称, d.标本类型, a.制片类型, a.制片方式, a.制片数,a.当前状态,a.清单状态 " & _
                 " from 病理制片信息 a, 病理取材信息 b, 病理检查信息 c, 病理标本信息 d, 病人医嘱记录 e,病理号码规则 f " & _
                " Where a.材块id = b.材块id And b.病理医嘱ID = c.病理医嘱ID And c.医嘱ID = e.ID And b.标本id = d.标本id " & _
                " and c.病理医嘱ID=[1] and a.当前状态 <> 2 and c.号码规则ID=f.ID" & _
                " order by 号别名称,病理号,材块ID,当前状态 "
    'If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngPatholAdviceId)
    
    Call ufgData.RefreshData
End Sub


Private Sub InitSlicesWorkList()
'初始化制片清单显示列
    Dim strTemp As String
    
    '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
	ufgData.DefaultColNames = gstrSlicesWorkCols
    
    '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("批量制片列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
     
    If strTemp = "" Then
        '初始化标本显示列表
        ufgData.ColNames = gstrSlicesWorkCols
    Else
        ufgData.ColNames = strTemp
    End If
    
    
    ufgData.ColConvertFormat = gstrSlicesWorkConvertFormat
    ufgData.IsShowPopupMenu = False
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Select Case control.ID
        Case TMenuType.mtLabView                    '标签预览
            Call Menu_File_LabView(control)
        
        Case TMenuType.mtLabPrint                   '标签打印
            Call Menu_File_LabPrint(control)
            
        Case TMenuType.mtWorkView                   '清单预览
            Call Menu_File_WorkView(control)
        
        Case TMenuType.mtWorkPrint                  '清单打印
            Call Menu_File_WorkPrint(control)
        
        Case TMenuType.mtAccept                     '制片接受
            Call Menu_Edit_Accept
        
        Case TMenuType.mtComplete                   '制片完成
            Call Menu_Edit_Complete

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
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_Exit()
    blnIsOk = False
    Me.Hide
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
    picMain.Left = Left
    picMain.Top = Top
    picMain.Width = Right - Left
    picMain.Height = Bottom - Top
End Sub

Private Sub picMain_Resize()
On Error Resume Next
    framFilter.Left = 120
    framFilter.Top = 0
    framFilter.Width = picMain.Width - 240
    
    tabFilter.Left = 120
    tabFilter.Top = framFilter.Top + framFilter.Height
    tabFilter.Width = picMain.Width - chkWCL.Width * 5 - 720
    
    chkWCL.Left = tabFilter.Width
    chkWCL.Top = tabFilter.Top + 40
    
    chkYJS.Left = chkWCL.Left + chkWCL.Width + 240
    chkYJS.Top = chkWCL.Top + 40
    
    chkYWC.Left = chkYJS.Left + chkYJS.Width + 240
    chkYWC.Top = chkYJS.Top
    
    ufgData.Left = 120
    ufgData.Top = tabFilter.Top + tabFilter.Height
    ufgData.Width = picMain.Width - 240
    ufgData.Height = picMain.Height - framFilter.Height - tabFilter.Height
End Sub

Private Sub ufgData_OnColFormartChange()
'保存列表配置
    zlDatabase.SetPara "批量制片列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub

Private Sub SlicesBatAccept()
'特检批量接受
    Dim i As Long
    Dim curPatholAdviceID As String
    Dim strSql As String
    Dim blnUpdateCallWind As Boolean
    
    blnUpdateCallWind = False
    
    For i = 1 To ufgData.GridRows - 1
        '如果选中的检查，才进行接收
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            If curPatholAdviceID <> ufgData.Text(i, gstrSlicesWork_病理医嘱ID) Then
                curPatholAdviceID = ufgData.Text(i, gstrSlicesWork_病理医嘱ID)
                
                strSql = "Zl_病理制片_接受(" & Val(curPatholAdviceID) & ",'" & UserInfo.姓名 & "')"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
            
            '更新当前列表状态
            If ufgData.Text(i, gstrSlicesWork_当前状态) = "未处理" Then
                ufgData.Text(i, gstrSlicesWork_当前状态) = "已接受"
            End If
            
            '更新调用界面列表状态
            If Val(curPatholAdviceID) = mlngPatholAdviceId Then
                blnUpdateCallWind = True
            End If
        End If
    Next i
    
    If blnUpdateCallWind And Not (mufgParGrid Is Nothing) Then
        For i = 1 To mufgParGrid.GridRows - 1
            If mufgParGrid.Text(i, gstrSlicesWork_当前状态) = "未处理" Then
                mufgParGrid.Text(i, gstrSlices_当前状态) = "已接受"
                mufgParGrid.Text(i, gstrSlices_制片人) = UserInfo.姓名
            End If
        Next i
    End If
End Sub




Private Sub SlicesBatSure()
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
            If curPatholAdviceID <> ufgData.Text(i, gstrSlicesWork_病理医嘱ID) Then
                curPatholAdviceID = ufgData.Text(i, gstrSlicesWork_病理医嘱ID)
                
                strSql = "Zl_病理制片_确认(" & Val(curPatholAdviceID) & "," & zlStr.To_Date(dtServicesTime) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
            
            '更新当前列表状态
            If ufgData.Text(i, gstrSlicesWork_当前状态) = "已接受" Then
                ufgData.Text(i, gstrSlicesWork_当前状态) = "已完成"
            End If
            
            '更新调用界面列表状态
            If Val(curPatholAdviceID) = mlngPatholAdviceId Then
                blnUpdateCallWind = True
            End If
        End If
    Next i
    
    If blnUpdateCallWind And Not (mufgParGrid Is Nothing) Then
        For i = 1 To mufgParGrid.GridRows - 1
            If mufgParGrid.Text(i, gstrSlicesWork_当前状态) = "已接受" Then
                mufgParGrid.Text(i, gstrSlices_当前状态) = "已完成"
                mufgParGrid.Text(i, gstrSlices_制片人) = UserInfo.姓名
            End If
        Next i
    End If
End Sub



Private Function CheckAllowSureOrAccept(Optional ByVal blnIsSure As Boolean = True) As Boolean
'判断是否需要进行核收
    Dim i As Long
    
    CheckAllowSureOrAccept = False
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetRowCheck(i) = True And (ufgData.Text(i, gstrSlices_当前状态) = IIf(blnIsSure, "已接受", "未处理")) Then
            CheckAllowSureOrAccept = True
            Exit Function
        End If
    Next i
End Function


Private Sub chkWCL_Click()
On Error GoTo ErrHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSlicesData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkYJS_Click()
On Error GoTo ErrHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSlicesData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkYWC_Click()
On Error GoTo ErrHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSlicesData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_Accept()
'制片接受
On Error GoTo ErrHandle
    If Not CheckAllowSureOrAccept(False) Then
        Call MsgBoxD(Me, "尚无需要接受的制片信息。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call SlicesBatAccept
    
    blnIsOk = True
    
    Call MsgBoxD(Me, "已完成对所选检查的接受处理。", vbOKOnly, Me.Caption)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Edit_Complete()
'制片完成
On Error GoTo ErrHandle
    If Not CheckAllowSureOrAccept(True) Then
        Call MsgBoxD(Me, "尚无需要完成的制片信息。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call SlicesBatSure
    
    blnIsOk = True
    
    Call MsgBoxD(Me, "已完成对所选检查的制片处理。", vbOKOnly, Me.Caption)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdFilter_Click()
On Error GoTo ErrHandle
    Call GetSlicesData
    
    Call RefreshSilcesCount
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub PrintSlicesLabel(ByVal cbrControl As CommandBarControl)
'打印预览特检项目标签
    Dim i As Long
    Dim j As Long
    Dim strValue(5) As String
    
    Dim strSliceId As String
    Dim k As Long
    Dim lngCount As Long
    Dim bytStyle As Byte
    
    j = 0
    strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0"
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            If zlCommFun.ActualLen(strValue(j)) > 2000 Then
                j = j + 1
                strValue(j) = ""
            End If

            strSliceId = ufgData.KeyValue(i)
            lngCount = Val(ufgData.Text(i, gstrSlices_制片数))
    
            If strValue(j) <> "" Then strValue(j) = strValue(j) & ","
    
            strValue(j) = strValue(j) & strSliceId
            
            If lngCount > 1 Then
                For k = 1 To lngCount - 1
                    strValue(j) = strValue(j) & "," & strSliceId
                Next k
            End If
        End If
    Next i
    
    If cbrControl.ID = TMenuType.mtLabView Then
        bytStyle = 1
    Else
        bytStyle = 2
    End If
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_09", Me, "制片ID1=" & strValue(0), "制片ID2=" & strValue(1), "制片ID3=" & strValue(2), "制片ID4=" & strValue(3), "制片ID5=" & strValue(4), "制片ID6=" & strValue(5), bytStyle)
End Sub

Private Sub Menu_File_LabView(ByVal cbrControl As CommandBarControl)
'标签预览
On Error GoTo ErrHandle
    Call PrintSlicesLabel(cbrControl)
    
    blnIsOk = True
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_LabPrint(ByVal cbrControl As CommandBarControl)
'标签打印
On Error GoTo ErrHandle
    Call PrintSlicesLabel(cbrControl)
    
    blnIsOk = True
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub




Private Sub PrintWorkList(ByVal cbrControl As CommandBarControl)
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
    
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_08", Me, "制片ID1=" & strValue(0), "制片ID2=" & strValue(1), "制片ID3=" & strValue(2), "制片ID4=" & strValue(3), "制片ID5=" & strValue(4), "制片ID6=" & strValue(5), bytStyle)
    
End Sub

Private Sub Menu_File_WorkView(ByVal cbrControl As CommandBarControl)
On Error GoTo ErrHandle
    
    Call PrintWorkList(cbrControl)
    
    blnIsOk = True
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_WorkPrint(ByVal cbrControl As CommandBarControl)
On Error GoTo ErrHandle
    
    Call PrintWorkList(cbrControl)
    
    blnIsOk = True
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Initialize()
    Set zlReport = New zl9Report.clsReport
    
    mblnAutoAcceptOfAfterPrint = False
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
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAccept, "制片接受(&R)"): cbrControl.IconId = 747
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtComplete, "制片完成(&S)"): cbrControl.IconId = 3200
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
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtAccept, "制片接受"): cbrControl.IconId = 747
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtComplete, "制片完成"): cbrControl.IconId = 3200
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub InitFilterPage()
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim i As Long
    
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
        
        strSql = "select ID,名称 from 病理号码规则"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "获得病理名称")

        If rsData.RecordCount > 0 Then
            
            rsData.MoveFirst
        
            For i = 0 To rsData.RecordCount - 1
                If NVL(rsData!名称, "  ") <> "  " Then
                    .InsertItem i, rsData!名称, picTag.hWnd, 0
                    .Item(tabFilter.ItemCount - 1).tag = rsData!名称
                End If
                rsData.MoveNext
            Next
            
            .InsertItem rsData.RecordCount, "所  有", picTag.hWnd, 0
            .Item(tabFilter.ItemCount - 1).tag = "所有"
            
        End If
        
    End With
    
    tabFilter.Item(mlngFilterTabIndex).Selected = True
End Sub


Private Sub LoadFilterParameter()
    mlngFilterTabIndex = Val(zlDatabase.GetPara("制片批量过滤页面", glngSys, glngModul, 0))
    chkWCL.value = Val(zlDatabase.GetPara("制片批量未处理", glngSys, glngModul, 1))
    chkYJS.value = Val(zlDatabase.GetPara("制片批量已接受", glngSys, glngModul, 0))
    chkYWC.value = Val(zlDatabase.GetPara("制片批量已完成", glngSys, glngModul, 0))
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Dim curDate As Date
    
    Call InitCommandBars
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call LoadFilterParameter
    
    Call InitFilterPage
    
    '初始化数据列表
    Call InitSlicesWorkList
    
    curDate = zlDatabase.Currentdate
    
    dtpStart.value = Format(curDate - 1, "yyyy-mm-dd 00:00")
    dtpEnd.value = Format(curDate - 1, "yyyy-mm-dd 23:59")
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SaveFilterParameter()
    Call zlDatabase.SetPara("制片批量过滤页面", tabFilter.Selected.Index, glngSys, glngModul)
    Call zlDatabase.SetPara("制片批量未处理", chkWCL.value, glngSys, glngModul)
    Call zlDatabase.SetPara("制片批量已接受", chkYJS.value, glngSys, glngModul)
    Call zlDatabase.SetPara("制片批量已完成", chkYWC.value, glngSys, glngModul)
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    Call SaveFilterParameter
    
    Set zlReport = Nothing
End Sub

Private Sub UpdateSlicesPrintState()
'在打印后，接受打印过的制片数据
    Dim strSql As String
    Dim i As Long
    Dim strPrintIds As String
        
    strPrintIds = ""
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, ufgData.GetColIndexWithRowCheck()) Then
            strPrintIds = strPrintIds & "," & ufgData.KeyValue(i)
            
            strSql = "Zl_病理制片_清单打印(" & ufgData.KeyValue(i) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            
            ufgData.Text(i, gstrSlices_清单状态) = "已打印"
        End If
    Next i
    
    '更新当前检查的制片记录状态
    If Trim(strPrintIds) <> "" And Not (mufgParGrid Is Nothing) Then
        strPrintIds = strPrintIds & ","

        For i = 1 To mufgParGrid.GridRows - 1
            If UCase(strPrintIds) Like "*," & UCase(mufgParGrid.KeyValue(i)) & ",*" Then

                mufgParGrid.Text(i, gstrSpeExam_清单状态) = "已打印"
            End If
        Next i
    End If
End Sub



Private Sub tabFilter_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo ErrHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterSlicesData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnColsNameReSet()
On Error GoTo ErrHandle

    If ufgData.DataGrid.Rows > 1 Then Call GetSlicesData
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub zlReport_AfterPrint(ByVal ReportNum As String)
'清单已打印
On Error GoTo ErrHandle
    '如果不是制片清单打印，则直接退出
    If ReportNum <> "ZL1_PATHOLSLICES_01" Then Exit Sub
    
    Call UpdateSlicesPrintState
    
    '打印后自动核收
    If mblnAutoAcceptOfAfterPrint Then
        Call SlicesBatAccept
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

