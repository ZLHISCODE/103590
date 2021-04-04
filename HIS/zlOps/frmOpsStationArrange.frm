VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOpsStationArrange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "手术安排"
   ClientHeight    =   5505
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6570
   Icon            =   "frmOpsStationArrange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      Caption         =   "附加内容"
      Height          =   2055
      Left            =   30
      TabIndex        =   10
      Top             =   3375
      Width           =   5145
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   2
         Left            =   4620
         TabIndex        =   23
         Text            =   "1"
         Top             =   315
         Width           =   390
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   660
         Width           =   1845
      End
      Begin VB.ListBox lst 
         Height          =   900
         Left            =   1455
         Style           =   1  'Checkbox
         TabIndex        =   20
         Top             =   1035
         Width           =   3555
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   300
         Width           =   1845
      End
      Begin VB.CheckBox chk 
         Caption         =   "感染手术(&3)"
         Height          =   195
         Index           =   3
         Left            =   3330
         TabIndex        =   15
         Top             =   720
         Width           =   1305
      End
      Begin VB.CheckBox chk 
         Caption         =   "污染手术(&2)"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   14
         Top             =   1065
         Width           =   1305
      End
      Begin VB.CheckBox chk 
         Caption         =   "接台手术(&F)"
         Height          =   195
         Index           =   0
         Left            =   3330
         TabIndex        =   13
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "手术性质(&X)"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "紧急程度(&J)"
         Height          =   180
         Index           =   4
         Left            =   135
         TabIndex        =   11
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "手术人员(&R)"
      Height          =   2370
      Left            =   30
      TabIndex        =   8
      Top             =   945
      Width           =   5145
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1995
         Left            =   105
         TabIndex        =   9
         Top             =   255
         Width           =   4950
         _cx             =   8731
         _cy             =   3519
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   5355
      TabIndex        =   18
      Top             =   1350
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5355
      TabIndex        =   16
      Top             =   60
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5355
      TabIndex        =   17
      Top             =   570
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Height          =   945
      Left            =   45
      TabIndex        =   19
      Top             =   -45
      Width           =   5115
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   525
         Width           =   3510
      End
      Begin VB.CommandButton cmd 
         Height          =   330
         Index           =   1
         Left            =   4635
         Picture         =   "frmOpsStationArrange.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "多选，快捷键：F3"
         Top             =   495
         Width           =   345
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   0
         Left            =   3945
         TabIndex        =   2
         Text            =   "1"
         Top             =   180
         Visible         =   0   'False
         Width           =   390
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Left            =   1125
         TabIndex        =   1
         Top             =   165
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   106692611
         CurrentDate     =   38083
      End
      Begin MSComCtl2.UpDown udp 
         Height          =   300
         Left            =   4350
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   165
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt(0)"
         BuddyDispid     =   196610
         BuddyIndex      =   0
         OrigLeft        =   4395
         OrigTop         =   165
         OrigRight       =   4635
         OrigBottom      =   465
         Max             =   12
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chk 
         Caption         =   "时长(&H)"
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   24
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "开始时间(&T)"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "手 术 间(&M)"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   5
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "小时"
         Height          =   180
         Left            =   4650
         TabIndex        =   4
         Top             =   210
         Visible         =   0   'False
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmOpsStationArrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
'（１）窗体级变量定义

Private mblnReading As Boolean
Private mblnDataChanged As Boolean
Private mblnOK As Boolean
Private mlngKey As Long
Private mlngDeptKey As Long
Private mfrmMain As Form
Private mstrPrivs As String
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

'######################################################################################################################

Public Function ShowEdit(ByVal frmMain As Form, Optional lngKey As Long = 0, Optional lngDeptKey As Long = 0, Optional ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '功能：打开编辑窗体进行数据的新增、修改操作
    '参数：
    '返回：
    '******************************************************************************************************************
    mlngKey = lngKey
    mlngDeptKey = lngDeptKey
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    If ExecuteCommand("初始控件") = False Then Exit Function
    If ExecuteCommand("初始数据") = False Then Exit Function
    
    Call ExecuteCommand("读取数据")
    
    DataChanged = False
    
    Me.Show 1, mfrmMain
    
    ShowEdit = mblnOK
    
End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：对新增、修改的数据进行合法性校验
    '返回：校验合法返回True，否则返回False
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset

    If chk(1).Value = 1 Then
        If Val(txt(0).Text) < 1 Or Val(txt(0).Text) > 12 Then
            ShowSimpleMsg "手术时长必须大于1小时而小于12小时！"
            
            zlControl.TxtSelAll txt(0)
            txt(0).SetFocus
            Exit Function
        End If
    End If
    
    gstrSQL = "SELECT 1 FROM 医技执行房间 WHERE 执行间=[1] AND 科室id=[2]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txt(1).Text, mlngDeptKey)
    
    If rs.BOF Then
        ShowSimpleMsg "安排手术间了一个不存在的手术间！"
        zlControl.TxtSelAll txt(1)
        txt(1).SetFocus
        Exit Function
    End If
    
    '检查手术时间与申请时间的关系
    gstrSQL = "SELECT 开嘱时间 FROM 病人医嘱记录 WHERE ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
    If rs.BOF = False Then
        If Format(dtp.Value, "yyyy-MM-dd HH:mm") < Format(rs("开嘱时间").Value, "yyyy-MM-dd HH:mm") Then
            
            If MsgBox("手术开始时间(" & Format(dtp.Value, "yyyy-MM-dd HH:mm") & ")早于申请时间(" & Format(rs("开嘱时间").Value, "yyyy-MM-dd HH:mm") & ")" & vbCrLf & "是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                dtp.SetFocus
                Exit Function
            End If
            
        End If
    End If
    
    '检查一个病人是否在同一时间段内做两种手术
    gstrSQL = "SELECT 1 FROM 病人手术记录 B " & _
                "WHERE B.手术状态 In (2,3) AND  " & _
                       "B.医嘱id <> [3] AND  " & _
                       "(B.病人id, NVL(B.主页id,0)) IN (SELECT 病人id, NVL(主页id,0) FROM 病人医嘱记录 WHERE ID = [3]) AND  " & _
                       "((B.手术开始时间 BETWEEN [1] AND [2]) OR  " & _
                       "(B.手术结束时间 BETWEEN [1] AND [2]))"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(DateAdd("h", Val(txt(0).Text), dtp.Value), "YYYY-MM-DD HH:MM:SS")), mlngKey)
    If rs.BOF = False Then
        ShowSimpleMsg "当前病人不能同时进行二场手术。"
        dtp.SetFocus
        Exit Function
    End If
    
    ValidData = True
    
End Function

Private Function SaveData() As Boolean
    '******************************************************************************************************************
    '功能：对新增、修改后的数据进行保存/更新处理
    '参数：返回参lngKey，表示更新记录的关键字
    '返回：保存成功返回True，否则返回False
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim rsSQL As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim str污染手术 As String
    Dim blnTrans As Boolean
    
    On Error GoTo errHand
    
    Call SQLRecord(rsSQL)
    
    With vsf
        For lngLoop = 1 To .Rows - 1
            If .TextMatrix(lngLoop, .ColIndex("岗位")) <> "" And .TextMatrix(lngLoop, .ColIndex("姓名")) <> "" Then
                strTmp = strTmp & ";" & Val(.RowData(lngLoop)) & "," & .TextMatrix(lngLoop, .ColIndex("岗位")) & "," & .TextMatrix(lngLoop, .ColIndex("姓名")) & "," & .TextMatrix(lngLoop, .ColIndex("编号"))
            End If
        Next
    End With
    
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    
    For lngLoop = 0 To lst.ListCount - 1
        If lst.Selected(lngLoop) Then
            str污染手术 = str污染手术 & ";" & lst.List(lngLoop)
        End If
    Next
    If str污染手术 <> "" Then str污染手术 = Mid(str污染手术, 2)
    
    gstrSQL = "zl_病人手术记录_Arrange(" & mlngKey & "," & _
                                        "To_Date('" & Format(dtp.Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                        IIf(chk(1).Value = 0, "Null", "TO_DATE('" & Format(DateAdd("h", Val(txt(0).Text), dtp.Value), "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')") & ",'" & _
                                        txt(1).Text & "'," & _
                                        mlngDeptKey & ",'" & _
                                        strTmp & "',2,'" & cbo(0).Text & "'," & Val(txt(2).Text) & ",'" & zlCommFun.GetNeedName(cbo(1).Text) & "'," & chk(2).Value & ",'" & str污染手术 & "'," & chk(3).Value & ")"
    Call SQLRecordAdd(rsSQL, gstrSQL)
    
    gstrSQL = "Zl_病人手术记录_Updateadvice(" & mlngKey & ")"
    Call SQLRecordAdd(rsSQL, gstrSQL)
    
    '特殊SQL执行,由具体执行时决定
    '------------------------------------------------------------------------------------------------------------------
    If mfrmMain.NotAutoCharge = False Then
        gstrSQL = "Select b.医嘱id From 病人医嘱记录 a,病人手术记录 b Where b.ID=[1] And b.医嘱id=a.ID And a.医嘱状态 Not In (4,8)"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            Call SQLRecordAdd(rsSQL, "", 0, 1, Val(rs("医嘱id").Value))
        End If
    End If
            
    '开始执行SQL,即提交到数据库中
    '------------------------------------------------------------------------------------------------------------------
    If rsSQL.RecordCount > 0 Then
        
        rsSQL.MoveFirst
        
        blnTrans = True
        gcnOracle.BeginTrans

        For lngLoop = 1 To rsSQL.RecordCount
            
            If Val(rsSQL("Trans").Value) = 1 And blnTrans = False Then
                blnTrans = True
                gcnOracle.BeginTrans
            End If
            
            '生成体检项目的费用,
            If Val(rsSQL("Custom").Value) = 1 Then
                If CreateOrderCharge(Val(rsSQL("Parameter").Value), mstrPrivs) = False Then
                    blnTrans = False
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
            Else
                If Val(rsSQL("Trans").Value) = 2 Then
                    If blnTrans Then
                        gcnOracle.CommitTrans
                        blnTrans = False
                        SaveData = True
                    End If
                Else
                    gstrSQL = CStr(rsSQL("SQL").Value)
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
            End If
            rsSQL.MoveNext
        Next
        
        If blnTrans Then
            gcnOracle.CommitTrans
            blnTrans = False
            SaveData = True
        End If
    End If
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf, True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)

            Call .AppendColumn("岗位", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("姓名", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("编号", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .InitializeEdit(True, True, True)
            
            Call .InitializeEditColumn(.ColIndex("岗位"), True, vbVsfEditCombox)
            Call .InitializeEditColumn(.ColIndex("姓名"), True, vbVsfEditCommand)
            
            .IndicatorCol = 0
            Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
            
            .AppendRows = True
        End With
        txt(1).BackColor = COLOR.锁色

    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
        
        '手术紧急程度
        '--------------------------------------------------------------------------------------------------------------
        With cbo(0)
            .Clear
            .AddItem ""
            strSQL = "Select 编码,名称,简码,缺省标志 From 手术紧急程度"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    .AddItem rs("编码").Value & "-" & rs("名称").Value
                    If rs("缺省标志").Value = 1 Then .ListIndex = .NewIndex
                    rs.MoveNext
                Loop
            End If
            If .ListCount > 0 And .ListIndex = -1 Then .ListIndex = 0
        End With
            

        '手术性质分类
        '--------------------------------------------------------------------------------------------------------------
        With cbo(1)
            .Clear
            .AddItem ""
            strSQL = "Select 编码,名称,简码,缺省标志 From 手术性质分类"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    .AddItem rs("编码").Value & "-" & rs("名称").Value
                    If rs("缺省标志").Value = 1 Then .ListIndex = .NewIndex
                    rs.MoveNext
                Loop
            End If
            If .ListCount > 0 And .ListIndex = -1 Then .ListIndex = 0
        End With
        
        
        '手术污染分类
        '--------------------------------------------------------------------------------------------------------------
        With lst
            .Clear
            strSQL = "Select 编码,名称,简码 From 手术污染分类"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    .AddItem rs("名称").Value
                    rs.MoveNext
                Loop
            End If
        End With
        
        '手术岗位
        '--------------------------------------------------------------------------------------------------------------
        gstrSQL = "SELECT 编码||'-'||名称 As 名称 FROM 手术岗位 Order by 编码"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
        Call mclsVsf.InitializeEditColumn(mclsVsf.ColIndex("岗位"), True, vbVsfEditCombox, vsf.BuildComboList(rs, "名称", "名称"))
        
        
        dtp.Value = Format(zlDatabase.Currentdate + 1, dtp.CustomFormat)
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    
    '------------------------------------------------------------------------------------------------------------------
    Case "清空数据"
        
        mblnReading = True
        
        
        mblnReading = False
    '------------------------------------------------------------------------------------------------------------------
    Case "读取数据"
        gstrSQL = "SELECT a.紧急程度,a.接台手术,a.手术性质,a.污染手术,a.污染内容,a.感染手术,a.手术开始时间,Ceil((a.手术结束时间-a.手术开始时间)*24) As 小时,a.手术间 FROM 病人手术记录 a Where a.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey)
        If rs.BOF = False Then
            
            zlControl.CboLocate cbo(0), zlCommFun.NVL(rs("紧急程度").Value)
            txt(2).Text = zlCommFun.NVL(rs("接台手术").Value, 0)
            chk(0).Value = IIf(Val(txt(2).Text) > 0, 1, 0)
            
            Call zlControl.CboLocate(cbo(1), zlCommFun.NVL(rs("手术性质").Value))

            chk(2).Value = zlCommFun.NVL(rs("污染手术").Value, 0)
            If chk(2).Value = 1 Then
                strTmp = ";" & zlCommFun.NVL(rs("污染内容").Value) & ";"
                For intLoop = 0 To lst.ListCount - 1
                    If InStr(strTmp, ";" & lst.List(intLoop) & ";") > 0 Then
                        lst.Selected(intLoop) = True
                    End If
                Next
            End If
            
            chk(3).Value = zlCommFun.NVL(rs("感染手术").Value, 0)
            
            txt(1).Text = zlCommFun.NVL(rs("手术间").Value)
            dtp.Value = Format(zlCommFun.NVL(rs("手术开始时间").Value, zlDatabase.Currentdate), dtp.CustomFormat)
            txt(0).Text = zlCommFun.NVL(rs("手术间").Value, 1)
            
        End If
        
        
        '读取已安排的手术人员
        '--------------------------------------------------------------------------------------------------------------
        mclsVsf.ClearGrid
        gstrSQL = "Select A.人员id As ID,a.岗位,B.编号,a.姓名 From 病人手术人员 a,人员表 b,手术岗位 c Where c.名称=a.岗位 And Nvl(a.期间,1)=1 And a.记录id=[1] And a.人员id=b.ID(+) order by c.编码"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then Call mclsVsf.LoadGrid(rs)
        
    End Select

    ExecuteCommand = True

    Exit Function
    
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub chk_Click(Index As Integer)
    If Index = 1 Then
        txt(0).Visible = (chk(Index).Value = 1)
        udp.Visible = txt(0).Visible
        Label2.Visible = txt(0).Visible
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 1      '手术执行间
        
        gstrSQL = "Select RowNum As ID,执行间,Decode(b.手术间,Null,'空闲',Decode(b.手术状态,2,'预订',3,'在用')) As 状态" & vbNewLine & _
                    "From 医技执行房间 a," & vbNewLine & _
                    "     (" & vbNewLine & _
                    "      Select 手术间,Max(手术状态) As 手术状态" & vbNewLine & _
                    "      From 病人手术记录" & vbNewLine & _
                    "      Where Not (手术结束时间<[2] OR 手术开始时间>[3]) AND 手术室id=[1] AND 手术状态 In (2,3) Group By 手术间" & vbNewLine & _
                    "     ) b" & vbNewLine & _
                    "Where a.科室id=[1]" & vbNewLine & _
                    "      And a.执行间=b.手术间(+)"
                        
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngDeptKey, CDate(Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(DateAdd("h", Val(txt(0).Text), dtp.Value))))
 
        If ShowPubSelect(Me, txt(1), 2, "执行间,2100,0,;状态,900,0,", Me.Name & "\手术执行间选择", "请从下表中选择一个手术执行间", rsData, rs, 3600, 4200) = 1 Then
            txt(1).Text = zlCommFun.NVL(rs("执行间").Value)
            DataChanged = True
        End If

    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    If ValidData = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    DataChanged = False
    
    Unload Me
End Sub

Private Sub mclsVsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf
        If .TextMatrix(Row, .ColIndex("岗位")) = "" And .TextMatrix(Row, .ColIndex("姓名")) = "" Then
            Cancel = True
        End If
    End With
End Sub

Private Sub txt_Change(Index As Integer)
    If mblnReading Then Exit Sub
    
    DataChanged = True

End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        zlCommFun.PressKey vbKeyTab

    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Cancel Then Exit Sub
    
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf.AfterEdit(Row, Col)
    DataChanged = True
    
    With vsf
        Select Case Col
        Case .ColIndex("岗位")
            If .ComboIndex > -1 Then
                .TextMatrix(Row, Col) = zlCommFun.GetNeedName(.ComboItem(.ComboIndex))
            End If
        End Select
    End With
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    Dim strTmp As String
    
    With vsf
        If Col = .ColIndex("姓名") Then

            strTmp = zlCommFun.GetNeedName(.TextMatrix(Row, .ColIndex("岗位")))
            
            gstrSQL = "Select 是否唯一,是否医生,是否护士 From 手术岗位 Where 名称=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTmp)
            If rs.BOF = False Then
                If zlCommFun.NVL(rs("是否医生").Value, 0) = 1 Then strTmp = "医生"
                If zlCommFun.NVL(rs("是否护士").Value, 0) = 1 Then strTmp = "护士"
            Else
                strTmp = "医生"
            End If
                
            gstrSQL = GetPublicSQL(SQL.人员安排选择)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strTmp, mlngDeptKey, mlngKey)
            bytRet = ShowPubSelect(Me, vsf, 2, "编号,1200,0,;姓名,1200,0,;简码,900,0,;科室,1200,0,;状态,900,0,", Me.Name & "\人员安排选择", "请从下表中选择一个手术人员", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                        
            If bytRet = 1 Then
            
'                If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
'                    ShowSimpleMsg "选择的人员“" & zlCommFun.NVL(rs("姓名").Value) & "”已被选择！"
'                    Exit Sub
'                End If
                       
                .EditText = zlCommFun.NVL(rs("姓名").Value)
                .TextMatrix(Row, .ColIndex("姓名")) = zlCommFun.NVL(rs("姓名").Value)
                .TextMatrix(Row, .ColIndex("编号")) = zlCommFun.NVL(rs("编号").Value)
                .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
    
                DataChanged = True
    
            End If
            
        End If
    End With
End Sub

Private Sub vsf_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    
    With vsf
        Select Case Col
        Case .ColIndex("岗位")
            
            Call mclsVsf.ComboLocation(Row, Col)

        End Select
    End With
End Sub

Private Sub vsf_DblClick()
    Call mclsVsf.DbClick
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
    
    
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytMode As Byte
    Dim bytRet As Byte
    Dim strDoctor As String
    
    With vsf
        If KeyCode = vbKeyReturn Then
            If Col = .ColIndex("姓名") Then
            
                If InStr(.EditText, "'") > 0 Then
                    KeyCode = 0
                    .EditText = ""
                    Exit Sub
                End If
    
                strTmp = zlCommFun.GetNeedName(.TextMatrix(Row, .ColIndex("岗位")))
                
                gstrSQL = "Select 是否唯一,是否医生,是否护士 From 手术岗位 Where 名称=[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTmp)
                If rs.BOF = False Then
                    If zlCommFun.NVL(rs("是否医生").Value, 0) = 1 Then strDoctor = "医生"
                    If zlCommFun.NVL(rs("是否护士").Value, 0) = 1 Then strDoctor = "护士"
                Else
                    strDoctor = "医生"
                End If
            
                strText = UCase(.EditText)
                bytMode = GetApplyMode(strText)
                
                strText = strText & "%"
                strTmp = IIf(ParamInfo.项目输入匹配方式 = 1, strText, "%" & strText)
                
                gstrSQL = GetPublicSQL(SQL.人员安排过滤, bytMode)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strDoctor, mlngDeptKey, mlngKey, strText, strTmp)
    
                If ShowPubSelect(Me, vsf, 2, "编号,1200,0,;姓名,1200,0,;简码,900,0,;科室,1200,0,;状态,900,0,", Me.Name & "\人员安排过滤", "请从下表中选择一个人员", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then

'                    If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
'                        ShowSimpleMsg "选择的人员“" & zlCommFun.NVL(rs("姓名").Value) & "”已被选择！"
'                        Exit Sub
'                    End If
                           
                    .EditText = zlCommFun.NVL(rs("姓名").Value)
                    .TextMatrix(Row, .ColIndex("姓名")) = zlCommFun.NVL(rs("姓名").Value)
                    .TextMatrix(Row, .ColIndex("编号")) = zlCommFun.NVL(rs("编号").Value)
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                    
                    DataChanged = True
                Else
                    .Cell(flexcpData, Row, Col) = .EditText
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    DataChanged = True
                End If

            End If
        Else
            DataChanged = True
        End If
    End With
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    Call mclsVsf.KeyPress(KeyAscii)
    
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsf.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsf.MouseRow, vsf.MouseCol)
    End Select
End Sub

Private Sub vsf_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub


