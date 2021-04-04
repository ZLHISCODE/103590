VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAdviceEditEx 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4275
   ControlBox      =   0   'False
   Icon            =   "frmAdviceEditEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton optPosition 
      Caption         =   "输入部位(&I)"
      Height          =   180
      Index           =   1
      Left            =   1455
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1995
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.OptionButton optPosition 
      Caption         =   "选择部位(&S)"
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1995
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CommandButton cmdData 
      Caption         =   "…"
      Height          =   240
      Left            =   2475
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "选择项目(*)"
      Top             =   1950
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Left            =   525
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3555
      Picture         =   "frmAdviceEditEx.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "取消(Esc)"
      Top             =   1920
      Width           =   450
   End
   Begin VB.ComboBox cboData 
      Height          =   300
      Left            =   525
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VSFlex8Ctl.VSFlexGrid vsExt 
      Align           =   1  'Align Top
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4275
      _cx             =   7541
      _cy             =   3254
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
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAdviceEditEx.frx":0596
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   240
         Left            =   3435
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   1035
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.ComboBox cbo标本 
      Height          =   300
      Left            =   525
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton cmdOK 
      Height          =   315
      Left            =   3015
      Picture         =   "frmAdviceEditEx.frx":06A2
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "确认(F2)"
      Top             =   1920
      Width           =   450
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "麻醉"
      Height          =   180
      Left            =   105
      TabIndex        =   10
      Top             =   1980
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frmAdviceEditEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入口参数：
Public mstrPrivs As String
Public mlngHwnd As Long '用于定位的控件句柄
Public mint期效 As Integer
Public mstr性别 As String
Public mint服务对象 As Integer '1-门诊,2-住院
'0-检查组合,1-手术输入,2-中药配方,3-检验标本,4-检验组合
Public mintType As Integer
'入/出:检查="部位ID1,部位ID2,..."
'      手术="手术ID1,手术ID2,...;麻醉ID",其中可能没有附加手术和麻醉
'      中药="中药ID1,单量1,脚注1;中药ID2,单量2,脚注2;...|煎法ID"
'      检验标本="项目ID1,项目ID2,...;检验标本"
'      检验组合="项目ID1,项目ID2,...;检验标本"
Public mstrExtData As String '新增时为空;医嘱新增检验时为"项目ID;"
'主诊疗项目ID,中药配方时为配方ID或单味中药ID,检验组合时表示诊疗单据ID
Public mlng项目ID As Long

'处方职务检查要求
Public mbln护士站 As Boolean '是否护士站调用
Public mbln医保 As Boolean '是否医保或公费病人

'出口参数：
Public mblnOK As Boolean '出

'程序变量
Private mlng中药房 As Long
Private mint简码 As Integer
Private mstrLike As String
Private mblnReturn As Boolean '是否了回车确认
Private mblnNotAddNew As Boolean '是否不允许增加
'-----------------------------------------------------------------------------------------------------
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long

Private Sub cboData_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboData.ListIndex <> -1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        lngIdx = zlControl.CboMatchIndex(cboData.Hwnd, KeyAscii)
        If lngIdx = -1 And cboData.ListCount > 0 Then lngIdx = 0
        cboData.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo标本_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo标本.ListIndex <> -1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        lngIdx = zlControl.CboMatchIndex(cbo标本.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo标本.ListCount > 0 Then lngIdx = 0
        cbo标本.ListIndex = lngIdx
    End If
End Sub

Private Sub cmd_Click()
'功能：打开项目选择器
    Dim rsTmp As ADODB.Recordset, i As Long
    Dim strSQL As String, int性别 As Integer, strSQLItem As String
    Dim strStock As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, str药品 As String
    Dim strSamples As String
    
    If mstr性别 Like "*男*" Then
        int性别 = 1
    ElseIf mstr性别 Like "*女*" Then
        int性别 = 2
    End If
    
    On Error GoTo errH
    
    If mintType = 0 And optPosition(1).Value Then
        '输入检查部位
        strSQL = _
            "Select A.ID, A.编码, A.标本部位 As 检查部位" & vbNewLine & _
            "From 诊疗项目目录 A, 诊疗项目组合 B" & vbNewLine & _
            "Where A.ID = B.诊疗项目id And B.诊疗组合id = [1] And A.服务对象 In ([2], 3) And" & vbNewLine & _
            "      (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & vbNewLine & _
            "Order By B.序号"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "部位", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, mlng项目ID, mint服务对象)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到可用的检查部位，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
        
        '检查重复输入
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "该检查部位已经在其它行录入。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Call Set部位输入(vsExt.Row, rsTmp)
    ElseIf mintType = 1 Then
        '输入附加手术:这里不是单独应用,因此不限制
        strSQLItem = _
            " From 诊疗项目目录 A Where A.类别='F' And A.ID<>" & mlng项目ID & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                " And A.服务对象 IN([1],3) And Nvl(A.执行频率,0) IN(0,[2]) And Nvl(A.适用性别,0) IN(0,[3])"
        
        strSQL = "Select Distinct 0 as 末级,ID,上级ID,编码,名称,NULL as 单位,NULL as 规模" & _
            " From 诊疗分类目录 Where 类型=5" & _
            " Start With ID In (Select 分类ID" & strSQLItem & ") Connect by Prior 上级ID=ID"
        strSQL = strSQL & " Union ALL" & _
            " Select Distinct 1 as 末级,A.ID,分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 规模" & _
            strSQLItem & " Order By 编码"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "手术", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mint服务对象, IIF(mint期效 = 0, 2, 1), int性别)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到可用的手术项目，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
        
        '检查重复输入
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "该附加手术已经在其它行录入。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Call Set手术输入(vsExt.Row, rsTmp)
    ElseIf mintType = 2 And CellCanEdit(vsExt.Row, vsExt.Col) Then
        If vsExt.Col Mod 4 = 0 Then
            '中药库存,中药房未指定时,读不出库存记录
            If mlng中药房 <> 0 Then
                strStock = _
                    "Select 药品ID,Sum(Nvl(可用数量,0)) as 库存 From 药品库存" & _
                    " Where (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate))" & _
                    " And 性质 = 1 And 库房ID=" & mlng中药房 & _
                    " Group by 药品ID" & _
                    " Having Sum(Nvl(可用数量,0))<>0"
            Else
                strStock = "Select NULL as 药品ID,NULL as 库存 From Dual"
            End If
            
            '特殊药品权限
            str药品 = ""
            If InStr(mstrPrivs, "下达麻醉药嘱") = 0 Then
                str药品 = str药品 & " And E.毒理分类<>'麻醉药'"
            End If
            If InStr(mstrPrivs, "下达毒性药嘱") = 0 Then
                str药品 = str药品 & " And E.毒理分类<>'毒性药'"
            End If
            If InStr(mstrPrivs, "下达贵重药嘱") = 0 Then
                str药品 = str药品 & " And E.价值分类 Not IN('贵重','昂贵')"
            End If
            
            '选择单味中草药:这里不是单独应用,因此不限制
            strSQL = "Select 0 as 末级,-1 as ID,-NULL as 上级ID,NULL as 编码," & _
                " CHR(13)||'常用中药' as 名称,NULL as 单位,NULL as 规格,NULL as 产地,NULL as 库存,NULL as 处方职务ID From Dual"
            strSQL = strSQL & " Union ALL" & _
                " Select 0 as 末级,ID,上级ID,编码,名称,NULL as 单位,NULL as 规格,NULL as 产地,NULL as 库存,NULL as 处方职务ID" & _
                " From 诊疗分类目录 Where 类型=3" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID"
            strSQL = strSQL & " Union ALL" & _
                " Select Distinct 1 as 末级,A.ID,A.分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位,D.规格,D.产地," & _
                " Decode(X.库存,NULL,NULL,X.库存/C.住院包装||C.住院单位) AS 库存,E.处方职务 as 处方职务ID" & _
                " From 诊疗项目目录 A,药品特性 E,药品规格 C,收费项目目录 D,(" & strStock & ") X" & _
                " Where A.类别='7' And A.ID=E.药名ID And A.ID=C.药名ID And C.药品ID=D.ID" & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) And A.服务对象 IN([1],3)" & _
                    " And (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 IS NULL) And D.服务对象 IN([1],3)" & _
                    " And Nvl(A.执行频率,0) IN(0,[2])" & str药品 & _
                    " And Nvl(A.适用性别,0) IN(0,[3]) And C.药品ID=X.药品ID(+)"
            strSQL = strSQL & " Union ALL" & _
                " Select Distinct 1 as 末级,A.ID,-1 as 上级ID,A.编码,A.名称,A.计算单位 as 单位,D.规格,D.产地," & _
                " Decode(X.库存,NULL,NULL,X.库存/C.住院包装||C.住院单位) AS 库存,E.处方职务 as 处方职务ID" & _
                " From 诊疗项目目录 A,药品特性 E,药品规格 C,收费项目目录 D,诊疗个人项目 T,(" & strStock & ") X" & _
                " Where A.类别='7' And A.ID=E.药名ID And A.ID=C.药名ID And C.药品ID=D.ID" & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) And A.服务对象 IN([1],3)" & _
                    " And (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 IS NULL) And D.服务对象 IN([1],3)" & _
                    " And Nvl(A.执行频率,0) IN(0,[2])" & str药品 & _
                    " And Nvl(A.适用性别,0) IN(0,[3]) And C.药品ID=X.药品ID(+)" & _
                    " And T.诊疗项目ID=A.ID And T.人员ID=[4]"
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "中药", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
                mint服务对象, IIF(mint期效 = 0, 2, 1), int性别, UserInfo.ID)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到可用的中药项目，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
                End If
                Exit Sub
            End If
            
            '检查重复输入
            If ItemExist(rsTmp!ID, vsExt.Row, vsExt.Col) Then
                MsgBox "该味中药在配方中已经录入。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '处方职务检查
            If Not mbln护士站 Then
                strSQL = CheckOneDuty(rsTmp!名称, Nvl(rsTmp!处方职务ID), UserInfo.姓名, mbln医保)
                If strSQL <> "" Then
                    MsgBox strSQL, vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            
            '获取输入值
            vsExt.TextMatrix(vsExt.Row, vsExt.Col) = rsTmp!名称
            vsExt.TextMatrix(vsExt.Row, vsExt.Col + 2) = rsTmp!单位
            vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = vsExt.TextMatrix(vsExt.Row, vsExt.Col)
            vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col + 2) = CLng(rsTmp!ID) '记录中药ID
            
            Call EnterNextCell(vsExt.Row, vsExt.Col)
        ElseIf vsExt.Col Mod 4 = 3 Then
            '选择脚注
            strSQL = "Select Rownum as ID,编码,名称,简码 From 中药煎服脚注 Order by 编码"
            vPoint = GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "脚注", , , , , , True, vPoint.x, vPoint.y, vsExt.CellHeight, blnCancel, , True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到可用的煎服脚注，请先到基础编码管理中设置。", vbInformation, gstrSysName
                End If
                Exit Sub
            End If
            
            '获取输入值
            vsExt.TextMatrix(vsExt.Row, vsExt.Col) = rsTmp!名称
            vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = vsExt.TextMatrix(vsExt.Row, vsExt.Col)
            
            Call EnterNextCell(vsExt.Row, vsExt.Col)
        End If
    ElseIf mintType = 4 Then
        '检验项目
        With Me.cbo标本
            For i = 0 To .ListCount - 1
                strSamples = strSamples & ",'" & .List(i) & "'"
            Next
        End With
        If Len(strSamples) > 0 Then
            strSamples = Mid(strSamples, 2)
        Else
            strSamples = "''"
        End If
        If mlng项目ID > 0 Then '指定了诊疗单据
            strSQLItem = "From 诊疗项目目录 A,诊疗单据应用 B,检验项目参考 C,检验报告项目 D " & _
                "Where A.ID=B.诊疗项目ID And A.id=D.诊疗项目id(+) And D.报告项目ID=C.项目id(+)" & _
                " And B.应用场合=[2] And B.病历文件ID=[1]" & _
                " And A.类别='C' And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) In (0,[3])" & _
                " And A.服务对象 IN([2],3) And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                " And (C.标本类型 In (" & strSamples & ") Or C.标本类型 Is Null)"
        Else
            strSQLItem = "From 诊疗项目目录 A,诊疗单据应用 B,检验项目参考 C,检验报告项目 D " & _
                "Where A.ID=B.诊疗项目ID(+) And A.id=D.诊疗项目id(+) And D.报告项目ID=C.项目id(+)" & _
                " And A.类别='C' And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) In (0,[3])" & _
                " And A.服务对象 IN([2],3) And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                " And (C.标本类型 In (" & strSamples & ") Or C.标本类型 Is Null)"
'                " And (B.诊疗项目ID is Null Or B.应用场合=" & mint服务对象 & ")"
        End If
        
        strSQL = "Select Distinct 0 as 末级,ID,上级ID,编码,名称,' ' As 检验类型,0 As 病历文件ID,' ' As 标本部位" & _
            " From 诊疗分类目录 Where 类型=5" & _
            " Start With ID In (Select A.分类ID " & strSQLItem & ") Connect by Prior 上级ID=ID"
        strSQL = strSQL & " Union ALL" & _
            " Select Distinct 1 as 末级,A.ID,分类ID as 上级ID,A.编码,A.名称,A.操作类型 as 检验类型,B.病历文件ID,A.标本部位 " & strSQLItem & " Order By 编码"
        
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "检验项目", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mlng项目ID, mint服务对象, int性别)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到可用的检验项目，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
        If rsTmp("检验类型") = "微生物" And vsExt.Rows > 2 Then
            If vsExt.RowData(2) <> 0 Or vsExt.Row > 1 Then '整个申请只能开一个微生物项目
                MsgBox "微生物项目只能单独申请！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        If mlng项目ID = 0 Then mlng项目ID = Nvl(rsTmp("病历文件ID"), 0)
        
        '检查重复输入
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "该检验项目已经录入！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '检查检验类型是否相同
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 And i <> vsExt.Row Then
                If Not (vsExt.TextMatrix(i, 1) = Nvl(rsTmp!检验类型) _
                    Or vsExt.TextMatrix(i, 1) = "" Or Nvl(rsTmp!检验类型) = "") Then
                    MsgBox "请输入相同检验类型的项目，已输入项目的检验类型为""" & vsExt.TextMatrix(i, 1) & """。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
        
        '重新初始标本
        If Not InitCombox(rsTmp("ID"), Nvl(rsTmp("标本部位"))) Then Exit Sub
        
        Call Set检验项目(vsExt.Row, rsTmp)
        If rsTmp("检验类型") = "微生物" Then
            mblnNotAddNew = True
            vsExt.Rows = 2
        Else
            mblnNotAddNew = False
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdData_Click()
'功能：打开项目选择器
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str性别 As String, blnCancel As Boolean
    Dim strSQLItem As String
    
    If mstr性别 Like "*男*" Then
        str性别 = "0,1"
    ElseIf mstr性别 Like "*女*" Then
        str性别 = "0,2"
    Else
        str性别 = "0"
    End If
    
    If mintType = 1 Then
        '输入麻醉项目:这里不是单独应用,因此不限制
        strSQLItem = " From 诊疗项目目录 A Where A.类别='G'" & _
                " And A.服务对象 IN([2],3) And A.ID<>[1]" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)"

        strSQL = "Select Distinct 0 as 末级,ID,上级ID,编码,名称,NULL as 单位,NULL as 麻醉类型" & _
            " From 诊疗分类目录 Where 类型=5" & _
            " Start With ID In (Select 分类ID" & strSQLItem & ") Connect by Prior 上级ID=ID"
        strSQL = strSQL & " Union ALL" & _
            " Select Distinct 1 as 末级,A.ID,分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 麻醉类型" & _
            strSQLItem & " Order By 编码"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "麻醉项目", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mlng项目ID, mint服务对象)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到匹配项目！", vbInformation, gstrSysName
            End If
            txtData.SetFocus: Exit Sub
        End If
        txtData.Tag = rsTmp!ID
        txtData.Text = "[" & rsTmp!编码 & "]" & rsTmp!名称
        cmdData.Tag = txtData.Text
        
        txtData.SetFocus
    ElseIf mintType = 4 Then
        '输入标本
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim str中药IDs As String, blnSkip As Boolean
    Dim strMsg As String, strTmp As String
    Dim strSQL As String, i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    
    If mintType = 0 Then '检查部位组合
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 Then
                If optPosition(0).Value Then
                    If Val(vsExt.TextMatrix(i, vsExt.Cols - 1)) <> 0 Then
                        strTmp = strTmp & "," & vsExt.RowData(i)
                    End If
                Else
                    strTmp = strTmp & "," & vsExt.RowData(i)
                End If
            End If
        Next
        strTmp = Mid(strTmp, 2)
        If strTmp = "" Then
            MsgBox "至少需要一个检查部位。", vbInformation, gstrSysName
            vsExt.SetFocus: Exit Sub
        End If
    ElseIf mintType = 1 Or mintType = 4 Then '附加手术及麻醉项目；检验项目及标本
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 Then
                strTmp = strTmp & "," & vsExt.RowData(i)
            End If
        Next
        strTmp = Mid(strTmp, 2)
        If strTmp = "" And mintType = 4 Then
            MsgBox "至少要选择一个检验项目。", vbInformation, gstrSysName
            vsExt.SetFocus: Exit Sub
        End If
        strTmp = strTmp & ";" & IIF(mintType = 4, Me.cbo标本.Text, IIF(Val(txtData.Tag) = 0, "", Val(txtData.Tag)))
    ElseIf mintType = 2 Then '中药配方项及煎法
        blnSkip = False
        For i = vsExt.FixedRows To vsExt.Rows - 1
            For j = 0 To vsExt.Cols - 1 Step 4
                If CLng(vsExt.Cell(flexcpData, i, j + 2)) <> 0 Then
                    If Val(vsExt.TextMatrix(i, j + 1)) = 0 Then
                        If Not blnSkip Then
                            If MsgBox("""" & vsExt.TextMatrix(i, j) & """没有输入单味用量，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                vsExt.Row = i: vsExt.Col = j + 1
                                Call vsExt.ShowCell(i, j + 1)
                                vsExt.SetFocus: Exit Sub
                            End If
                            blnSkip = True
                        End If
                    End If
                    If Val(vsExt.TextMatrix(i, j + 1)) <> 0 Then
                        strTmp = strTmp & ";" & vsExt.Cell(flexcpData, i, j + 2) & "," & vsExt.TextMatrix(i, j + 1) & "," & vsExt.TextMatrix(i, j + 3)
                        str中药IDs = str中药IDs & "," & CLng(vsExt.Cell(flexcpData, i, j + 2))
                    End If
                End If
            Next
        Next
        strTmp = Mid(strTmp, 2)
        str中药IDs = Mid(str中药IDs, 2)
        
        If strTmp = "" Then
            MsgBox "请在配方中至少输入一味中药。", vbInformation, gstrSysName
            vsExt.Row = vsExt.FixedRows: vsExt.Col = 0
            vsExt.SetFocus: Exit Sub
        End If
        If cboData.ListIndex = -1 Then
            MsgBox "请确定中药配方的煎法。", vbInformation, gstrSysName
            cboData.SetFocus: Exit Sub
        End If
        
        '处方职务检查
        If Not mbln护士站 Then
            strSQL = "Select 药名ID,处方职务 From 药品特性 Where 药名ID IN(" & str中药IDs & ")"
            On Error GoTo errH
            Set rsTmp = New ADODB.Recordset
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption) 'IN
            For i = vsExt.FixedRows To vsExt.Rows - 1
                For j = 0 To vsExt.Cols - 1 Step 4
                    If CLng(vsExt.Cell(flexcpData, i, j + 2)) <> 0 Then
                        If Val(vsExt.TextMatrix(i, j + 1)) <> 0 Then
                            rsTmp.Filter = "药名ID=" & CLng(vsExt.Cell(flexcpData, i, j + 2))
                            If Not rsTmp.EOF Then
                                strMsg = CheckOneDuty(vsExt.TextMatrix(i, j), Nvl(rsTmp!处方职务), UserInfo.姓名, mbln医保)
                                If strMsg <> "" Then
                                    vsExt.Row = i: vsExt.Col = j
                                    Call vsExt.ShowCell(i, j)
                                    MsgBox strMsg, vbInformation, gstrSysName
                                    vsExt.SetFocus: Exit Sub
                                End If
                            End If
                        End If
                    End If
                Next
            Next
        End If
        
        '药品禁忌检查
        If Not Check中药禁忌(str中药IDs) Then Exit Sub
        
        strTmp = strTmp & "|" & cboData.ItemData(cboData.ListIndex)
    ElseIf mintType = 3 Then '检验标本
        For i = 1 To vsExt.Rows - 1
            If Val(vsExt.TextMatrix(i, 1)) <> 0 Then
                strTmp = vsExt.TextMatrix(i, 0)
                Exit For
            End If
        Next
        If strTmp = "" Then
            MsgBox "请选择该检验项目的检验标本。", vbInformation, gstrSysName
            vsExt.SetFocus: Exit Sub
        End If
    End If
    
    mstrExtData = strTmp
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mintType = 0 Then
        optPosition(0).TabStop = False: optPosition(1).TabStop = False '不然无效
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = vbKeyF2 Then
        If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
    ElseIf mintType = 0 And Shift = vbCtrlMask And KeyCode = vbKeyA Then
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 Then vsExt.TextMatrix(i, 2) = 1
        Next
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '不允许输入分隔符及单引号
    If InStr(",;|'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim blnMulti As Boolean, vRect As RECT
    
    Call zlControl.CboSetHeight(cboData, Me.Height * 2)
    Call zlControl.CboSetWidth(cboData.Hwnd, cboData.Width * 1.2)
    
    '输入匹配
    mstrLike = IIF(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    mint简码 = Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "简码生成", 0)) '简码匹配方式：0-拼音,1-五笔
    If mint服务对象 = 0 Then mint服务对象 = 2 '缺省为住院
    mblnOK = False
    mblnNotAddNew = False
            
    '初始化表格样式
    If mintType = 0 Then
        optPosition(0).Visible = True: optPosition(1).Visible = True
        optPosition(Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "部位确定方式", 0))).Value = True
        If Not Init检查组合 Then Unload Me: Exit Sub
    ElseIf mintType = 1 Then
        lblData.Visible = True
        txtData.Visible = True
        cmdData.Visible = True
        lblData.Caption = "麻醉"
        If Not Init手术项目 Then Unload Me: Exit Sub
    ElseIf mintType = 2 Then
        mlng中药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(mint服务对象 = 1, "门诊", "住院") & "缺省中药房", 0))
        
        lblData.Visible = True
        cboData.Visible = True
        lblData.Caption = "煎法"
        If Not Init中药配方 Then Unload Me: Exit Sub
    ElseIf mintType = 3 Then
        If Not Init检验标本 Then Unload Me: Exit Sub
    ElseIf mintType = 4 Then
        lblData.Visible = True
        lblData.Caption = "标本"
        With cbo标本
            .Left = txtData.Left: .Top = txtData.Top: .Width = txtData.Width
            .Visible = True
        End With
        If Not Init检验组合 Then Unload Me: Exit Sub
        If Not InitCombox(DefaultValue:=Me.txtData) Then Unload Me: Exit Sub
    
        blnMulti = GetSysParVal(84) = "1" '是否允许一条医嘱申请多个检验项目
        
        If Len(Trim(mstrExtData)) > 0 Then
            If Len(Trim(Split(mstrExtData, ";")(0))) > 0 And Not blnMulti Then
                vsExt.Enabled = False
                '如果只有一个标本则不显示本窗口
                If cbo标本.ListCount < 2 Then cmdOK_Click: Exit Sub
            End If
        End If
    End If
    
    '窗体定位
    GetWindowRect mlngHwnd, vRect
    Me.Left = (vRect.Left - 1) * Screen.TwipsPerPixelX
    Me.Top = (vRect.Top - 1) * Screen.TwipsPerPixelY - Me.Height
    
    Call Form_Resize
End Sub

Private Function Init中药配方() As Boolean
'功能：初始化中药配方表格格式及数据
'参数：mstrExtData=包含每味中药信息及煎法信息的串,为空时表示新输入中药配方
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim lngRow As Long, lngCol As Long
    Dim str中药IDs As String, lng煎法ID As Long
    Dim arr中药 As Variant

    vsExt.Clear
    vsExt.Cols = 12: vsExt.Rows = 7
    vsExt.FixedCols = 0: vsExt.FixedRows = 1
    vsExt.ColWidth(0) = 795: vsExt.ColAlignment(0) = 1 '单味中药
    vsExt.ColWidth(1) = 450: vsExt.ColAlignment(1) = 7 '单味用量
    vsExt.ColWidth(2) = 285: vsExt.ColAlignment(2) = 1 '单位
    vsExt.ColWidth(3) = 750: vsExt.ColAlignment(3) = 1 '脚注
    For i = 4 To vsExt.Cols - 1
        vsExt.ColWidth(i) = vsExt.ColWidth(i - 4)
        vsExt.ColAlignment(i) = vsExt.ColAlignment(i - 4)
    Next
    vsExt.MergeCells = flexMergeFixedOnly
    vsExt.MergeRow(0) = True
    vsExt.Cell(flexcpAlignment, 0, 0, 0, vsExt.Cols - 1) = 4
    vsExt.Cell(flexcpText, 0, 0, 0, vsExt.Cols - 1) = "请依次输入中草药,单味用量,脚注。按*键选取中药或脚注。"
    
    Me.Width = (Me.Width - Me.ScaleWidth) + 2280 * 3 + 250
    vsExt.GridColor = vsExt.BackColor
    vsExt.Editable = flexEDKbdMouse
    
    On Error GoTo errH
    
    If mstrExtData <> "" Then '修改
        lng煎法ID = Val(Split(mstrExtData, "|")(1))
        arr中药 = Split(Split(mstrExtData, "|")(0), ";")
        
        For i = 0 To UBound(arr中药)
            str中药IDs = str中药IDs & "," & CStr(Split(arr中药(i), ",")(0))
        Next
        str中药IDs = Mid(str中药IDs, 2)
        
        strSQL = "Select A.ID,A.名称,A.计算单位 From 诊疗项目目录 A Where ID IN(" & str中药IDs & ")"
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'IN
        
        If vsExt.Rows < -Int(rsTmp.RecordCount / -3) + 1 Then
            vsExt.Rows = -Int(rsTmp.RecordCount / -3) + 1
        End If
        lngRow = vsExt.FixedRows: lngCol = 0
        
        '按照现在的内容和次序显示
        For i = 0 To UBound(arr中药)
            rsTmp.Filter = "ID=" & CStr(Split(arr中药(i), ",")(0))
            If Not rsTmp.EOF Then
                vsExt.TextMatrix(lngRow, lngCol) = rsTmp!名称
                vsExt.TextMatrix(lngRow, lngCol + 1) = CStr(Split(arr中药(i), ",")(1))
                vsExt.TextMatrix(lngRow, lngCol + 2) = Nvl(rsTmp!计算单位)
                vsExt.TextMatrix(lngRow, lngCol + 3) = CStr(Split(arr中药(i), ",")(2))
                
                '用于恢复显示的记录
                vsExt.Cell(flexcpData, lngRow, lngCol) = vsExt.TextMatrix(lngRow, lngCol)
                vsExt.Cell(flexcpData, lngRow, lngCol + 1) = vsExt.TextMatrix(lngRow, lngCol + 1)
                vsExt.Cell(flexcpData, lngRow, lngCol + 2) = CLng(rsTmp!ID) '记录中药ID
                vsExt.Cell(flexcpData, lngRow, lngCol + 3) = vsExt.TextMatrix(lngRow, lngCol + 3)
                                
                '下一位置
                If lngCol + 4 > vsExt.Cols - 1 Then
                    lngRow = lngRow + 1: lngCol = 0
                Else
                    lngCol = lngCol + 4
                End If
            End If
        Next
    Else '新增
        strSQL = "Select ID,类别,名称,计算单位 From 诊疗项目目录 Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng项目ID)
        If rsTmp!类别 = "7" Then
            '输入了单味中草药
            vsExt.TextMatrix(vsExt.FixedRows, 0) = rsTmp!名称
            vsExt.TextMatrix(vsExt.FixedRows, 2) = Nvl(rsTmp!计算单位)
            
            '用于恢复显示的记录
            vsExt.Cell(flexcpData, vsExt.FixedRows, 0) = vsExt.TextMatrix(vsExt.FixedRows, 0)
            vsExt.Cell(flexcpData, vsExt.FixedRows, 2) = CLng(rsTmp!ID) '记录中药ID
        Else
            '输入了配方项目
            strSQL = "Select A.ID,A.名称,A.计算单位,B.单次用量,B.医生嘱托" & _
                " From 诊疗项目目录 A,诊疗项目组合 B,药品规格 C,收费项目目录 D" & _
                " Where A.ID=B.诊疗项目ID And A.ID=C.药名ID And C.药品ID=D.ID And B.诊疗组合ID=[1]" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) And A.服务对象 IN([2],3)" & _
                " And (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 is NULL) And D.服务对象 IN([2],3)" & _
                " Order by B.序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng项目ID, mint服务对象)
            If rsTmp.EOF Then
                MsgBox "该中药配方当前无有效的配方组成，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
                Exit Function
            End If
            
            If vsExt.Rows < -Int(rsTmp.RecordCount / -3) + 1 Then
                vsExt.Rows = -Int(rsTmp.RecordCount / -3) + 1
            End If
            lngRow = vsExt.FixedRows: lngCol = 0
            
            '按照设置的内容的次序显示
            For i = 1 To rsTmp.RecordCount
                vsExt.TextMatrix(lngRow, lngCol) = rsTmp!名称
                vsExt.TextMatrix(lngRow, lngCol + 1) = Nvl(rsTmp!单次用量)
                vsExt.TextMatrix(lngRow, lngCol + 2) = Nvl(rsTmp!计算单位)
                vsExt.TextMatrix(lngRow, lngCol + 3) = Nvl(rsTmp!医生嘱托)
                
                '用于恢复显示的记录
                vsExt.Cell(flexcpData, lngRow, lngCol) = vsExt.TextMatrix(lngRow, lngCol)
                vsExt.Cell(flexcpData, lngRow, lngCol + 1) = vsExt.TextMatrix(lngRow, lngCol + 1)
                vsExt.Cell(flexcpData, lngRow, lngCol + 2) = CLng(rsTmp!ID) '记录中药ID
                vsExt.Cell(flexcpData, lngRow, lngCol + 3) = vsExt.TextMatrix(lngRow, lngCol + 3)
                
                '下一位置
                If lngCol + 4 > vsExt.Cols - 1 Then
                    lngRow = lngRow + 1: lngCol = 0
                Else
                    lngCol = lngCol + 4
                End If
                rsTmp.MoveNext
            Next
            
            '获取配方项目的缺省煎法
            strSQL = "Select 用法ID From 诊疗用法用量 Where 性质=1 And 项目ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng项目ID)
            If Not rsTmp.EOF Then lng煎法ID = rsTmp!用法ID
        End If
    End If
        
    '中药煎法
    strSQL = "Select A.ID,A.编码,A.名称 From 诊疗项目目录 A" & _
        " Where A.类别='E' And A.操作类型='3' And A.服务对象 IN([1],3)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint服务对象)
    If rsTmp.Filter <> 0 Then rsTmp.Filter = 0
    If rsTmp.EOF Then
        MsgBox "未找到有效的中药煎法，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    
    For i = 1 To rsTmp.RecordCount
        cboData.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboData.ItemData(cboData.NewIndex) = rsTmp!ID
        If rsTmp!ID = lng煎法ID Then
            Call zlControl.CboSetIndex(cboData.Hwnd, cboData.NewIndex)
        End If
        rsTmp.MoveNext
    Next
    
    Call SetSplitLine
    vsExt.Row = vsExt.FixedRows: vsExt.Col = 0
    Init中药配方 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetSplitLine()
'功能：设置中药配方输入的三列分隔线
    Dim lngRow As Long, lngCol As Long
        
    vsExt.Redraw = False
    lngRow = vsExt.Row: lngCol = vsExt.Col
    
    vsExt.Select vsExt.FixedRows, 3, vsExt.Rows - 1, 3
    vsExt.CellBorder &HC0C0C0, 0, 0, 1, 0, 0, 0
    vsExt.Select vsExt.FixedRows, 7, vsExt.Rows - 1, 7
    vsExt.CellBorder &HC0C0C0, 0, 0, 1, 0, 0, 0

    vsExt.Row = lngRow: vsExt.Col = lngCol
    vsExt.Redraw = True
End Sub

Private Function Init手术项目() As Boolean
'功能：初始化手术表格格式及数据
'参数：mstrExtData=包含附加手术及麻醉项目的信息,其中可能没有附加手术；为空时表示新输入手术项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng麻醉ID As Long
    Dim arr手术IDs As Variant, str手术IDs As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    strSQL = mstrExtData
    If strSQL = "" Then strSQL = ";"
    str手术IDs = CStr(Split(strSQL, ";")(0))
    lng麻醉ID = Val(Split(strSQL, ";")(1))
    
    '附加手术
    If str手术IDs <> "" Then
        strSQL = "Select A.ID,A.编码,A.名称,A.操作类型" & _
            " From 诊疗项目目录 A" & _
            " Where A.类别='F' And A.ID IN(" & str手术IDs & ")" & _
            " Order by A.编码"
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'IN
        i = rsTmp.RecordCount
    End If
        
    vsExt.Clear
    vsExt.Rows = IIF(i = 0, 2, i + 1)
    vsExt.Cols = 2
    vsExt.FixedRows = 1: vsExt.FixedCols = 0
    vsExt.TextMatrix(0, 0) = "附加手术"
    vsExt.TextMatrix(0, 1) = "规模"
    vsExt.ColWidth(0) = 3200: vsExt.ColWidth(1) = 800
    vsExt.FixedAlignment(0) = 4: vsExt.FixedAlignment(1) = 4
    vsExt.ColAlignment(0) = 1: vsExt.ColAlignment(1) = 1
    vsExt.Editable = flexEDKbdMouse
    
    If str手术IDs <> "" And i <> 0 Then
        arr手术IDs = Split(str手术IDs, ",") '按照原有输入顺序
        For i = 0 To UBound(arr手术IDs)
            rsTmp.Filter = "ID=" & CStr(arr手术IDs(i))
            If Not rsTmp.EOF Then
                j = j + 1
                vsExt.RowData(j) = CLng(rsTmp!ID)
                vsExt.TextMatrix(j, 0) = "[" & rsTmp!编码 & "]" & rsTmp!名称
                vsExt.Cell(flexcpData, j, 0) = vsExt.TextMatrix(j, 0) '用于恢复显示
                vsExt.TextMatrix(j, 1) = Nvl(rsTmp!操作类型, 0)
            End If
        Next
    End If
    
    '麻醉项目
    If lng麻醉ID <> 0 Then
        strSQL = "Select A.ID,A.编码,A.名称,操作类型 From 诊疗项目目录 A Where A.类别='G' And A.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng麻醉ID)
        If rsTmp.Filter <> 0 Then rsTmp.Filter = 0
        If Not rsTmp.EOF Then
            txtData.Tag = rsTmp!ID
            txtData.Text = "[" & rsTmp!编码 & "]" & rsTmp!名称
            cmdData.Tag = txtData.Text '用于恢复显示
        End If
    End If
    
    vsExt.Row = 1: vsExt.Col = 0
    Init手术项目 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init检查组合() As Boolean
'功能：初始化检查部位表格格式及数据
'参数：mstrExtData=包含检查部位的信息,为空时表示新输入检查组合项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strPosition As String, arrPosition As Variant
    
    On Error GoTo errH
    
    If Not Visible Then
        strPosition = mstrExtData
    Else
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 Then
                If vsExt.Cols = 3 Then
                    If Val(vsExt.TextMatrix(i, 2)) <> 0 Then
                        strPosition = strPosition & "," & vsExt.RowData(i)
                    End If
                Else
                    strPosition = strPosition & "," & vsExt.RowData(i)
                End If
            End If
        Next
        strPosition = Mid(strPosition, 2)
    End If
    
    '按照设置的部位顺序号
    strSQL = "Select A.类别,A.编码,A.名称,A.标本部位,B.诊疗项目ID" & _
        " From 诊疗项目目录 A,诊疗项目组合 B" & _
        " Where A.ID=B.诊疗项目ID And B.诊疗组合ID=[1]" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And A.服务对象 IN([2],3)" & _
        " Order by B.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng项目ID, mint服务对象)
    If rsTmp.EOF Then
        MsgBox "该检查组合项目当前无有效部位，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    
    With vsExt
        .FixedRows = 0: .FixedCols = 0
        .Rows = 0: .Cols = 0
        If optPosition(0).Value Then '部位选择模式
            .Rows = IIF(rsTmp.EOF, 2, rsTmp.RecordCount + 1)
            .FixedRows = 1: .Cols = 3: .FixedCols = 0
            
            .TextMatrix(0, 0) = "检查项目"
            .TextMatrix(0, 1) = "检查部位"
            .TextMatrix(0, 2) = "选择"
            .FixedAlignment(0) = 4: .ColAlignment(0) = 1: .ColWidth(0) = 2000
            .FixedAlignment(1) = 4: .ColAlignment(1) = 1: .ColWidth(1) = 1500
            .FixedAlignment(2) = 4: .ColAlignment(2) = 4: .ColWidth(2) = 500
            .ColDataType(2) = flexDTBoolean
            .Editable = flexEDKbdMouse
            
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = CLng(rsTmp!诊疗项目ID) '一定要明确类型
                .TextMatrix(i, 0) = rsTmp!名称 '"[" & rsTmp!编码 & "]" & rsTmp!名称
                .TextMatrix(i, 1) = Nvl(rsTmp!标本部位)
                If InStr("," & strPosition & ",", "," & rsTmp!诊疗项目ID & ",") > 0 Then
                    .TextMatrix(i, 2) = 1
                End If
                rsTmp.MoveNext
            Next
                
            .Row = 1: .Col = 2
        Else '部位输入模式
            arrPosition = Split(strPosition, ",")
            .Rows = 1 + (UBound(arrPosition) + 1) + 1
            .FixedRows = 1: .Cols = 2: .FixedCols = 0
            
            .TextMatrix(0, 0) = "检查项目"
            .TextMatrix(0, 1) = "检查部位"
            .FixedAlignment(0) = 4: .ColAlignment(0) = 1: .ColWidth(0) = 2000
            .FixedAlignment(1) = 4: .ColAlignment(1) = 1: .ColWidth(1) = 2000
            .Editable = flexEDKbdMouse
                        
            For i = 0 To UBound(arrPosition)
                rsTmp.Filter = "诊疗项目ID=" & arrPosition(i)
                If Not rsTmp.EOF Then
                    .RowData(i + 1) = CLng(rsTmp!诊疗项目ID)
                    .TextMatrix(i + 1, 0) = rsTmp!名称
                    .TextMatrix(i + 1, 1) = Nvl(rsTmp!标本部位)
                    .Cell(flexcpData, i + 1, 1) = .TextMatrix(i + 1, 1) '用于恢复显示
                End If
            Next
            
            rsTmp.Filter = 0
            .TextMatrix(.Rows - 1, 0) = rsTmp!名称
            .Row = .Rows - 1: .Col = .Cols - 1
        End If
        .ShowCell .Row, .Col
        .LeftCol = 0 '要加在ShowCell后,不然选择模式有问题
    End With
    
    Init检查组合 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init检验标本() As Boolean
'功能：初始化检验标本表格格式及标本数据
'参数：mstrExtData=包含缺省的检验标本的信息,为空时表示新输入检验项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strTmp As String, lngItemCount As Long
    Dim aTmp() As String, strSample As String, blnChecked As Boolean
    
    On Error GoTo errH
    
    If Len(mstrExtData) > 0 Then
        aTmp = Split(mstrExtData, ";")
        strTmp = aTmp(0)
        lngItemCount = UBound(Split(strTmp, ",")) + 1
        If UBound(aTmp) > 0 Then strSample = aTmp(1)
    End If
    If lngItemCount = 0 Then
        strSQL = "Select 名称 From 诊疗检验标本"
    Else
        strSQL = _
            " Select 标本类型,Sum(1) From (" & _
            "   Select Distinct A.ID,B.名称 As 标本类型" & _
            "   From 诊疗项目目录 A,诊疗检验标本 B,检验项目参考 C,检验报告项目 D" & _
            "   Where A.ID=D.诊疗项目ID(+) And D.报告项目ID=C.项目ID(+)" & _
            "       And (C.标本类型 Is Null Or C.标本类型=B.名称) And A.ID In (" & strTmp & ")" & _
            " ) Group By 标本类型 Having Sum(1)=" & lngItemCount
    End If
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In
    If rsTmp.EOF Then
        MsgBox Switch(lngItemCount = 0, "未设置检验标本，请到字典管理工具中设置。", _
            lngItemCount = 1, "选取的检验项目未定义检验标本，请先到检验项目管理中设置", _
            lngItemCount > 1, "选取的检验项目的检验标本不一致，请先到检验项目管理中设置"), vbInformation, gstrSysName
        Exit Function
    ElseIf rsTmp.RecordCount = 1 And mstrExtData = "" Then
        '新输入项目时,如果只有一个标本时,直接选择退出
        mstrExtData = rsTmp(0)
        mblnOK = True: Exit Function
    End If
    
    vsExt.Clear
    vsExt.Rows = IIF(rsTmp.EOF, 2, rsTmp.RecordCount + 1)
    vsExt.FixedRows = 1: vsExt.Cols = 2: vsExt.FixedCols = 0
    vsExt.Row = 1: vsExt.Col = 1
    
    vsExt.TextMatrix(0, 0) = "检验标本"
    vsExt.TextMatrix(0, 1) = "选择"
    vsExt.FixedAlignment(0) = 4: vsExt.ColAlignment(0) = 1: vsExt.ColWidth(0) = 3500
    vsExt.FixedAlignment(1) = 4: vsExt.ColAlignment(1) = 4: vsExt.ColWidth(1) = 500
    vsExt.ColDataType(1) = flexDTBoolean
    vsExt.Editable = flexEDKbdMouse
    
    For i = 1 To rsTmp.RecordCount
        vsExt.TextMatrix(i, 0) = rsTmp(0)
        If strSample = vsExt.TextMatrix(i, 0) Then
            vsExt.TextMatrix(i, 1) = 1
            vsExt.Row = i
            blnChecked = True
        End If
        rsTmp.MoveNext
    Next
    If Not blnChecked Then
        vsExt.TextMatrix(1, 1) = 1
        vsExt.Row = 1
    End If
    
    Init检验标本 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init检验组合() As Boolean
'功能：初始化检验项目
'参数：mstrExtData=包含缺省的检验项目的信息,为空时表示新输入检验项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim arrItems As Variant, strItems As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    strSQL = mstrExtData
    If strSQL = "" Then strSQL = ";"
    strItems = CStr(Split(strSQL, ";")(0))
    Me.txtData = Split(strSQL, ";")(1)
    cmdData.Tag = txtData.Text
    
    If strItems <> "" Then
        If mlng项目ID > 0 Then '指定了诊疗单据
            strSQL = "Select A.* From 诊疗项目目录 A,诊疗单据应用 B " & _
                " Where A.ID=B.诊疗项目ID And B.应用场合=" & mint服务对象 & " And B.病历文件ID=" & mlng项目ID & _
                " And A.类别='C' And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) In (0" & IIF(Len(Trim(mstr性别)) = 0, ") ", IIF(mstr性别 Like "*男*", ",1) ", ",2) ")) & _
                " And A.服务对象 IN(" & mint服务对象 & ",3) And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                " And A.ID In(" & strItems & ")" & _
                " Order by A.编码"
        Else
            strSQL = "Select A.*,B.病历文件ID From 诊疗项目目录 A,诊疗单据应用 B " & _
                " Where A.ID=B.诊疗项目ID(+)" & _
                " And A.类别='C' And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) In (0" & IIF(Len(Trim(mstr性别)) = 0, ") ", IIF(mstr性别 Like "*男*", ",1) ", ",2) ")) & _
                " And A.服务对象 IN(" & mint服务对象 & ",3) And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                " And A.ID In(" & strItems & ")" & _
                " Order by A.编码"
'                " And (B.诊疗项目ID is Null Or B.应用场合=" & mint服务对象 & ")"
        End If
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In
        i = rsTmp.RecordCount
        If i > 0 And mlng项目ID = 0 Then mlng项目ID = Nvl(rsTmp("病历文件ID"), 0)
    End If
        
    vsExt.Clear
    vsExt.Rows = IIF(i = 0, 2, i + 1)
    vsExt.Cols = 2
    vsExt.FixedRows = 1: vsExt.FixedCols = 0
    vsExt.TextMatrix(0, 0) = "检验项目"
    vsExt.ColWidth(0) = 4000: vsExt.ColHidden(1) = True
    vsExt.FixedAlignment(0) = 4
    vsExt.ColAlignment(0) = 1
    vsExt.Editable = flexEDKbdMouse
    
    If i > 0 Then
        arrItems = Split(strItems, ",") '按照原有输入顺序
        For i = 0 To UBound(arrItems)
            rsTmp.Filter = "ID=" & arrItems(i)
            If Not rsTmp.EOF Then
                j = j + 1
                vsExt.RowData(j) = CLng(rsTmp!ID)
                vsExt.TextMatrix(j, 0) = "[" & rsTmp!编码 & "]" & rsTmp!名称
                vsExt.Cell(flexcpData, j, 0) = vsExt.TextMatrix(j, 0) '用于恢复显示
                vsExt.TextMatrix(j, 1) = Nvl(rsTmp!操作类型)
                If rsTmp("操作类型") = "微生物" Then mblnNotAddNew = True '微生物只能开一个检验项目
            End If
        Next
    End If
    
    vsExt.Row = 1: vsExt.Col = 0
    Init检验组合 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitCombox(Optional ByVal strNewItemID As String = "", Optional ByVal DefaultValue As String = "") As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strTmp As String, lngItemCount As Long
    InitCombox = False
    
    On Error GoTo DBError
    strTmp = "": lngItemCount = 0
    For i = 1 To vsExt.Rows - 1
        If vsExt.RowData(i) <> 0 And (i <> vsExt.Row Or Len(strNewItemID) = 0) Then
            lngItemCount = lngItemCount + 1
            strTmp = strTmp & "," & vsExt.RowData(i)
        End If
    Next
    If Len(strNewItemID) > 0 Then
        lngItemCount = lngItemCount + 1
        strTmp = strTmp & "," & strNewItemID
    End If
    If Len(strTmp) > 0 Then strTmp = Mid(strTmp, 2)

    If lngItemCount = 0 Then
        strSQL = "Select 名称 From 诊疗检验标本"
    Else
        strSQL = "Select 标本类型,Sum(1) From (" & _
            "   Select Distinct A.ID,B.名称 As 标本类型" & _
            "   From 诊疗项目目录 A,诊疗检验标本 B,检验项目参考 C,检验报告项目 D" & _
            "   Where A.ID=D.诊疗项目ID(+) And D.报告项目ID=C.项目ID(+)" & _
            "       And (C.标本类型 Is Null Or C.标本类型=B.名称) And A.ID In (" & strTmp & ")" & _
            " ) Group By 标本类型 Having Sum(1)=" & lngItemCount
    End If
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In
    If rsTmp.EOF Then
        MsgBox Switch(lngItemCount = 0, "未设置检验标本，请到字典管理工具中设置。", _
            lngItemCount = 1, "选取的检验项目未定义检验标本，请先到检验项目管理中设置", _
            lngItemCount > 1, "选取的检验项目的检验标本与其他项目的不一致，请先到检验项目管理中设置"), vbInformation, gstrSysName
        Exit Function
    End If
    
    With cbo标本
        strTmp = .Text
        
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp(0)
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
        On Error Resume Next
        If Len(DefaultValue) > 0 Then
            .Text = DefaultValue
        Else
            .Text = strTmp
        End If
    End With
    InitCombox = True
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Height - cmdCancel.Width
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
    mlngHwnd = 0
    mint期效 = 0
    mstr性别 = ""
    mintType = 0
    mlng项目ID = 0
    mint服务对象 = 0
    mbln护士站 = False
    mbln医保 = False
    
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "部位确定方式", IIF(optPosition(0).Value, 0, 1)
End Sub

Private Sub optPosition_Click(Index As Integer)
    If Visible Then
        Call Init检查组合: vsExt.SetFocus
        optPosition(0).TabStop = False: optPosition(1).TabStop = False '不然无效
    End If
End Sub

Private Sub txtData_GotFocus()
    zlControl.TxtSelAll txtData
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset, vRect As RECT
    Dim strSQL As String, str性别 As String
    Dim strLike As String, blnCancel As Boolean
    
    If mstr性别 Like "*男*" Then
        str性别 = "0,1"
    ElseIf mstr性别 Like "*女*" Then
        str性别 = "0,2"
    Else
        str性别 = "0"
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtData.Text = "" Then
            If mintType = 1 Then '手术可以不输入麻醉项目
                Call zlCommFun.PressKey(vbKeyTab)
            End If
            Exit Sub
        ElseIf txtData.Text = cmdData.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        
        '优化
        strLike = mstrLike
        If Len(txtData.Text) < 2 Then strLike = ""
        
        If mintType = 1 Then
            '输入麻醉项目
            strSQL = _
                " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 麻醉类型" & _
                " From 诊疗项目目录 A,诊疗项目别名 B" & _
                " Where A.ID=B.诊疗项目ID And A.类别='G' And A.服务对象 IN([3],3)" & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                    " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]) And B.码类=[4]" & _
                " Order by A.编码"
            vRect = GetControlRect(txtData.Hwnd)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "麻醉项目", False, "", "", False, False, True, vRect.Left, vRect.Top, txtData.Height, blnCancel, False, True, _
                UCase(txtData.Text) & "%", strLike & UCase(txtData.Text) & "%", mint服务对象, mint简码 + 1)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配项目！", vbInformation, gstrSysName
                End If
                txtData.Text = cmdData.Tag
                zlControl.TxtSelAll txtData
                Exit Sub
            End If
            txtData.Tag = rsTmp!ID
            txtData.Text = "[" & rsTmp!编码 & "]" & rsTmp!名称
            cmdData.Tag = txtData.Text
            
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf mintType = 4 Then
            '检验标本
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call cmdData_Click
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtData_Validate(Cancel As Boolean)
'功能：恢复显示原内容
    If txtData.Text <> cmdData.Tag Then
        txtData.Text = cmdData.Tag
    End If
End Sub

Private Sub vsExt_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'功能:显示选择按钮,并保证当前单元格可见
    
    '保证当前单元格可见
    If NewRow >= vsExt.FixedRows And NewRow <= vsExt.Rows - 1 Then
        If vsExt.LeftCol >= vsExt.FixedCols And vsExt.LeftCol <= vsExt.Cols - 1 Then
            Call vsExt.ShowCell(NewRow, vsExt.LeftCol)
        End If
    End If
    
    If mintType = 0 And optPosition(1).Value Then
        If NewCol = 1 Then
            cmd.Height = vsExt.CellHeight - 30
            cmd.Left = vsExt.CellLeft + vsExt.CellWidth - cmd.Width - 15
            cmd.Top = vsExt.CellTop + 15
            
            cmd.Visible = True
        Else
            cmd.Visible = False
        End If
    ElseIf mintType = 1 Or mintType = 4 Then
        '显示/隐藏手术选择按钮
        If NewCol = 0 Then
            cmd.Height = vsExt.CellHeight - 30
            cmd.Left = vsExt.CellLeft + vsExt.CellWidth - cmd.Width - 15
            cmd.Top = vsExt.CellTop + 15
            
            cmd.Visible = True
        Else
            cmd.Visible = False
        End If
    End If
End Sub

Private Sub vsExt_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'功能:限制某些列宽的范围
    If Row = -1 Then
        If mintType = 0 Then
            '选择列宽度不变
            If optPosition(0).Value Then
                If 3500 - vsExt.ColWidth(0) <= 0 Then vsExt.ColWidth(0) = 3000
                vsExt.ColWidth(1) = 3500 - vsExt.ColWidth(0)
            Else
                If 4000 - vsExt.ColWidth(0) <= 0 Then vsExt.ColWidth(0) = 2000
                vsExt.ColWidth(1) = 4000 - vsExt.ColWidth(0)
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) '使按钮可见及调整按钮位置
        End If
    End If
End Sub

Private Sub vsExt_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
    '单位列鼠标不可进入
    If mintType = 2 And Button = 1 And vsExt.MouseCol Mod 4 = 2 Then Cancel = True
End Sub

Private Sub vsExt_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mintType = 2 Then
        '单位列按键不可进入
        If NewCol Mod 4 = 2 Then
            Cancel = True
            If OldCol > NewCol Then '按键移动时跳过
                vsExt.Col = NewCol - 1
            Else
                vsExt.Col = NewCol + 1
            End If
            vsExt.Row = NewRow
        End If
    End If
End Sub

Private Sub vsExt_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If cmd.Visible Then cmd.Visible = False
End Sub

Private Sub vsExt_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'功能：限制某些列不能改变列宽
    If Row = -1 Then
        If mintType = 0 Then
            '只允许改变前两列列宽
            If Col <> 0 Then Cancel = True
        ElseIf mintType = 3 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub vsExt_GotFocus()
    Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) '使按钮可见
End Sub

Private Sub vsExt_KeyDown(KeyCode As Integer, Shift As Integer)
'功能：删除数据行
    Dim i As Long, j As Long, k As Long
    
    If KeyCode = vbKeyDelete Then
        If (mintType = 0 And optPosition(1).Value Or mintType = 1 Or mintType = 4) And vsExt.RowData(vsExt.Row) <> 0 Then
            If MsgBox("要删除当前行吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            vsExt.RowData(vsExt.Row) = 0
            If mintType = 0 Then
                vsExt.TextMatrix(vsExt.Row, vsExt.Cols - 1) = ""
                vsExt.Cell(flexcpData, vsExt.Row, vsExt.Cols - 1) = ""
            Else
                For i = 0 To vsExt.Cols - 1
                    vsExt.TextMatrix(vsExt.Row, i) = ""
                    vsExt.Cell(flexcpData, vsExt.Row, i) = ""
                Next
            End If
            If Not (vsExt.Rows = vsExt.FixedRows + 1 And vsExt.Row = vsExt.FixedRows) Then
                vsExt.RemoveItem vsExt.Row
            End If
            
            '重新初始标本
            If mintType = 4 Then InitCombox
        ElseIf mintType = 2 Then
            If CLng(vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)) <> 0 Then
                If MsgBox("要删除""" & vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                '清除当前味药信息
                For i = 0 To 3
                    vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4 + i) = ""
                    vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + i) = Empty
                Next
                '后面的内容向前移
                For i = vsExt.Row To vsExt.Rows - 1
                    For j = 0 To vsExt.Cols - 1 Step 4
                        If Not (i = vsExt.Row And j <= (vsExt.Col \ 4) * 4) Then
                            For k = 0 To 3
                                If j = 0 Then
                                    vsExt.TextMatrix(i - 1, vsExt.Cols - (4 - k)) = vsExt.TextMatrix(i, j + k)
                                    vsExt.Cell(flexcpData, i - 1, vsExt.Cols - (4 - k)) = vsExt.Cell(flexcpData, i, j + k)
                                Else
                                    vsExt.TextMatrix(i, j + k - 4) = vsExt.TextMatrix(i, j + k)
                                    vsExt.Cell(flexcpData, i, j + k - 4) = vsExt.Cell(flexcpData, i, j + k)
                                End If
                                vsExt.TextMatrix(i, j + k) = ""
                                vsExt.Cell(flexcpData, i, j + k) = Empty
                            Next
                        End If
                    Next
                Next
                '删除多余的空行(至少保留可以最多显示的行数7)
                If vsExt.Rows > 7 Then
                    For i = vsExt.Rows - 1 To 7 Step -1
                        If CLng(vsExt.Cell(flexcpData, i - 1, 2)) = 0 Then
                            vsExt.RemoveItem i
                        End If
                    Next
                End If
                Call vsExt.ShowCell(vsExt.Row, vsExt.Col)
            End If
        End If
    End If
End Sub

Private Sub vsExt_LostFocus()
    If Not ActiveControl Is cmd Then cmd.Visible = False
End Sub

Private Sub vsExt_KeyPress(KeyAscii As Integer)
'功能：非编辑状态时，自动移动单元格
    If KeyAscii = 13 Then
        KeyAscii = 0
        '定位到下一应输入单元格
        If mintType = 0 Then
            If vsExt.Col <> vsExt.Cols - 1 Then
                vsExt.Col = vsExt.Cols - 1
            ElseIf vsExt.Row + 1 <= vsExt.Rows - 1 Then
                vsExt.Row = vsExt.Row + 1
                vsExt.Col = vsExt.Cols - 1
            Else
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            If vsExt.Row = vsExt.Rows - 1 Then
                If vsExt.RowData(vsExt.Row) = 0 Or mblnNotAddNew Then
                    Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                Else
                    vsExt.AddItem ""
                End If
            End If
            If vsExt.Row + 1 <= vsExt.Rows - 1 Then
                vsExt.Row = vsExt.Row + 1
                vsExt.Col = 0
            End If
        ElseIf mintType = 2 Then
            If CLng(vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)) = 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            Else
                Call EnterNextCell(vsExt.Row, vsExt.Col)
            End If
        ElseIf mintType = 3 Then
            If vsExt.Col <> 1 Then
                vsExt.Col = 1
            ElseIf vsExt.Row + 1 <= vsExt.Rows - 1 Then
                vsExt.Row = vsExt.Row + 1
                vsExt.Col = 1
            Else
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
        End If
    ElseIf KeyAscii = Asc("*") Then
        If mintType = 0 Or mintType = 1 Or mintType = 4 Then
            KeyAscii = 0
            If cmd.Visible Then cmd_Click
        ElseIf mintType = 2 Then
            KeyAscii = 0
            cmd_Click '选择单味中草药或脚注
        End If
    End If
End Sub

Private Sub vsExt_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'功能：非回车确认完后编辑的处理(这里Text:=EditText,但ValidateEdit事件中还没有)
    Dim i As Long
    If Not mblnReturn Then
        If mintType = 0 And optPosition(1).Value Then
            If Col = vsExt.Cols - 1 Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            If Col = 0 Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                            
                '重新初始标本
                If mintType = 4 Then InitCombox
            End If
        ElseIf mintType = 2 Then
            If Col Mod 4 = 0 Then '中药
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
            ElseIf Col Mod 4 = 1 Then '单味用量
                If Not IsNumeric(vsExt.TextMatrix(Row, Col)) _
                    Or Val(vsExt.TextMatrix(Row, Col)) <= 0 _
                    Or Val(vsExt.TextMatrix(Row, Col)) > LONG_MAX Then
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Else
                    vsExt.Cell(flexcpData, Row, Col) = vsExt.TextMatrix(Row, Col)
                End If
            ElseIf Col Mod 4 = 3 Then '脚注
                If zlCommFun.ActualLen(vsExt.TextMatrix(Row, Col)) > 100 Then
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Else
                    vsExt.Cell(flexcpData, Row, Col) = vsExt.TextMatrix(Row, Col)
                End If
            End If
        ElseIf mintType = 3 Then
            '取消其它标本选择(单选)
            If Val(vsExt.TextMatrix(Row, 1)) <> 0 Then
                For i = vsExt.FixedRows To vsExt.Rows - 1
                    If i <> Row And Val(vsExt.TextMatrix(i, 1)) <> 0 Then
                        vsExt.TextMatrix(i, 1) = 0
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub vsExt_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'功能：输入数据确认
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, int性别 As Integer, str药品 As String
    Dim strStock As String, blnCancel As Boolean, i As Long
    Dim vPoint As POINTAPI, strSamples As String, strLike As String
    
    If mstr性别 Like "*男*" Then
        int性别 = 1
    ElseIf mstr性别 Like "*女*" Then
        int性别 = 2
    End If
    
    If KeyAscii = 13 Then
        mblnReturn = True '标记是按回车确认编辑
        KeyAscii = 0
        
        '优化
        strLike = mstrLike
        If Len(vsExt.EditText) < 2 Then strLike = ""
        
        On Error GoTo errH
        
        If mintType = 0 And optPosition(1).Value Then
            '输入检查部位
            strSQL = _
                "Select A.ID, A.编码, A.标本部位 As 检查部位" & vbNewLine & _
                "From 诊疗项目目录 A, 诊疗项目组合 B" & vbNewLine & _
                "Where A.ID = B.诊疗项目id And B.诊疗组合id = [1] And A.服务对象 In ([4], 3) And" & vbNewLine & _
                "      (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & vbNewLine & _
                "      (A.编码 Like [2] Or Upper(A.标本部位) Like [3] Or " & IIF(mint简码 = 0, "zlSpellCode", "zlWBCode") & "(A.标本部位) Like [3])" & vbNewLine & _
                "Order By B.序号"
            vPoint = GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "部位", False, "", "", False, False, True, vPoint.x, vPoint.y, vsExt.CellHeight, blnCancel, False, True, _
                mlng项目ID, UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mint服务对象)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配项目！", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            '检查重复输入
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "该检查部位已经在其它行录入。", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            Call Set部位输入(Row, rsTmp)
        ElseIf mintType = 1 Then
            '输入附加手术:这里不是单独应用,因此不限制
            strSQL = _
                " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 规模" & _
                " From 诊疗项目目录 A,诊疗项目别名 B" & _
                " Where A.ID=B.诊疗项目ID And A.类别='F' And A.ID<>[3]" & IIF(strLike = "", "", " And Rownum<=100") & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                    " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]) And B.码类=[4]" & _
                    " And A.服务对象 IN([5],3) And Nvl(A.执行频率,0) IN(0,[6]) And Nvl(A.适用性别,0) IN(0,[7])" & _
                " Order by A.编码"
            vPoint = GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "手术", False, "", "", False, False, True, vPoint.x, vPoint.y, vsExt.CellHeight, blnCancel, False, True, _
                UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mlng项目ID, mint简码 + 1, mint服务对象, IIF(mint期效 = 0, 2, 1), int性别)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配项目！", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            '检查重复输入
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "该附加手术已经在其它行录入。", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            Call Set手术输入(Row, rsTmp)
        ElseIf mintType = 2 Then
            '截取回车后,如果用Msgbox使Edit焦点丢失,则会完成编辑,但不会激活AfterEdit事件
            If Col Mod 4 = 0 Then '中药
                '中药库存,中药房未指定时,读不出库存记录
                If mlng中药房 <> 0 Then
                    strStock = _
                        "Select 药品ID,Sum(Nvl(可用数量,0)) as 库存 From 药品库存" & _
                        " Where (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期>Trunc(Sysdate))" & _
                        " And 性质=1 And 库房ID=[3] Group by 药品ID" & _
                        " Having Sum(Nvl(可用数量,0))<>0"
                End If
                
                '特殊药品权限
                str药品 = ""
                If InStr(mstrPrivs, "下达麻醉药嘱") = 0 Then
                    str药品 = str药品 & " And E.毒理分类<>'麻醉药'"
                End If
                If InStr(mstrPrivs, "下达毒性药嘱") = 0 Then
                    str药品 = str药品 & " And E.毒理分类<>'毒性药'"
                End If
                If InStr(mstrPrivs, "下达贵重药嘱") = 0 Then
                    str药品 = str药品 & " And E.价值分类 Not IN('贵重','昂贵')"
                End If
                
                '输入单味中药:这里不是单独应用,因此不限制
                strSQL = "Select A.ID,A.编码,A.名称,A.计算单位" & _
                    " From 诊疗项目目录 A,诊疗项目别名 B" & _
                    " Where A.ID=B.诊疗项目ID And A.类别='7'" & _
                    " And (A.编码 Like [1] And B.码类=[5] Or B.名称 Like [2] And B.码类=[5] Or B.简码 Like [2] And B.码类 IN([5],3))" & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) And A.服务对象 IN([4],3)" & _
                    " And Nvl(A.执行频率,0) IN(0,[6]) And Nvl(A.适用性别,0) IN(0,[7])"
                If strLike = "" And strStock <> "" Then
                    '当可以利用简码索引时(单向匹配),如果有(+)连接(药品库存),则需要Group By一下(奇怪)
                    strSQL = strSQL & " Group by A.ID,A.编码,A.名称,A.计算单位"
                End If
                If strStock = "" Then
                    strSQL = _
                        " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位,D.规格,D.产地,NULL AS 库存,E.处方职务 as 处方职务ID" & _
                        " From 药品特性 E,药品规格 C,收费项目目录 D,(" & strSQL & ") A" & _
                        " Where A.ID=E.药名ID And A.ID=C.药名ID And C.药品ID=D.ID" & _
                            " And (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 IS NULL) And D.服务对象 IN([4],3)" & _
                            IIF(strLike = "", "", " And Rownum<=100") & str药品 & _
                        " Order by A.编码"
                Else
                    strSQL = _
                        " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位,D.规格,D.产地," & _
                        " Decode(X.库存,NULL,NULL,X.库存/C.住院包装||C.住院单位) AS 库存,E.处方职务 as 处方职务ID" & _
                        " From 药品特性 E,药品规格 C,收费项目目录 D,(" & strSQL & ") A,(" & strStock & ") X" & _
                        " Where A.ID=E.药名ID And A.ID=C.药名ID And C.药品ID=D.ID And C.药品ID=X.药品ID(+)" & _
                            " And (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 IS NULL) And D.服务对象 IN([4],3)" & _
                            IIF(strLike = "", "", " And Rownum<=100") & str药品 & _
                        " Order by A.编码"
                End If
                vPoint = GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "中药", False, "", "", False, False, True, vPoint.x, vPoint.y, vsExt.CellHeight, blnCancel, False, True, _
                    UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mlng中药房, mint服务对象, mint简码 + 1, IIF(mint期效 = 0, 2, 1), int性别)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "未找到匹配项目！", vbInformation, gstrSysName
                    End If
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                
                '检查重复输入
                If ItemExist(rsTmp!ID, Row, Col) Then
                    MsgBox "该味中药在配方中已经录入。", vbInformation, gstrSysName
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                
                '处方职务检查
                If Not mbln护士站 Then
                    strSQL = CheckOneDuty(rsTmp!名称, Nvl(rsTmp!处方职务ID), UserInfo.姓名, mbln医保)
                    If strSQL <> "" Then
                        MsgBox strSQL, vbInformation, gstrSysName
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Exit Sub
                    End If
                End If
                
                '获取输入值
                vsExt.EditText = rsTmp!名称 '直接输入匹配时必要
                vsExt.TextMatrix(Row, Col) = rsTmp!名称
                vsExt.TextMatrix(Row, Col + 2) = rsTmp!单位
                vsExt.Cell(flexcpData, Row, Col + 2) = CLng(rsTmp!ID) '记录中药ID
            ElseIf Col Mod 4 = 1 Then '单量
                If Not IsNumeric(vsExt.EditText) Or Val(vsExt.EditText) <= 0 Or Val(vsExt.EditText) > LONG_MAX Then
                    MsgBox "单味用量输入错误，不是大于零的数字或输入数值过大！", vbInformation, gstrSysName
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                vsExt.TextMatrix(Row, Col) = vsExt.EditText
            ElseIf Col Mod 4 = 3 Then '脚注
                If vsExt.EditText <> "" Then
                    strSQL = "Select Rownum as ID,编码,名称,简码 From 中药煎服脚注" & _
                        " Where Upper(编码) Like [1] Or Upper(名称) Like [2] Or Upper(简码) Like [2]" & _
                        " Order by 编码"
                    vPoint = GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "脚注", False, "", "", False, False, True, vPoint.x, vPoint.y, vsExt.CellHeight, blnCancel, False, True, _
                        UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%")
                End If
                If rsTmp Is Nothing Then
                    If blnCancel Then
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Exit Sub
                    End If
                    '无匹配当作直接输入
                    If zlCommFun.ActualLen(vsExt.EditText) > 100 Then
                        MsgBox "脚注输入内容过长，最多只允许 50 个汉字或 100 个字符。", vbInformation, gstrSysName
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Exit Sub
                    End If
                    vsExt.TextMatrix(Row, Col) = vsExt.EditText
                Else
                    vsExt.EditText = rsTmp!名称 '直接输入匹配时必要
                    vsExt.TextMatrix(Row, Col) = rsTmp!名称
                End If
            End If
            vsExt.Cell(flexcpData, Row, Col) = vsExt.TextMatrix(Row, Col)
            Call EnterNextCell(Row, Col)
        ElseIf mintType = 4 Then
            '检验项目
            With Me.cbo标本
                For i = 0 To .ListCount - 1
                    strSamples = strSamples & ",'" & .List(i) & "'"
                Next
            End With
            If Len(strSamples) > 0 Then
                strSamples = Mid(strSamples, 2)
            Else
                strSamples = "''"
            End If
            strSQL = "Select A.ID,A.编码,A.名称,A.操作类型,A.标本部位" & _
                " From 诊疗项目目录 A,诊疗项目别名 C Where A.ID=C.诊疗项目ID" & _
                " And (A.编码 Like [1] Or C.名称 Like [2] Or C.简码 Like [2]) And C.码类=[3]" & _
                " And A.类别='C' And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) In (0,[6])" & _
                " And A.服务对象 IN([4],3) And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)"
            If strLike = "" Then
                '当可以利用简码索引时(单向匹配),如果有(+)连接,则需要Group By一下(奇怪)
                strSQL = strSQL & " Group by A.ID,A.编码,A.名称,A.操作类型,A.标本部位"
            End If
            If mlng项目ID > 0 Then '指定了诊疗单据
                strSQL = "Select Distinct A.ID,A.编码,A.名称,A.操作类型 as 检验类型,A.标本部位" & _
                    " From 诊疗单据应用 B,检验项目参考 D,检验报告项目 E,(" & strSQL & ") A" & _
                    " Where A.ID=B.诊疗项目ID And A.ID=E.诊疗项目id(+) And E.报告项目ID=D.项目id(+)" & _
                    " And B.应用场合+0=[4] And B.病历文件ID+0=[5]" & _
                    " And (D.标本类型 In (" & strSamples & ") Or D.标本类型 Is Null)" & _
                    " Order by A.编码"
            Else
                strSQL = "Select Distinct A.ID,A.编码,A.名称,A.操作类型 as 检验类型,B.病历文件ID,A.标本部位" & _
                    " From 诊疗单据应用 B,检验项目参考 D,检验报告项目 E,(" & strSQL & ") A" & _
                    " Where A.ID=B.诊疗项目ID(+) And A.ID=E.诊疗项目id(+) And E.报告项目ID=D.项目id(+)" & _
                    " And 0+B.应用场合(+)=[4] And (D.标本类型 In (" & strSamples & ") Or D.标本类型 Is Null)" & _
                    " Order by A.编码"
            End If
            vPoint = GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "检验项目", False, "", "", False, False, True, vPoint.x, vPoint.y, vsExt.CellHeight, blnCancel, False, True, _
                UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mint简码 + 1, mint服务对象, mlng项目ID, int性别)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配项目！", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            If rsTmp("检验类型") = "微生物" And vsExt.Rows > 2 Then
                If vsExt.RowData(2) <> 0 Or vsExt.Row > 1 Then '整个申请只能开一个微生物项目
                    MsgBox "微生物项目只能单独申请！", vbInformation, gstrSysName
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                    Exit Sub
                End If
            End If
            If mlng项目ID = 0 Then mlng项目ID = Nvl(rsTmp("病历文件ID"), 0)
            
            '检查重复输入
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "该检验项目已经录入！", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            '检查检验类型是否相同
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 And i <> Row Then
                    If Not (vsExt.TextMatrix(i, 1) = Nvl(rsTmp!检验类型) _
                        Or vsExt.TextMatrix(i, 1) = "" Or Nvl(rsTmp!检验类型) = "") Then
                        MsgBox "请输入相同检验类型的项目，已输入项目的检验类型为""" & vsExt.TextMatrix(i, 1) & """。", vbInformation, gstrSysName
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                        Exit Sub
                    End If
                End If
            Next
            
            '重新初始标本
            If Not InitCombox(rsTmp("ID"), Nvl(rsTmp("标本部位"))) Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            Call Set检验项目(Row, rsTmp)
            If rsTmp("检验类型") = "微生物" Then
                mblnNotAddNew = True
                vsExt.Rows = 2
            Else
                mblnNotAddNew = False
            End If
        End If
    Else
        If mintType = 2 Then
            '单味用量只允许输入数字
            If Col Mod 4 = 1 Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Set部位输入(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    vsExt.EditText = rsInput!检查部位 '对于输入直接匹配时有必要
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, vsExt.Cols - 1) = rsInput!检查部位
    vsExt.Cell(flexcpData, lngRow, vsExt.Cols - 1) = vsExt.TextMatrix(lngRow, vsExt.Cols - 1)
    
    '下一输入行
    If vsExt.RowData(vsExt.Rows - 1) <> 0 Then
        vsExt.AddItem ""
        vsExt.TextMatrix(vsExt.Rows - 1, 0) = vsExt.TextMatrix(vsExt.Rows - 2, 0)
    End If
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = vsExt.Cols - 1
End Sub

Private Sub Set手术输入(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    '附加手术
    vsExt.EditText = "[" & rsInput!编码 & "]" & rsInput!名称 '对于输入直接匹配时有必要
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, 0) = "[" & rsInput!编码 & "]" & rsInput!名称
    vsExt.Cell(flexcpData, lngRow, 0) = vsExt.TextMatrix(lngRow, 0)
    vsExt.TextMatrix(lngRow, 1) = Nvl(rsInput!规模)
    
    '下一输入行
    If vsExt.RowData(vsExt.Rows - 1) <> 0 And Not mblnNotAddNew Then vsExt.AddItem ""
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = 0
End Sub

Private Sub Set检验项目(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    '检验项目
    vsExt.EditText = "[" & rsInput!编码 & "]" & rsInput!名称 '对于输入直接匹配时有必要
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, 0) = "[" & rsInput!编码 & "]" & rsInput!名称
    vsExt.Cell(flexcpData, lngRow, 0) = vsExt.TextMatrix(lngRow, 0)
    vsExt.TextMatrix(lngRow, 1) = Nvl(rsInput!检验类型)
    
    '下一输入行
    If vsExt.RowData(vsExt.Rows - 1) <> 0 And Not mblnNotAddNew Then vsExt.AddItem ""
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = 0
End Sub

Private Sub vsExt_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsExt.EditSelStart = 0
    vsExt.EditSelLength = zlCommFun.ActualLen(vsExt.EditText)
End Sub

Private Sub vsExt_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'功能：限制某些列不允许编辑(该事件后于BeforeEdit,在EditText赋值之前)
    mblnReturn = False
        
    If mintType = 0 Then
        '只允许选择部位
        If optPosition(0).Value Then
            If Col <> 2 Or vsExt.RowData(Row) = 0 Then Cancel = True
        Else
            If cmd.Visible Then cmd.Visible = False '开始编辑了则隐藏按钮
            If Col <> 1 Then Cancel = True
        End If
    ElseIf mintType = 1 Or mintType = 4 Then
        '只允许编辑附加手术
        If cmd.Visible Then cmd.Visible = False '开始编辑了则隐藏按钮
        If Col <> 0 Then Cancel = True
    ElseIf mintType = 2 Then
        '必须依次输入
        If Not CellCanEdit(Row, Col) Then Cancel = True
        
        If Col Mod 4 = 1 Then
            vsExt.EditMaxLength = 8
        Else
            vsExt.EditMaxLength = 0
        End If
    ElseIf mintType = 3 Then
        '只允许选择标本
        If Col <> 1 Then
            Cancel = True
        ElseIf Val(vsExt.TextMatrix(Row, Col)) <> 0 Then
            Cancel = True '不允许取消选择(单选)
        End If
    End If
End Sub

Private Function CellCanEdit(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'功能：输入中药配方时,判断指定的单元格当前是否输入内容
'说明：在配方输入表格中,如果前一个未输入,则当前不允许输入
    '定位到上一个中药输入单元
    On Error Resume Next
    lngCol = (lngCol \ 4) * 4
    If lngCol - 4 >= vsExt.FixedCols Then
        lngCol = lngCol - 4
    Else
        If lngRow - 1 >= vsExt.FixedRows Then
            lngRow = lngRow - 1
            lngCol = vsExt.Cols - 4
        Else
            CellCanEdit = True
            Exit Function
        End If
    End If
    CellCanEdit = CLng(vsExt.Cell(flexcpData, lngRow, lngCol + 2)) <> 0
End Function

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'功能：进入下一个中药配方的输入单元格

    '当前位置未输入中药
    If CLng(vsExt.Cell(flexcpData, lngRow, (lngCol \ 4) * 4 + 2)) = 0 Then Exit Sub
    
    '单量未输入
    If lngCol Mod 4 = 1 And vsExt.TextMatrix(lngRow, lngCol) = "" Then Exit Sub
    
    If lngCol + 1 <= vsExt.Cols - 1 Then
        lngCol = lngCol + 1
    Else
        If lngRow + 1 > vsExt.Rows - 1 Then
            vsExt.AddItem "", vsExt.Rows
            Call SetSplitLine
        End If
        lngRow = lngRow + 1
        lngCol = vsExt.FixedCols
    End If
    
    vsExt.Row = lngRow: vsExt.Col = lngCol
End Sub

Private Function ItemExist(ByVal lng中药ID As Long, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'功能：判断中药配方输入表格中,指定的中药是否已经输入
    Dim i As Long, j As Long
    
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = 0 To vsExt.Cols - 1 Step 4
            If Not (lngRow = i And (lngCol \ 4) * 4 = j) Then
                If CLng(vsExt.Cell(flexcpData, i, j + 2)) = lng中药ID Then
                    ItemExist = True
                    Exit Function
                End If
            End If
        Next
    Next
End Function

Private Function Check中药禁忌(ByVal str中药IDs As String) As Boolean
'功能：检查一个配方中的中药配伍禁忌
'参数：str中药IDs="1,2,3,..."
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str慎用 As String, str禁用 As String, lng组编号 As Long
    
    On Error GoTo errH
    
    strSQL = "Select 组编号 From 诊疗互斥项目" & _
        " Where 项目ID+0 IN(" & str中药IDs & ") Group by 组编号 Having Count(*)>1"
    strSQL = "Select A.组编号,A.类型,B.名称" & _
        " From 诊疗互斥项目 A,诊疗项目目录 B" & _
        " Where A.项目ID=B.ID And A.组编号 IN(" & strSQL & ")" & _
        " And A.项目ID+0 IN(" & str中药IDs & ")" & _
        " Order by A.组编号,B.编码"
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp!组编号 <> lng组编号 Then
                If rsTmp!类型 = 1 Then
                    str慎用 = str慎用 & vbCrLf & "●"
                Else
                    str禁用 = str禁用 & vbCrLf & "●"
                End If
                lng组编号 = rsTmp!组编号
            End If
            If rsTmp!类型 = 1 Then
                str慎用 = str慎用 & "，" & rsTmp!名称
            Else
                str禁用 = str禁用 & "，" & rsTmp!名称
            End If
            rsTmp.MoveNext
        Next
        If str禁用 <> "" Then
            MsgBox "当前配方中发现下列药品互相禁用：" & Replace(str禁用, "●，", "● "), vbInformation, gstrSysName
            Exit Function
        ElseIf str慎用 <> "" Then
            If MsgBox("当前配方中发现下列药品互相慎用：" & Replace(str慎用, "●，", "● ") & vbCrLf & vbCrLf & "要继续吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    Check中药禁忌 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
