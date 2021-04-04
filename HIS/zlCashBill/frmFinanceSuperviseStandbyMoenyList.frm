VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFinanceSuperviseStandbyMoenyList 
   BorderStyle     =   0  'None
   Caption         =   "备用金列表"
   ClientHeight    =   8565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11625
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11625
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      Begin VB.CheckBox chkCancel 
         Caption         =   "含回收记录(&C)"
         Height          =   210
         Left            =   6240
         TabIndex        =   8
         Top             =   180
         Width           =   2130
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "重新过滤数据(&R)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   8655
         TabIndex        =   3
         Top             =   105
         Width           =   1605
      End
      Begin VB.ComboBox cboPerson 
         Height          =   330
         Left            =   1020
         TabIndex        =   2
         Text            =   "cboPerson"
         Top             =   120
         Width           =   2040
      End
      Begin VB.TextBox txtNO 
         Height          =   345
         Left            =   3600
         TabIndex        =   1
         Top             =   113
         Width           =   2415
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2355
         TabIndex        =   6
         Top             =   150
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblPerson 
         AutoSize        =   -1  'True
         Caption         =   "收费员"
         Height          =   210
         Left            =   315
         TabIndex        =   5
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lblNo 
         AutoSize        =   -1  'True
         Caption         =   "NO"
         Height          =   210
         Left            =   3330
         TabIndex        =   4
         Top             =   165
         Width           =   210
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   1800
      Left            =   360
      TabIndex        =   7
      Top             =   1110
      Width           =   8625
      _cx             =   15214
      _cy             =   3175
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFinanceSuperviseStandbyMoenyList.frx":0000
      ScrollTrack     =   -1  'True
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
Attribute VB_Name = "frmFinanceSuperviseStandbyMoenyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsPerson As ADODB.Recordset
Private mlngModule As Long, mstrPrivs As String
Private mblnDrop As Boolean

Public Sub zlInitVar(ByVal lngModule As Long, ByVal strPrivs As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关变量
    '入参:lngModule-模块号
    '       strPrivs-权限串
    '编制:刘兴洪
    '日期:2013-09-09 14:41:46
    '说明:加载窗体后,立即调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
End Sub
Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格控件
    '编制:刘兴洪
    '日期:2013-09-11 17:34:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strHead As String, varData As Variant
    strHead = "ID,单据号,金额,备注,领用人,收回人,收回时间,登记人,登记时间"
    varData = Split(strHead, ",")
    With vsList
        .Clear
        .Rows = 2: .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = varData(i)
            .ColKey(i) = varData(i)
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "单据号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*金额" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Name, "备用金信息列表", False
    End With
End Sub
Private Function LoadData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载历史收款数据
    '返回:数据加载成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-11 17:08:50
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strPerson As String, strWhere As String, i As Long, blnDel As Boolean
    
    
    On Error GoTo errHandle
     strPerson = zlStr.NeedName(cboPerson.Text)
    If txtNO.Text <> "" Then
        strWhere = strWhere & " And A.NO=[1]"
    Else
        strWhere = strWhere & " And A.收款员=[2]"
    End If
    If chkCancel.Value <> 1 Then
        strWhere = strWhere & " And A.收回时间 is null "
    End If
    strSQL = "" & _
    "   Select A.ID, A.NO As 单据号, LTrim(To_Char(A.金额, '99999999990.00')) As 金额, A.备注, " & _
    "        A.收款员 As 领用人,  " & _
    "        A.收回人, to_char(A.收回时间,'yyyy-mm-dd hh24:mi:ss') as 收回时间, " & _
    "        A.登记人,  to_char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') as 登记时间 " & _
    "   From 人员暂存记录 A" & _
    "   Where MOD(A.记录性质,10) = 1 " & strWhere & _
    "   Order by 登记时间 Desc,NO Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(Trim(txtNO.Text)), strPerson)
    
    With vsList
        .Clear 1: .Rows = 2
        .FixedRows = 1
        If rsTemp.RecordCount <> 0 Then
            Set .DataSource = rsTemp
        End If
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "单据号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*金额" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        For i = 1 To .Rows - 1
            blnDel = Trim(.TextMatrix(i, .ColIndex("收回时间"))) <> ""
            If blnDel Then
                '作废记录，用红色字体
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            End If
        Next
        .Row = 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Name, "备用金信息列表", False
        If .Enabled And .Visible Then .SetFocus
    End With
    LoadData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Sub cboPerson_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim rsTemp As ADODB.Recordset
    If KeyAscii <> 13 Then Exit Sub
    
    If cboPerson.Locked Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    strText = UCase(cboPerson.Text)
    If cboPerson.ListIndex <> -1 Then
        '弹出列表时,又在文本框输入了内容
        If strText <> UCase(cboPerson.List(cboPerson.ListIndex)) Then Call zlcontrol.CboSetIndex(cboPerson.hWnd, -1)
    End If
    If strText = "" Then cboPerson.ListIndex = -1: Exit Sub
    '69061,刘尔旋,2013-12-30,触发刷新列表的方式调整
    If cboPerson.ListIndex >= 0 Then
        Call cmdRefresh_Click
        Exit Sub
    End If
    intIdx = -1
    '先复制记录集
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsPerson)
    strCompents = Replace(gstrLike, "%", "*") & strText & "*"
    If IsNumeric(strText) Then
        intInputType = 0 '0-输入的是全数字
    ElseIf zlCommFun.IsCharAlpha(strText) Then
        intInputType = 1 '1-输入的是全字母
    Else
        intInputType = 2 '2-其他
    End If
    mrsPerson.Filter = 0: iCount = 0
    With mrsPerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not mrsPerson.EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编号) = strText Then strResult = Nvl(!姓名): iCount = 0: Exit Do
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编号)) = Val(strText) Then
                    If iCount = 0 Then strResult = Nvl(!姓名)
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Val(mrsPerson!编号) Like strText & "*" Then
                    If CheckPersonExists(Nvl(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrsPerson, rsTemp)
                 End If
                 
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strText Then
                    If iCount = 0 Then strResult = Nvl(!姓名)   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If CheckPersonExists(Nvl(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrsPerson, rsTemp)
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编号) = strText Or Trim(!简码) = strText Or Trim(!姓名) = strText Then
                    If iCount = 0 Then strResult = Nvl(!姓名)   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If Trim(!编号) Like strText & "*" Or Trim(Nvl(!简码)) Like strCompents Or Trim(Nvl(!姓名)) Like strCompents Then
                    If CheckPersonExists(Nvl(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrsPerson, rsTemp)
                End If
            End Select
            mrsPerson.MoveNext
        Loop
    End With
    
    If iCount > 1 Then strResult = ""
    If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!姓名)
    '直接定位
    If strResult <> "" Then
        rsTemp.Close: Set rsTemp = Nothing
        If CheckPersonExists(strResult, True) Then Call cmdRefresh_Click
        Exit Sub
    End If
     If rsTemp.RecordCount = 0 Then
        '未找到
        rsTemp.Close: Set rsTemp = Nothing
        KeyAscii = 0: zlcontrol.TxtSelAll cboPerson: Exit Sub
     End If
     
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = "编号"
    Case 1 '输入全拼音
        rsTemp.Sort = "简码"
    Case Else
        '根据选择来定
        rsTemp.Sort = "编号"
    End Select
    '弹出选择器
    Dim rsReturn As ADODB.Recordset
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, cboPerson, rsTemp, True, "", "", rsReturn) Then
        If cboPerson.Enabled Then cboPerson.SetFocus
        If Not rsReturn Is Nothing Then
            If rsReturn.RecordCount <> 0 Then
                '进行定位
                If CheckPersonExists(Nvl(rsReturn!姓名), True) Then
                    'zlCommFun.PressKey vbKeyTab
                End If
            End If
        End If
    End If
    rsTemp.Close: Set rsTemp = Nothing
End Sub

Private Sub cboPerson_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub cboPerson_Validate(Cancel As Boolean)
    If cboPerson.Text <> "" Then
        If cbo.FindIndex(cboPerson, zlStr.NeedName(cboPerson.Text), True) = -1 Then cboPerson.ListIndex = -1: cboPerson.Text = ""
    End If
    If cboPerson.Text = "" Then Call cboPerson_KeyPress(vbKeyReturn)
    '有数据，必须输入
    If cboPerson.ListIndex = -1 And cboPerson.ListCount <> 0 Then Cancel = True
End Sub
Private Sub cboPerson_GotFocus()
    Call zlCommFun.OpenIme(True)
    Call zlcontrol.TxtSelAll(cboPerson)
End Sub
 

Private Sub cmdRefresh_Click()
    txtNO.Text = ""
    Call LoadData
End Sub
Private Sub Form_Load()
    Call InitGrid
    Call LoadPerson
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsList
        .Left = ScaleLeft + 50
        .Top = picTop.Top + picTop.Height + 50
        .Height = ScaleHeight - .Top + 50
        .Width = ScaleWidth - .Left * 2
    End With
End Sub
Private Function LoadPerson() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收费员
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-23 11:59:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    gstrSQL = "" & _
    "   Select distinct A.ID,A.编号,A.姓名,A.简码  " & _
    "   From 人员表 A,人员性质说明 B " & _
    "   Where A.id = B.人员ID  " & _
    "               And B.人员性质 In ('门诊挂号员','门诊收费员','预交收款员','住院结帐员','入院登记员','发卡登记人')  " & _
    "               And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
    "               And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
    "   Order By 编号"
    Set mrsPerson = zlDatabase.OpenSQLRecord(gstrSQL, "检查当前操作员是否为相应性质人员")
    With cboPerson
        Do While Not mrsPerson.EOF
            .AddItem Nvl(mrsPerson!编号) & "-" & Nvl(mrsPerson!姓名)
            .ItemData(.NewIndex) = Val(Nvl(mrsPerson!ID))
            If .ListIndex < 0 Then .ListIndex = .NewIndex
            mrsPerson.MoveNext
        Loop
        If .ListCount <> 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    LoadPerson = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function isCheckPersonExists(ByVal str姓名 As String, _
    Optional blnLocateItem As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查姓名是否在收费员下拉列表中
    '入参:str姓名-姓名
    '     blnLocateItem:是否直接定位
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-23 14:34:47
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cboPerson.ListCount - 1
        If zlStr.NeedName(cboPerson.List(i)) = str姓名 Then
            If blnLocateItem Then cboPerson.ListIndex = i
            isCheckPersonExists = True
            Exit Function
        End If
    Next
End Function
Private Sub txtNO_GotFocus()
    zlcontrol.TxtSelAll txtNO
    zlCommFun.OpenIme False

End Sub
Private Sub txtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Or Trim(txtNO.Text) = "" Then Exit Sub
    txtNO.Text = GetFullNO(Trim(txtNO.Text), 141)
    Call LoadData
End Sub

Public Sub zlPrint(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输出列表信息
    '入参:bytMode=1-打印,2-预览,3-输出到Excel
    '编制:刘兴洪
    '日期:2013-09-13 10:23:30
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As New zlPrint1Grd, objRow As New zlTabAppRow
    
    Err = 0: On Error GoTo ErrHand:
    
    '输出收款信息
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstr单位名称 & "备用金发放清册"
    Set objRow = New zlTabAppRow
    If txtNO.Text <> "" Then
        objRow.Add "单据号：" & txtNO.Text
    Else
        objRow.Add "收费员：" & cboPerson.Text
    End If
    If chkCancel.Value = 1 Then
        objRow.Add "含作废的备用金"
    End If
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    Set objPrint.Body = vsList
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
Public Sub RePrintBill()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重打收款收据
    '编制:刘兴洪
    '日期:2013-09-13 16:00:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String
    If Not (zlStr.IsHavePrivs(mstrPrivs, "备用金领用单打印") And zlStr.IsHavePrivs(mstrPrivs, "重打备用金领用单")) Then Exit Sub
    With vsList
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("单据号")))
    End With
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1500_1", Me, "NO=" & strNO, 2)
End Sub
Public Sub zlRefresh()
    '重新进行数据刷新
    Call cmdRefresh_Click
End Sub
Public Sub CallCustomRpt(ByVal frmMain As Object, ByVal lngSys As Long, ByVal strRptCode As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用自定义报表
    '入参:lngSys-系统号
    '        strRptCode-报表编号
    '编制:刘兴洪
    '日期:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngDeptID As Long
    Dim lngID As Long, dtStartDate As Date, dtEndDate As Date, blnDel As Boolean
    Dim strNO As String
    With vsList
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("单据号")))
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("收回时间"))) <> ""
    End With
    Call ReportOpen(gcnOracle, lngSys, strRptCode, frmMain, _
        "NO=" & strNO, _
        "ID=" & lngID, _
        "作废标志=" & IIf(blnDel, 1, 0))
End Sub

Public Function zlPayOnWorkMoney(ByVal frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发放上岗备用金
    '返回:发放成功返回true,否则返回False
    '编制:刘尔旋
    '日期:2013-12-4
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str领用人 As String
    Dim frmNew As New frmFinanceSuperviseStandbyMoneyEdit
    Dim blnReturn As Boolean
    On Error GoTo errHandle
    If cboPerson.ListIndex > 0 Then
        If cboPerson.ItemData(cboPerson.ListIndex) <> 0 Then
            str领用人 = zlStr.NeedName(cboPerson.Text)
        End If
    End If
    blnReturn = frmNew.EditCard(frmMain, EM_ED_上岗, mlngModule, mstrPrivs, str领用人, 0)
    If Not frmNew Is Nothing Then Unload frmNew
    Set frmNew = Nothing
    If blnReturn Then zlRefresh
    zlPayOnWorkMoney = blnReturn
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPayStandbyMoney(ByVal frmMain As Object) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发放备用金
    '返回:发放成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-12 16:45:53
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str领用人 As String
    Dim frmNew As New frmFinanceSuperviseStandbyMoneyEdit
    Dim blnReturn As Boolean
    On Error GoTo errHandle
    If cboPerson.ListIndex > 0 Then
        If cboPerson.ItemData(cboPerson.ListIndex) <> 0 Then
            str领用人 = zlStr.NeedName(cboPerson.Text)
        End If
    End If
    blnReturn = frmNew.EditCard(frmMain, EM_ED_增加, mlngModule, mstrPrivs, str领用人, 0)
    If Not frmNew Is Nothing Then Unload frmNew
    Set frmNew = Nothing
    If blnReturn Then zlRefresh
    zlPayStandbyMoney = blnReturn
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckPersonExists(ByVal str姓名 As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查姓名是否在你收费员下拉列表中.
    '入参:str姓名-姓名
    '     blnLocateItem:是否直接定位
    '出参:
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cboPerson.ListCount - 1
        If zlStr.NeedName(cboPerson.List(i)) = str姓名 Then
            If blnLocateItem Then cboPerson.ListIndex = i
            CheckPersonExists = True
            Exit Function
        End If
    Next
End Function

Private Sub vsList_DblClick()
    Dim lngID As Long
    Dim frmNew As frmFinanceSuperviseStandbyMoneyEdit
    With vsList
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Sub
        lngID = .TextMatrix(.Row, .ColIndex("ID"))
        If lngID = 0 Then Exit Sub
    End With
    On Error GoTo errHandle
    Set frmNew = New frmFinanceSuperviseStandbyMoneyEdit
    Call frmNew.EditCard(Me, EM_ED_查看, mlngModule, mstrPrivs, "", lngID)
    If Not frmNew Is Nothing Then Unload frmNew
    Set frmNew = Nothing
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_GotFocus()
    Call zl_VsGridGotFocus(vsList)
End Sub
Private Sub vsList_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsList, GRD_LOSTFOCUS_COLORSEL)
    vsList.Tag = "0"
End Sub
Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Name, "备用金信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsList, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Name, "备用金信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Public Function CancelStandbyMoney() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废备用金
    '返回:作废成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-12 18:00:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim strSQL As String, lngID As Long, strNO As String
    Dim strTime As String
    On Error GoTo errHandle
    With vsList
        If .Row < 0 Or .Row > .Rows - 1 Then
            Exit Function
        End If
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        strNO = Trim(.TextMatrix(.Row, .ColIndex("单据号")))
        If lngID = 0 Then Exit Function
    End With
    If MsgBox("你是否真的要收回单号为" & strNO & "的备用金吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    strTime = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    ' Zl_人员暂存记录_Cancel
    strSQL = "Zl_人员暂存记录_Cancel("
    '  Id_In       In 人员暂存记录.Id%Type,
    strSQL = strSQL & "" & lngID & ","
    '  收回人_In   In 人员暂存记录.收回人%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  收回时间_In In 人员暂存记录.收回时间%Type
    strSQL = strSQL & "to_date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'))"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    With vsList
        If chkCancel.Value = 1 Then
            .TextMatrix(.Row, .ColIndex("收回人")) = UserInfo.姓名
            .TextMatrix(.Row, .ColIndex("收回时间")) = strTime
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbRed
        Else
            lngRow = .Row
            If (.Row < .Rows - 1 Or .Row >= 1) And .Rows - 1 > 1 Then
                If .Row = .Rows - 1 Then
                    .Row = lngRow - 1
                Else
                    .Row = lngRow + 1
                End If
                If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                .RemoveItem lngRow
            ElseIf .Row = 1 Then
                .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
            Else
                Call zlRefresh
            End If
        End If
    End With
    If vsList.Enabled And vsList.Visible Then vsList.SetFocus
    Call vsList_GotFocus
    CancelStandbyMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Property Get IsAllowCancel() As Boolean
    '允许是否回收
    Dim lngID As Long
    With vsList
        If .Row < 0 Or .Row > .Rows - 1 Then IsAllowCancel = False: Exit Property
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then IsAllowCancel = False: Exit Property
        IsAllowCancel = Trim(.TextMatrix(.Row, .ColIndex("收回时间"))) = ""
    End With
End Property
 
