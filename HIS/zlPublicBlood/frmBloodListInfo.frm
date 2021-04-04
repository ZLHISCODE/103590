VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBloodListInfo 
   BorderStyle     =   0  'None
   Caption         =   "血液列表信息"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraExecUD 
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   90
      MousePointer    =   7  'Size N S
      TabIndex        =   2
      Top             =   1185
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   1290
      Left            =   75
      TabIndex        =   0
      Top             =   1335
      Width           =   7125
      _cx             =   12568
      _cy             =   2275
      Appearance      =   2
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
      BackColorSel    =   16444122
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
   Begin RichTextLib.RichTextBox rtfOther 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmBloodListInfo.frx":0000
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
End
Attribute VB_Name = "frmBloodListInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng医嘱ID As Long
Private mlngFontSize As Long
Private mblnMoved As Boolean
Private mbln用血 As Boolean
Private mclsVsf As clsVsf
Private mblnFistRefresh As Boolean
Private mblnShowInfo As Boolean

Public Function zlRefresh(ByVal lng医嘱ID As Long, Optional ByVal lngFontSize As Long = 9, Optional ByVal blnMoved As Boolean = False) As Boolean
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim blnShowBlood As Boolean
    
    'SQL语句相关变量
    Dim strWhere As String
    On Error GoTo ErrHand
    
    mlng医嘱ID = lng医嘱ID
    mlngFontSize = lngFontSize
    mblnMoved = blnMoved
    
    If ShowOtherAppend(blnShowBlood) = False Then Exit Function
    If blnShowBlood = True Then
        If mbln用血 = True Then
            strWhere = " And a.id=f.收发ID and f.配发ID=b.id "
        Else
            strWhere = " And a.id=f.收发ID(+) And a.配发ID=b.id "
        End If
        mblnShowInfo = True
        strSQL = _
            " Select a.Id, a.血液id, a.Abo,a.Rh, To_Char(a.效期, 'YYYY-MM-DD hh24:mi') 血液效期, a.颜色 血液颜色, a.外观 血袋外观, a.配血人," & vbNewLine & _
            "       To_Char(a.配血日期, 'YYYY-MM-DD hh24:mi') 配血时间, a.核对人 审核人, To_Char(a.核对日期, 'YYYY-MM-DD hh24:mi') 审核时间," & vbNewLine & _
            "       Nvl(a.发血状态, 0) 发血状态,nvl(f.执行状态,0) 执行状态," & vbNewLine & _
            "       Decode(Nvl(f.执行状态, 0),0,Decode(Nvl(f.接收状态, 0), 0, Decode(Nvl(a.发血状态, 0),0,'新配',1,'已审核',2,'发出',9,'发出',3,Decode(a.审核人, Null, '退血', '拒发'),''), 2, '拒绝接收', '已接收'),1,'正在执行',2,'完成执行',3,'停止执行') 血液状态," & vbNewLine & _
            "       c.名称 库房, d.签名人, To_Char(d.签名时间, 'YYYY-MM-DD hh24:mi') 签名时间, a.血袋编号, a.实际数量 As 数量, e.名称 As 血液名称, e.规格 血液规格," & vbNewLine & _
            "       (Select f_List2str(Cast(Collect(g.名称) As t_Strlist))" & vbNewLine & _
            "         From 诊疗项目目录 g, 血液配血方法 f" & vbNewLine & _
            "         Where f.配血方法id = g.Id(+) And f.收发id = a.Id) 配血方法," & vbNewLine & _
            "       (Select Max(f.配血结论) From 诊疗项目目录 g, 血液配血方法 f Where f.配血方法id = g.Id(+) And f.收发id = a.Id) 配血结论,a.发血人,A.发血日期,a.取血人,a.摘要 配血摘要" & vbNewLine & _
            " From 部门表 c, 血库签名 d, 收费项目目录 e, 血液品种 k, 血液规格 l, 血液收发记录 a,血液发送记录 f, 血液配血记录 b" & vbNewLine & _
            " Where c.Id = a.库房id And a.配血签名id = d.Id(+) And d.性质(+) = 3 And a.血液id = e.Id And k.品种id = l.品种id And l.规格id = a.血液id And" & vbNewLine & _
            "      Nvl(a.填写数量, 0) <> 0 And a.单据 = 6 And Mod(a.记录状态, 3) = 1 " & strWhere & " And b.申请id = [1]" & vbNewLine & _
            " Order By a.配血日期, a.序号"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "已发血液信息提取", lng医嘱ID)
        Call mclsVsf.LoadGrid(rsTemp, "", True)
    Else
        mblnShowInfo = False
    End If
    Call SetFontSize(mlngFontSize)
    mblnFistRefresh = False
    Call Form_Resize
    zlRefresh = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowOtherAppend(blnShowBlood As Boolean) As Boolean
'功能：显示指定行医嘱的审核信息
'说明：只检查审核状态通过和未通过的医嘱
'返回：是否需要读取和显示血液列表
    Dim strSQL As String
    Dim int类型 As Integer
    Dim rsTmp As ADODB.Recordset
    Dim str操作员 As String, str时间 As String, str状态 As String, str操作说明 As String
    Dim str未用原因 As String
    
    Dim int执行标记 As Integer, int审核状态 As Integer
    Dim str检查方法 As String, bln用血 As Boolean
    Dim arrCode, arrItem
    Dim i As Integer, lngIdx As Long
    On Error GoTo errH
    
    mbln用血 = False
    blnShowBlood = False
    rtfOther.Text = "": rtfOther.SelStart = 0
    '提取医嘱相关数据
    strSQL = _
        " Select b.审核状态, b.检查方法,b.执行标记, c.操作类型, c.执行分类" & vbNewLine & _
        " From 诊疗项目目录 c, 病人医嘱记录 a, 病人医嘱记录 b" & vbNewLine & _
        " Where c.Id = a.诊疗项目id And a.相关id = b.Id And a.诊疗类别 = 'E' And b.Id = [1] And b.诊疗类别 = 'K'"
    If mblnMoved = True Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID)
    '血液医嘱肯定查得到数据，查不到数据则退出
    If rsTmp.EOF Then
        Exit Function
    End If
    int审核状态 = Val("" & rsTmp!审核状态)
    str检查方法 = "" & rsTmp!检查方法
    If str检查方法 = "" Then
        If Val("" & rsTmp!操作类型) = "8" And Val("" & rsTmp!执行分类) = 1 Then
            bln用血 = True
        End If
    Else
        bln用血 = Val(str检查方法) = 1
    End If
    mbln用血 = bln用血
    str操作员 = "审核人：": str时间 = "审核时间：": str状态 = ""
    
    If int执行标记 = -1 Then '读取标记未用的原因
        strSQL = "Select 操作人员,操作时间,操作说明 From 病人医嘱状态 Where 医嘱id = [1] And 操作类型 = [2]"
        If mblnMoved = True Then
            strSQL = Replace(strSQL, "病人医嘱状态", "H病人医嘱状态")
        End If
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, 17)
        If Not rsTmp.EOF Then
            str未用原因 = "未用原因：" & rsTmp!操作说明
            str未用原因 = str未用原因 & "(操作员：" & rsTmp!操作人员 & "  操作时间：" & Format(rsTmp!操作时间 & "", "YYYY-MM-DD HH:MM:SS") & ")"
        End If
    End If
    
    '以前的用血医嘱，无法对应到血液记录，什么都不显示
    If bln用血 = True And str检查方法 = "" Then
        Exit Function
    End If
    '目前输血处理的医嘱状态只有这几种，后面如果新增需要调整这部分内容
    Select Case int审核状态
        Case 2 '审核完成
            If bln用血 = False Then
                int类型 = 15 '血库审核通过
                str状态 = "完成配血"
                str操作员 = "配血完成人："
                str时间 = "配血完成时间："
            Else
                int类型 = 15 '血库审核通过
                str状态 = "完成发血"
                str操作员 = "发血操作人："
                str时间 = "发血操作时间："
            End If
            blnShowBlood = True
        Case 3 '(启用输血分级管理审核未通过，接收后审核未通过)
            If bln用血 = True Then
                '用血医嘱不用分级审核，审核状态=3说明是拒绝发血
                int类型 = 16
                str状态 = "拒绝发血"
                str操作员 = "拒绝发血人："
                str时间 = "拒绝发血时间："
                str操作说明 = "拒绝发血原因："
            Else
                '备血医嘱需要检查是输血审核拒绝还是拒绝配血
                strSQL = "Select 1 From 血液配血记录 Where 申请ID=[1] and 记录状态=[2]"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, 3)
                If rsTmp.RecordCount > 0 Then
                    int类型 = 16 '拒绝配血
                    str状态 = "拒绝配血"
                    str操作员 = "拒绝配血人："
                    str时间 = "拒绝配血时间："
                    str操作说明 = "拒绝配血原因："
                Else
                    int类型 = 12 '审核未通过
                    str状态 = "审核未通过"
                    str操作员 = "审核人："
                    str时间 = "审核时间："
                    str操作说明 = "审核未通过原因："
                End If
            End If
        Case 4
            str状态 = "等待配血"
            int类型 = 11 '正常医嘱，如果启用输血分机管理，则代表审核通过
        Case 5
            int类型 = 14
            str状态 = "正在配血"
            str操作员 = "配血接收人："
            str时间 = "配血接收时间："
            blnShowBlood = True
        Case 6 '输血科接收后停止配血
            int类型 = 17
            str状态 = "停止配血"
            str操作员 = "停止配血人："
            str时间 = "停止配血时间："
            str操作说明 = "停止配血原因："
        Case 1
            If bln用血 = True Then
                str状态 = "用血医嘱待核对"
                blnShowBlood = True
            End If
        Case Else
            If bln用血 = True Then
                str状态 = "等待发血"
            End If
    End Select
    rtfOther.Text = ""
    strSQL = "Select 操作人员,操作时间,操作说明 From 病人医嘱状态 Where 医嘱id = [1] And 操作类型 = [2] order by 操作时间"
    If mblnMoved = True Then
        strSQL = Replace(strSQL, "病人医嘱状态", "H病人医嘱状态")
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, int类型)
    
    strSQL = ""
    arrCode = Array()
    arrItem = Array()
    With rtfOther
        If Not rsTmp.EOF Then
            ReDim Preserve arrCode(UBound(arrCode) + 1)
            arrCode(UBound(arrCode)) = "状态：" & str状态
            ReDim Preserve arrItem(UBound(arrItem) + 1)
            arrItem(UBound(arrItem)) = "状态："
            
            Do While Not rsTmp.EOF
                ReDim Preserve arrCode(UBound(arrCode) + 1)
                arrCode(UBound(arrCode)) = str操作员 & rsTmp!操作人员
                ReDim Preserve arrCode(UBound(arrCode) + 1)
                arrCode(UBound(arrCode)) = str时间 & Format(rsTmp!操作时间 & "", "YYYY-MM-DD HH:MM:SS")
                If str操作说明 <> "" Then
                    ReDim Preserve arrCode(UBound(arrCode) + 1)
                    arrCode(UBound(arrCode)) = str操作说明 & rsTmp!操作说明
                End If
                ReDim Preserve arrItem(UBound(arrItem) + 1)
                arrItem(UBound(arrItem)) = str操作员
                ReDim Preserve arrItem(UBound(arrItem) + 1)
                arrItem(UBound(arrItem)) = str时间
                If str操作说明 <> "" Then
                    ReDim Preserve arrItem(UBound(arrItem) + 1)
                    arrItem(UBound(arrItem)) = str操作说明
                End If
                rsTmp.MoveNext
            Loop
            If str未用原因 <> "" Then
                ReDim Preserve arrCode(UBound(arrCode) + 1)
                arrCode(UBound(arrCode)) = str未用原因
                ReDim Preserve arrItem(UBound(arrItem) + 1)
                arrItem(UBound(arrItem)) = "未用原因："
            End If
        ElseIf str状态 <> "" Then
            strSQL = "状态：" & str状态
            ReDim Preserve arrCode(UBound(arrCode) + 1)
            arrCode(UBound(arrCode)) = strSQL
            ReDim Preserve arrItem(UBound(arrItem) + 1)
            arrItem(UBound(arrItem)) = "状态："
        End If
        .SelStart = 0
        For i = 0 To UBound(arrCode)
            .SelBold = False
            .SelText = IIf(.Text = "", "", vbCrLf) & CStr(arrCode(i))
            lngIdx = .Find(CStr(arrItem(i)), , , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then
                .SelStart = lngIdx
                .SelLength = Len(CStr(arrItem(i)))
                .SelBold = True
                .SelIndent = 100
            End If
            .SelStart = Len(.Text)
        Next i
        If UBound(arrItem) >= 0 Then
            lngIdx = .Find(CStr(arrItem(0)), 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(CStr(arrItem(0)))
        End If
    End With
    ShowOtherAppend = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Sub SetFontSize(ByVal lngFontSize As Long)
    With rtfOther
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelFontSize = lngFontSize
        .SelLength = 0
    End With
    Call gobjComlib.zlControl.VSFSetFontSize(vsList, lngFontSize)
    '首次刷新恢复，避免重新刷新恢复
    If mblnFistRefresh = True Then
        Call gobjComlib.RestoreWinState(Me, "zlPublicBlood")
    End If
End Sub

Private Sub Form_Load()
    mblnFistRefresh = True
    mblnShowInfo = False
    Set mclsVsf = New clsVsf
    Call InitTable
End Sub

Private Sub InitTable()
'表格初始化
    With mclsVsf
        Call .Initialize(Me.Controls, vsList, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False, , , True)
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)  '收发ID
        Call .AppendColumn("状态", 810, flexAlignLeftCenter, flexDTString, , "血液状态") '接收执行状态
        Call .AppendColumn("血液名称", 1800, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("规格", 810, flexAlignLeftCenter, flexDTString, , "血液规格")
        Call .AppendColumn("ABO", 810, flexAlignLeftCenter, flexDTString, , "ABO", True)
        Call .AppendColumn("Rh(D)", 600, flexAlignLeftCenter, flexDTString, , "RH", True)
        Call .AppendColumn("血袋编号", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("效期", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", "血液效期", True)
        Call .AppendColumn("数量", 500, flexAlignRightCenter, flexDTDecimal, , , , , , False)
        
        
        '血液配发信息
        Call .AppendColumn("配血人", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("配血时间", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        Call .AppendColumn("审核人", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("审核时间", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        Call .AppendColumn("配血方法", 1500, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("配血结论", 1500, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("配血摘要", 2000, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("发血人", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("取血人", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("发血时间", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", "发血日期")
        
        
        '隐藏列
        Call .AppendColumn("血液ID", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("发血状态", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("执行状态", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
            
        .AppendRows = False
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With rtfOther
        .Top = 0
        .Left = 0
        .Width = Me.Width
        .Height = IIf(mblnShowInfo = True, Me.Height - vsList.Height - fraExecUD.Height, Me.Height)
    End With
    With fraExecUD
        .Left = 0
        .Top = rtfOther.Top + rtfOther.Height
        .Visible = mblnShowInfo
    End With
    With vsList
        .Left = 0
        .Top = fraExecUD.Top + fraExecUD.Height
        .Width = Me.Width
        .Visible = mblnShowInfo
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call gobjComlib.SaveWinState(Me, "zlPublicBlood")
    Set mclsVsf = Nothing
End Sub

Private Sub fraExecUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If rtfOther.Height + Y < 700 Or vsList.Height - Y < 700 Then Exit Sub
        fraExecUD.Top = fraExecUD.Top + Y
        rtfOther.Height = rtfOther.Height + Y
        vsList.Top = vsList.Top + Y
        vsList.Height = vsList.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub vsList_DblClick()
    '输血执行记录查看
    If Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("ID"))) < 0 Then Exit Sub
    If Not (Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("执行状态"))) > 0) Then Exit Sub
    Call frmBloodExecEdit.ViewExecution(Me, Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("ID"))))
End Sub
