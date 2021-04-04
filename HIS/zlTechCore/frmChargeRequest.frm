VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmChargeRequest 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VSFlex8Ctl.VSFlexGrid vsMoney 
      Height          =   1290
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   6000
      _cx             =   10583
      _cy             =   2275
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
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
      GridLinesFixed  =   2
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
      Begin VB.Line lnX0 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line lnY0 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDetail 
      Height          =   1485
      Left            =   240
      TabIndex        =   1
      Top             =   2985
      Width           =   5625
      _cx             =   9922
      _cy             =   2619
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483644
      GridColorFixed  =   -2147483644
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
      GridLinesFixed  =   2
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
      Begin VB.Line lnX1 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line lnY1 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
   End
   Begin VB.Image imgX 
      Height          =   45
      Left            =   2685
      MousePointer    =   7  'Size N S
      Top             =   1470
      Width           =   5085
   End
End
Attribute VB_Name = "frmChargeRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COLOR_LOST = &HFFEBD7
Private Const COLOR_FOCUS = &HFFCC99

Private mrsPrice As ADODB.Recordset '未计费医嘱的主费用

Private mfrmParent As Form
Private pgbLoad As Object
Private AdviceID As Long
Private lngSendNO As Long
Private iPatientType As Integer
Private lngPatientID As Long
Private lngPatientDept As Long
Private lngPageId As Long
Private strCheckNo As String
Private str费别 As String
Private int记录性质 As Integer
Private int执行状态 As Integer

Private lng开单科室ID As Long
Private mstrPrivs As String

Private mSysName As String
Private mstrSys As String

Private msgl体检折扣 As Single
Private mblnDataMoved As Boolean
Private mblnChargeDataMoved As Boolean
Public mblnCash As Boolean  '是已经收费，只要存在记帐记录或一条为收费的记录，则认为未收费；


Public Property Get CashState() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    CashState = mblnCash
End Property

Public Sub zlRefresh(ByVal objParent As Object, ByVal lngAdviceID As Long, ByVal SendNO As Long, Optional ByVal strPrivs As String = "", Optional ByVal strClass As String = "检验", Optional ByVal strSys As String)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:  lngAdviceID         主医嘱id
    '------------------------------------------------------------------------------------------------------------------
    
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
        
    '初始化处理
    Set mrsPrice = New ADODB.Recordset
    mrsPrice.Fields.Append "医嘱ID", adBigInt
    mrsPrice.Fields.Append "开嘱科室ID", adBigInt
    mrsPrice.Fields.Append "类别", adVarChar, 1
    mrsPrice.Fields.Append "收费细目ID", adBigInt
    mrsPrice.Fields.Append "计算单位", adVarChar, 100, adFldIsNullable
    mrsPrice.Fields.Append "附加手术", adInteger
    mrsPrice.Fields.Append "执行科室", adInteger
    mrsPrice.Fields.Append "收入项目ID", adBigInt
    mrsPrice.Fields.Append "收据费目", adVarChar, 20, adFldIsNullable
    mrsPrice.Fields.Append "数量", adDouble
    mrsPrice.Fields.Append "单价", adDouble
    mrsPrice.Fields.Append "应收", adCurrency
    mrsPrice.Fields.Append "实收", adCurrency
    mrsPrice.Fields.Append "发送单据", adVarChar, 30
    mrsPrice.Fields.Append "发送号", adVarChar, 30
    
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
    
       
    iPatientType = 1
    lngPatientID = 0
    lngPageId = 0
    strCheckNo = ""
    lngPatientDept = 0
'    int计费状态 = 0
    str费别 = ""
    int记录性质 = 1
    mstrPrivs = strPrivs
    int执行状态 = 0
'    strNo = ""
    lng开单科室ID = 0
            
    '接口参数处理
    mSysName = strClass
    mstrSys = strSys
    mstrPrivs = strPrivs
    AdviceID = lngAdviceID
    lngSendNO = SendNO
    Set mfrmParent = objParent
        
    On Error GoTo DBError
    
    '数据转储处理
    mblnDataMoved = False
    strSQL = "Select b.开嘱时间 From 病人医嘱记录 b Where b.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    If rsTmp.BOF = False Then
        mblnDataMoved = False
        mblnChargeDataMoved = zlDatabase.DateMoved(Format(rsTmp("开嘱时间").Value, "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption)
    Else
        mblnDataMoved = True
        mblnChargeDataMoved = True
    End If
            
    strSQL = _
            "Select A.记录性质," & _
                   "A.执行状态," & _
                   "B.病人ID," & _
                   "B.主页ID," & _
                   "B.挂号单," & _
                   "B.病人科室ID," & _
                   "Nvl(F.费别, D.费别) as 费别," & _
                   "Decode(B.病人来源, 1, '门诊', 2, '住院', 3, '外来', 4, '体检') as 来源, "
    strSQL = strSQL & _
                   "A.执行部门ID " & _
              "From 病人医嘱发送 A," & _
                   "病人医嘱记录 B," & _
                   "病人信息     D," & _
                   "病案主页     F " & _
             "Where A.医嘱ID = B.ID " & _
                   "And B.病人ID = D.病人ID " & _
                   "And B.病人ID = F.病人ID(+) " & _
                   "And B.主页ID = F.主页ID(+) " & _
                   "And A.医嘱ID IN (SELECT ID FROM 病人医嘱记录 WHERE ID = [1] OR 相关ID = [1]) " & _
                   "And A.发送号 = [2] " & _
             "Order by A.发送时间 Desc,B.序号"
             
    '数据转储处理
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
                         
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID, SendNO)
        
    If rsTmp.BOF = False Then
        
        iPatientType = Decode(rsTmp("来源"), "门诊", 1, "体检", 1, 2)
        
        lngPatientID = rsTmp("病人ID")
        lngPageId = Nvl(rsTmp("主页ID"), 0)
        strCheckNo = Nvl(rsTmp("挂号单"), "")
        lngPatientDept = Nvl(rsTmp("病人科室ID"), 0)
        str费别 = Nvl(rsTmp!费别)
        int记录性质 = Nvl(rsTmp!记录性质, 1)
        int执行状态 = Nvl(rsTmp!执行状态, 0)
        lng开单科室ID = Nvl(rsTmp!执行部门ID, 0)
    End If
    
    
    '计算体检折扣
    If mstrSys = "体检" Then
        
        str费别 = ""
        
        strSQL = "SELECT NVL(B.体检价格,1) AS 结算折扣 FROM 体检项目医嘱 A,体检项目清单 B WHERE A.清单id=B.ID AND A.医嘱ID=[1] "
        
        '数据转储处理
        If mblnDataMoved Then
            strSQL = Replace(strSQL, "体检项目医嘱", "H体检项目医嘱")
            strSQL = Replace(strSQL, "体检项目清单", "H体检项目清单")
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, AdviceID)
        If rsTmp.BOF = False Then
            msgl体检折扣 = rsTmp("结算折扣").Value
        End If
        
    End If
    
    Call LoadMoneyList(AdviceID, lngSendNO, 0, str费别, int记录性质)
    
    Exit Sub
    
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlClearData()
    
    '----------------------------------------------------------------------------------------------------
    '功能:
    '----------------------------------------------------------------------------------------------------
    
    vsMoney.Rows = 2
    vsMoney.Cell(flexcpText, 1, 0, 1, vsMoney.Cols - 1) = ""
    
    vsDetail.Rows = 2
    vsDetail.Cell(flexcpText, 1, 0, 1, vsDetail.Cols - 1) = ""
    
End Sub

Public Function zlMenuClick(ByVal strFunc As String) As Boolean

    '----------------------------------------------------------------------------------------------------
    '功能:
    '----------------------------------------------------------------------------------------------------
    Select Case strFunc
    Case "生成主费用"
        zlMenuClick = MoneyMain
    Case "修改附加费用"
        zlMenuClick = MoneyModi
    Case "删除附加费用"
        zlMenuClick = MoneyDel
    Case "收费单据"
        zlMenuClick = MoneyNewBilling(1)
    Case "记帐单据"
        zlMenuClick = MoneyNewBilling(2)
    Case "零费耗用登记"
        zlMenuClick = MoneyNewBilling(2, True)
    End Select
    
End Function


Public Property Get Body(Optional ByVal lngIndex As Long) As Object
    Set Body = vsMoney
End Property

Private Function GetMaxNo(ByVal strNO As String, ByRef lngNO As Long, ByRef strDate As String) As Boolean
    
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT NVL(MAX(序号),0) AS 序号,NVL(MAX(登记时间),SYSDATE) AS 登记时间 FROM 病人费用记录 WHERE NO='" & strNO & "'"
    Call zlDatabase.OpenRecordset(rs, strSQL, Me.Caption)
    If rs.BOF = False Then
        lngNO = rs("序号").Value
        strDate = Format(rs("登记时间").Value, "yyyy-MM-dd HH:mm:ss")
    End If
            
    GetMaxNo = True
    
End Function

Private Function MoneyMain() As Boolean
    '----------------------------------------------------------------------------------------------------
    '功能:
    '----------------------------------------------------------------------------------------------------
    Dim rsPati As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim lng发送号
    Dim lng医嘱ID As Long
    Dim int来源 As Integer
    Dim int父号 As Integer
    Dim lng项目ID As Long
    Dim lng执行部门ID As Long
    Dim lng病人病区ID As Long
    Dim lng病人科室ID As Long
    Dim lng类别ID As Long
    Dim arrSQL As Variant
    Dim strSQL As String
    Dim strDate As String
    Dim i As Long
    Dim int保险项目否 As Integer
    Dim lng保险大类ID As Long
    Dim str保险编码 As String
    Dim cur统筹金额 As Currency
    Dim strMsg As String
    Dim strNO As String
    
    If mrsPrice Is Nothing Then Exit Function
    If AdviceID = 0 Then Exit Function

    
    If vsMoney.TextMatrix(vsMoney.Row, 2) <> "[未计费]" Then
        MsgBox "该执行项目无需计费或已经计费。" & vbCrLf & "如果需要，你可以手工补充附加费用。", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsPrice.Filter = "发送单据='" & vsMoney.TextMatrix(vsMoney.Row, 9) & "'"
    If mrsPrice.RecordCount = 0 Then
        MsgBox "该执行项目没有可以计费的主费用。" & vbCrLf & "如果需要，你可以手工补充附加费用。", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsPrice.MoveFirst
    
    If int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If MsgBox("确实要生成该项目的主费用吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Function
    End If
        
    Screen.MousePointer = 11
    
    
    'lng发送号 = lngSendNO
    lng发送号 = mrsPrice("发送号").Value
    lng病人ID = lngPatientID
    lng主页ID = lngPageId
    int来源 = iPatientType
    
    '获取病人的信息
    strSQL = "Select A.姓名,A.性别,A.年龄,Nvl(B.费别,A.费别) as 费别," & _
        " A.门诊号,A.住院号,Nvl(A.当前床号,B.出院病床) as 床号," & _
        " Nvl(A.当前病区ID,B.当前病区ID) as 病人病区ID," & _
        " Nvl(A.当前科室ID,B.出院科室ID) as 病人科室ID," & _
        " Nvl(B.险类,A.险类) as 险类,C.编码 as 付款码" & _
        " From 病人信息 A,病案主页 B,医疗付款方式 C" & _
        " Where A.病人ID=" & lng病人ID & " And A.病人ID=B.病人ID(+)" & _
        " And B.主页ID(+)=" & lng主页ID & " And A.医疗付款方式=C.名称(+)"
    Call zlDatabase.OpenRecordset(rsPati, strSQL, Me.Caption)
    
    '可能对照费用为药品费用
    If int记录性质 = 1 Then
        lng类别ID = ExistIOClass(8) '门诊划价单
    Else
        lng类别ID = ExistIOClass(9) '门诊/住院记帐单
    End If
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    arrSQL = Array()
    With mrsPrice
        .MoveFirst
        
        Dim lngMaxNo As Long
        
        strNO = mrsPrice("发送单据").Value
        
        Call GetMaxNo(strNO, lngMaxNo, strDate)
        
        If int记录性质 = 1 Then
            If BillExistBalance(strNO) Then
                MsgBox "要生成的收费单和" & strNO & "是同一张单据，" & strNO & "已经收费，不能再生成！", vbInformation, gstrSysName
                Exit Function
            End If
            strDate = "TO_DATE('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
        Else
            strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        End If
        
                       
        For i = lngMaxNo + 1 To lngMaxNo + .RecordCount
            '获取对应的医嘱信息
            If lng医嘱ID <> !医嘱ID Then
                strSQL = "Select 医嘱期效,病人科室ID,开嘱科室ID,婴儿,执行频次,计价特性" & _
                    " From 病人医嘱记录 Where ID=" & !医嘱ID
                Call zlDatabase.OpenRecordset(rsAdvice, strSQL, Me.Caption)
                
                '将当前这条计费医嘱标记为已计费
'                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱发送_计费(" & !医嘱ID & "," & lng发送号 & ")"
            End If
            lng医嘱ID = !医嘱ID
            
            '病人病区科室
            lng病人病区ID = Nvl(rsPati!病人病区ID, 0)
            lng病人科室ID = Nvl(rsPati!病人科室ID, 0)
            If lng病人科室ID = 0 Then
                lng病人病区ID = Nvl(rsAdvice!病人科室ID, 0)
                lng病人科室ID = Nvl(rsAdvice!病人科室ID, 0)
            End If
            If lng病人科室ID = 0 Then
                lng病人病区ID = UserInfo.部门ID
                lng病人科室ID = UserInfo.部门ID
            End If
            
            '每个收费项目的处理
            If lng项目ID <> !收费细目ID Then
                int父号 = i '获取价格父号
                lng执行部门ID = Get收费执行科室ID(lng病人ID, lng主页ID, !类别, !收费细目ID, !执行科室, Nvl(rsAdvice!病人科室ID, 0), Nvl(rsAdvice!开嘱科室ID, 0), int来源)
                            
                '获取保险项目信息
                If int来源 = 2 And Not IsNull(rsPati!险类) Then
                    strMsg = gclsInsure.GetItemInsure(lng病人ID, !收费细目ID, !实收, False, rsPati!险类)
                    If strMsg <> "" Then
                        int保险项目否 = Val(Split(strMsg, ";")(0))
                        lng保险大类ID = Val(Split(strMsg, ";")(1))
                        cur统筹金额 = Format(Val(Split(strMsg, ";")(2)), gstrDec)
                        str保险编码 = CStr(Split(strMsg, ";")(3))
                    End If
                End If
            End If
            lng项目ID = !收费细目ID
            
            
            Select Case mstrSys
            Case "体检"
                str费别 = ""
            End Select
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If int来源 = 1 Then
                If int记录性质 = 1 Then
                    '生成门诊划价单据
                    arrSQL(UBound(arrSQL)) = _
                        "zl_门诊划价记录_Insert('" & strNO & "'," & i & "," & lng病人ID & ",NULL," & _
                        ZVal(Nvl(rsPati!门诊号, 0)) & ",'" & Nvl(rsPati!付款码) & "','" & Nvl(rsPati!姓名) & "'," & _
                        "'" & Nvl(rsPati!性别) & "','" & Nvl(rsPati!年龄) & "','" & str费别 & "',NULL," & _
                        lng病人病区ID & "," & lng病人科室ID & "," & UserInfo.部门ID & ",'" & UserInfo.姓名 & "'," & _
                        "NULL," & lng项目ID & ",'" & !类别 & "','" & !计算单位 & "',NULL,1," & !数量 & "," & _
                        !附加手术 & "," & ZVal(lng执行部门ID) & "," & IIF(int父号 = i, "NULL", int父号) & "," & _
                        !收入项目ID & ",'" & Nvl(!收据费目) & "'," & !单价 & "," & !应收 & "," & !实收 & "," & _
                        strDate & "," & strDate & ",NULL,'" & UserInfo.姓名 & "'," & ZVal(lng类别ID) & ",NULL," & _
                        !医嘱ID & ",'" & Nvl(rsAdvice!执行频次) & "',NULL,NULL," & Nvl(rsAdvice!医嘱期效, 0) & "," & _
                        Nvl(rsAdvice!计价特性, 0) & ",1)"
                Else
                    '生成门诊记帐单据
                    arrSQL(UBound(arrSQL)) = _
                        "zl_门诊记帐记录_Insert('" & strNO & "'," & i & "," & lng病人ID & "," & _
                        ZVal(Nvl(rsPati!门诊号, 0)) & ",'" & Nvl(rsPati!姓名) & "','" & Nvl(rsPati!性别) & "'," & _
                        "'" & Nvl(rsPati!年龄) & "','" & str费别 & "',NULL," & ZVal(Nvl(rsAdvice!婴儿, 0)) & "," & _
                        lng病人病区ID & "," & lng病人科室ID & "," & UserInfo.部门ID & "," & _
                        "'" & UserInfo.姓名 & "',NULL," & lng项目ID & ",'" & !类别 & "'," & _
                        "'" & !计算单位 & "',1," & !数量 & "," & !附加手术 & "," & ZVal(lng执行部门ID) & "," & _
                        IIF(int父号 = i, "NULL", int父号) & "," & !收入项目ID & ",'" & Nvl(!收据费目) & "'," & !单价 & "," & _
                        !应收 & "," & !实收 & "," & strDate & "," & strDate & ",NULL,NULL,'" & UserInfo.编号 & "'," & _
                        "'" & UserInfo.姓名 & "'," & ZVal(lng类别ID) & ",NULL,NULL," & !医嘱ID & "," & _
                        "'" & Nvl(rsAdvice!执行频次) & "',NULL,NULL," & Nvl(rsAdvice!医嘱期效, 0) & "," & _
                        Nvl(rsAdvice!计价特性, 0) & ")"
                End If
            Else
                '生成住院记帐单据
                arrSQL(UBound(arrSQL)) = _
                    "zl_住院记帐记录_Insert('" & strNO & "'," & i & "," & lng病人ID & "," & ZVal(lng主页ID) & "," & _
                    ZVal(Nvl(rsPati!住院号, 0)) & ",'" & Nvl(rsPati!姓名) & "','" & Nvl(rsPati!性别) & "'," & _
                    "'" & Nvl(rsPati!年龄) & "','" & Nvl(rsPati!床号) & "','" & str费别 & "'," & _
                    lng病人病区ID & "," & lng病人科室ID & ",NULL," & ZVal(Nvl(rsAdvice!婴儿, 0)) & "," & _
                    UserInfo.部门ID & ",'" & UserInfo.姓名 & "',NULL," & lng项目ID & ",'" & !类别 & "'," & _
                    "'" & !计算单位 & "'," & int保险项目否 & "," & ZVal(lng保险大类ID) & ",'" & str保险编码 & "'," & _
                    "1," & !数量 & "," & !附加手术 & "," & ZVal(lng执行部门ID) & "," & _
                    IIF(int父号 = i, "NULL", int父号) & "," & !收入项目ID & ",'" & Nvl(!收据费目) & "'," & !单价 & "," & _
                    !应收 & "," & !实收 & "," & cur统筹金额 & "," & strDate & "," & strDate & ",NULL,NULL," & _
                    "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',NULL," & ZVal(lng类别ID) & ",NULL,NULL,NULL," & _
                    !医嘱ID & ",'" & Nvl(rsAdvice!执行频次) & "',NULL,NULL," & Nvl(rsAdvice!医嘱期效, 0) & "," & _
                    Nvl(rsAdvice!计价特性, 0) & ",NULL)"
            End If
            
            .MoveNext
        Next
    End With
    
    '设置医嘱计费标志
'    strSQL = _
'            "SELECT ID FROM 病人医嘱记录 WHERE ID=" & AdviceID
'    strSQL = strSQL & " UNION ALL " & _
'            "SELECT ID FROM 病人医嘱记录 WHERE 相关ID=" & AdviceID
    
    strSQL = "SELECT 医嘱ID FROM 病人医嘱发送 WHERE NVL(计费状态,0)=0 AND NO='" & strNO & "'"
    
    Call zlDatabase.OpenRecordset(rsAdvice, strSQL, Me.Caption)
    If rsAdvice.BOF = False Then
        Do While Not rsAdvice.EOF
            
            '将当前这条计费医嘱标记为已计费
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱发送_计费(" & rsAdvice("医嘱ID").Value & "," & lng发送号 & ")"
            
            rsAdvice.MoveNext
        Loop
    End If
    
    On Error GoTo errH
    
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    
    Dim strSQL1() As String
    ReDim strSQL1(0 To 1)
    strSQL1(0) = ""
    strSQL1(1) = ""
    
    If int记录性质 <> 1 Then
        If 材料自动发料(strSQL1, strNO) = False Then GoTo errH
    End If
    
    For i = 1 To UBound(strSQL1)
        If Trim(strSQL1(i)) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL1(i), Me.Caption)
    Next
    
    '在提交前进行医保传输
    If int来源 = 2 And Not IsNull(rsPati!险类) Then
        If gclsInsure.GetCapability(support记帐上传, , rsPati!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, , rsPati!险类) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strMsg, , rsPati!险类) Then
                gcnOracle.RollbackTrans
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    
    gcnOracle.CommitTrans
    
    '在提交后进行医保传输
    If int来源 = 2 And Not IsNull(rsPati!险类) Then
        If gclsInsure.GetCapability(support记帐上传, , rsPati!险类) And gclsInsure.GetCapability(support记帐完成后上传, , rsPati!险类) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strMsg, , rsPati!险类) Then
                If strMsg <> "" Then
                    MsgBox strMsg, vbInformation, gstrSysName
                Else
                    MsgBox "单据""" & strNO & """的数据向医保传送失败,该单据已保存！", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    On Error GoTo 0
    Screen.MousePointer = 0
    
    MsgBox "执行项目的主费用生成成功。", vbInformation, gstrSysName
    '刷新
    
    MoneyMain = True
    
    Me.Tag = "Loading": Call Form_Activate
    
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function 材料自动发料(ByRef strSQL() As String, ByVal strNO As String) As Boolean
    
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim bln自动发料 As Boolean
    
    On Error GoTo ErrHand
    
    bln自动发料 = GetSysParVal(63) = "1"
    If bln自动发料 = False Then
        材料自动发料 = True
        Exit Function
    End If
    
    strTmp = "SELECT DISTINCT A.执行部门id FROM 病人费用记录 A,材料特性 B WHERE A.收费细目id=B.材料id AND NVL(B.跟踪在用,0)=1 AND A.收费类别='4' and A.NO='" & strNO & "'"
    Call zlDatabase.OpenRecordset(rs, strTmp, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            If zlCommFun.Nvl(rs("执行部门id").Value, 0) > 0 Then
                strSQL(1) = "zl_材料收发记录_处方发料(" & rs("执行部门id").Value & ",25,'" & strNO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1,Sysdate)"
            End If
            rs.MoveNext
        Loop
    End If
    
    材料自动发料 = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function MoneyModi() As Boolean
    '----------------------------------------------------------------------------------------------------
    '功能:
    '----------------------------------------------------------------------------------------------------
    
    Dim lng病人ID As Long, lng主页ID As Long
    Dim lng病人科室ID As Long
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim int病人来源 As Integer, int记录性质 As Integer
    Dim strNO As String, bln零耗 As Boolean
    
    
    If AdviceID = 0 Then Exit Function
    If vsMoney.TextMatrix(vsMoney.Row, 0) = "主费用" Then
        MsgBox "执行项目的主费用不能修改。如果需要，你可以手工补充附加费用。", vbInformation, gstrSysName
        Exit Function
    End If
    If int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Function
    End If
    
    With vsMoney
        strNO = .TextMatrix(.Row, 2)
        If strNO = "" Or strNO = "[未计费]" Then Exit Function
        int记录性质 = .Cell(flexcpData, .Row, 1)
        
        If InStr(.TextMatrix(.Row, 1), "√") > 0 Then
            MsgBox "该单据已经收费，不能再修改。", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If zlDatabase.NOMoved("病人费用记录", strNO) Then
        MsgBox "该单据已经转出，不能再修改。", vbInformation, gstrSysName
        Exit Function
    End If
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    If mSysName = "检验" Then
        strSQL = "SELECT MIN(ID) AS ID FROM 病人医嘱记录 WHERE 相关id=" & AdviceID
        zlDatabase.OpenRecordset rs, strSQL, Me.Caption
        If rs.BOF Then Exit Function
        lng医嘱ID = zlCommFun.Nvl(rs("ID").Value)
    Else
        lng医嘱ID = AdviceID
    End If
    
    
    lng发送号 = lngSendNO
    lng病人ID = lngPatientID
    lng主页ID = lngPageId
    int病人来源 = iPatientType
    lng病人科室ID = lngPatientDept
    
    If int记录性质 = 2 Then
       bln零耗 = BillisZeroLog(strNO)
    End If
    
    
    frmTechnicExpense.mstrPrivs = mstrPrivs
    frmTechnicExpense.mbytInState = 0
    frmTechnicExpense.mbln费用登记 = bln零耗
    frmTechnicExpense.mstrInNO = strNO
    frmTechnicExpense.mlng医嘱ID = lng医嘱ID
    frmTechnicExpense.mlng发送号 = lng发送号
    frmTechnicExpense.mlng病人ID = lng病人ID
    frmTechnicExpense.mlng主页ID = lng主页ID
    frmTechnicExpense.mint病人来源 = int病人来源
    frmTechnicExpense.mint记录性质 = int记录性质
    frmTechnicExpense.mlng开单科室ID = lng开单科室ID
    frmTechnicExpense.mlng病人科室id = lng病人科室ID
    On Error Resume Next
    frmTechnicExpense.Show 1, Me
    On Error GoTo 0
    If gblnOK Then
        '刷新
        MoneyModi = True
        Me.Tag = "Loading": Call Form_Activate
        
    End If
End Function

Private Function MoneyDel() As Boolean
    '----------------------------------------------------------------------------------------------------
    '功能:
    '----------------------------------------------------------------------------------------------------
    
    Dim int病人来源 As Integer, int记录性质 As Integer
    Dim strNO As String
    
    If AdviceID = 0 Then Exit Function
    If vsMoney.TextMatrix(vsMoney.Row, 0) = "主费用" Then
        MsgBox "执行项目的主费用不能删除。", vbInformation, gstrSysName
        Exit Function
    End If
    If int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Function
    End If
    
    With vsMoney
        strNO = .TextMatrix(.Row, 2)
        If strNO = "" Or strNO = "[未计费]" Then Exit Function
        int记录性质 = .Cell(flexcpData, .Row, 1)
    
        If InStr(.TextMatrix(.Row, 1), "√") > 0 Then
            MsgBox "该单据已经收费，不能再删除。", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If zlDatabase.NOMoved("病人费用记录", strNO) Then
        MsgBox "该单据已经转出，不能再删除。", vbInformation, gstrSysName
        Exit Function
    End If
    
    int病人来源 = iPatientType
    
    frmTechnicExpense.mstrPrivs = mstrPrivs
    frmTechnicExpense.mbytInState = 3
    frmTechnicExpense.mstrInNO = strNO
    frmTechnicExpense.mint病人来源 = int病人来源
    frmTechnicExpense.mint记录性质 = int记录性质
    On Error Resume Next
    frmTechnicExpense.Show 1, Me
    On Error GoTo 0
    If gblnOK Then
        '刷新
        MoneyDel = True
        Me.Tag = "Loading": Call Form_Activate
    End If
End Function

Private Function MoneyNewBilling(ByVal iRecordType As Integer, Optional OnlyRecord As Boolean = False) As Boolean
    '----------------------------------------------------------------------------------------------------
    '功能:
    '----------------------------------------------------------------------------------------------------
    
    Dim lng病人ID As Long, lng主页ID As Long
    Dim lng病人科室ID As Long
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim int病人来源 As Integer
    
    If AdviceID = 0 Then Exit Function
    
    If int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Function
    End If
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    If mSysName = "检验" Then
        strSQL = "SELECT MIN(ID) AS ID FROM 病人医嘱记录 WHERE 相关id=" & AdviceID
        zlDatabase.OpenRecordset rs, strSQL, Me.Caption
        If rs.BOF Then Exit Function
        lng医嘱ID = zlCommFun.Nvl(rs("ID").Value)
    Else
        lng医嘱ID = AdviceID
    End If
        
    lng发送号 = lngSendNO
    lng病人ID = lngPatientID
    lng主页ID = lngPageId
    int病人来源 = iPatientType
    lng病人科室ID = lngPatientDept
    
    frmTechnicExpense.mstrPrivs = mstrPrivs
    frmTechnicExpense.mbytInState = 0
    frmTechnicExpense.mlng医嘱ID = lng医嘱ID
    frmTechnicExpense.mlng发送号 = lng发送号
    frmTechnicExpense.mlng病人ID = lng病人ID
    frmTechnicExpense.mlng主页ID = lng主页ID
    frmTechnicExpense.mint病人来源 = int病人来源
    frmTechnicExpense.mint记录性质 = iRecordType
    frmTechnicExpense.mbln费用登记 = OnlyRecord
    frmTechnicExpense.mlng开单科室ID = lng开单科室ID
    frmTechnicExpense.mlng病人科室id = lng病人科室ID
    On Error Resume Next
    frmTechnicExpense.Show 1, Me
    On Error GoTo 0
    If gblnOK Then
        '刷新
        MoneyNewBilling = True
        Me.Tag = "Loading": Call Form_Activate
    End If
End Function

Private Sub Form_Activate()
    On Error Resume Next

    If Me.Tag = "Loading" Then
        mfrmParent.Refresh
        Me.Tag = ""
        Call LoadMoneyList(AdviceID, lngSendNO, 0, str费别, int记录性质)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHand
    
    Call mfrmParent.ActiveFormKeyDown(KeyCode, Shift)

ErrHand:

End Sub

Private Sub Form_Load()
    
    On Error GoTo ShowError
    
'    Set mrsPrice = Nothing
    
    Call InitMoneyTable
    Call InitDetailTable
    
    Exit Sub
    
ShowError:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
'    If imgX.Top > Me.ScaleHeight - 1000 Then imgX.Top = Me.ScaleHeight - 1000
        
    With vsMoney
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = imgX.Top - .Top
    End With
    
    With imgX
        .Left = 0
        .Top = vsMoney.Top + vsMoney.Height
        .Width = Me.ScaleWidth
    End With
    
    With vsDetail
        .Left = 0
        .Top = imgX.Top + imgX.Height
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top
    End With
            
    Call AppendRows(vsMoney, lnX0, lnY0)
    Call AppendRows(vsDetail, lnX1, lnY1)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Set mrsPrice = Nothing
End Sub


Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 1 Then Exit Sub
    
    imgX.Top = imgX.Top + y
    
    If imgX.Top < 1500 Then imgX.Top = 1500
    If Me.Height - imgX.Top - imgX.Height < 1000 Then imgX.Top = Me.Height - imgX.Height - 1000

    Form_Resize
End Sub

Private Sub vsDetail_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsDetail, lnX1, lnY1)
End Sub

Private Sub vsDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsDetail, lnX1, lnY1)
End Sub

Private Sub vsDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub vsMoney_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = OldRow Then Exit Sub
    If NewCol >= vsMoney.FixedCols And NewRow >= vsMoney.FixedRows Then
        vsMoney.ForeColorSel = vsMoney.Cell(flexcpForeColor, NewRow, 0)
        Call LoadBillDetail(NewRow)
    End If
    
    On Error GoTo ErrHand
    
    Call mfrmParent.ActiveFormEnabled
    
ErrHand:
End Sub

Private Sub vsMoney_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsMoney, lnX0, lnY0)
End Sub

Private Sub vsMoney_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsMoney, lnX0, lnY0)
End Sub

Private Sub vsMoney_DblClick()
    If vsMoney.MouseRow >= vsMoney.FixedRows Then
        Call vsMoney_KeyPress(13)
    End If
End Sub

Private Sub vsMoney_GotFocus()
    vsMoney.BackColorSel = COLOR_FOCUS
End Sub

Private Sub vsMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub vsMoney_KeyPress(KeyAscii As Integer)
    
    Dim int病人来源 As Integer
    Dim int记录性质 As Integer
    Dim strNO As String
    
    If KeyAscii = 13 Or KeyAscii = 32 Then
        KeyAscii = 0
        With vsMoney
            strNO = .TextMatrix(.Row, 2)
            If strNO = "" Or strNO = "[未计费]" Then Exit Sub
            int记录性质 = .Cell(flexcpData, .Row, 1)
        End With
        int病人来源 = iPatientType
        
        frmTechnicExpense.mstrPrivs = mstrPrivs
        frmTechnicExpense.mbytInState = 1
        frmTechnicExpense.mstrInNO = strNO
        frmTechnicExpense.mint病人来源 = int病人来源
        frmTechnicExpense.mint记录性质 = int记录性质
        On Error Resume Next
        frmTechnicExpense.Show 1, Me
    End If
End Sub

Private Sub vsMoney_LostFocus()
    vsMoney.BackColorSel = COLOR_LOST
End Sub

Private Sub vsDetail_GotFocus()
    vsDetail.BackColorSel = COLOR_FOCUS
End Sub

Private Sub vsDetail_LostFocus()
    vsDetail.BackColorSel = COLOR_LOST
End Sub

Private Sub InitMoneyTable()
    '----------------------------------------------------------------------------------------------------
    '功能:
    '----------------------------------------------------------------------------------------------------
    
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "费用类型,900,1;单据类型,1000,1;单据号,900,1;费别,900,1;应收金额,1000,7;实收金额,1000,7;开单人,750,1;登记时间,1080,1;登记人,750,1;发送单据,0,1"
    arrHead = Split(strHead, ";")
    With vsMoney
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        
        .Cols = .Cols + 1
        .ExtendLastCol = True
        Call AppendRows(vsMoney, lnX0, lnY0)
        
    End With
End Sub

Private Sub InitDetailTable()
    '----------------------------------------------------------------------------------------------------
    '功能:
    '----------------------------------------------------------------------------------------------------
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "类别,650,1;项目,3000,1;数量,1000,1;单价,1000,7;应收金额,1000,7;实收金额,1000,7;执行科室,1000,1"
    arrHead = Split(strHead, ";")
    With vsDetail
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        
        .Cols = .Cols + 1
        .ExtendLastCol = True
        Call AppendRows(vsDetail, lnX1, lnY1)
    End With
End Sub


Private Function LoadMoneyList(ByVal lng医嘱ID As Long, _
                                ByVal lng发送号 As Long, _
                                ByVal int计费状态 As Integer, _
                                ByVal str费别 As String, _
                                ByVal int记录性质 As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：读取指定医嘱的主要费用及附加费用
    '说明：1.包含医嘱本身的主费用及附加费用,主费用可能尚未产生
    '      2.目前单据暂不支持部份退费,所以清单中只需简单显示
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rsList As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    Dim blnMain As Boolean
    Dim blnSub As Boolean
    Dim strPre As String
    Dim lngRow As Long
    Dim cur应收 As Currency
    Dim cur实收 As Currency
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errH
    
    '未计费的
    Dim strTmp As String
    
    If mstrSys = "LIS" Then
            
        strTmp = _
            "(Select ID From 病人医嘱记录 X,(Select " & lng医嘱ID & " as 医嘱id From dual union Select 医嘱id From 检验项目分布 A " & _
            "Where 标本ID In (Select ID From 检验标本记录 Where 医嘱id=" & lng医嘱ID & ") " & _
                  ") Y Where Y.医嘱id In (X.ID,X.相关ID)) "
    Else
        strTmp = _
            "(Select ID from 病人医嘱记录 WHERE " & lng医嘱ID & " IN (ID,相关id)) "
    End If
    
    If int执行状态 <> 1 Then
        strSQL = _
            "SELECT DISTINCT A.记录性质,A.NO,A.发送号 " & _
            "FROM 病人医嘱发送 A " & _
            "WHERE   A.医嘱ID IN " & strTmp & _
                    "AND NVL(A.计费状态,0)=0"
        
        '数据转储处理
        If mblnDataMoved Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "检验标本记录", "H检验标本记录")
            strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        End If
        
        Call zlDatabase.OpenRecordset(rs, strSQL, Me.Caption)
        strSQL = ""
        If rs.BOF = False Then
            Do While Not rs.EOF
                
                '未计费状态,直接读取收费关系显示
                cur应收 = 0
                cur实收 = 0
                Call NewAdvicePrice(rs, cur应收, cur实收)
                
                If cur实收 > 0 Then
                    
                    '判断此收费单据号是否收费
                    If int记录性质 = 1 Then
                        If Not BillExistBalance(rs("NO").Value) Then
                            strSQL = strSQL & IIF(strSQL <> "", " Union ALL ", "") & _
                                    "Select 1 as 费用类型," & _
                                            int记录性质 & " as 记录性质," & _
                                            "0 as 已收费," & _
                                            "'[未计费]' as NO," & _
                                            "'" & str费别 & "' as 费别," & _
                                            cur应收 & " as 应收金额," & _
                                            cur实收 & " as 实收金额,'" & _
                                            UserInfo.姓名 & "' as 开单人," & _
                                            "Sysdate as 登记时间,'" & _
                                            UserInfo.姓名 & "' as 操作员,'" & rs("NO").Value & "' AS 发送单据 " & _
                                    " From Dual"
                        End If
                    Else
                    
                        strSQL = strSQL & IIF(strSQL <> "", " Union ALL ", "") & _
                            "Select 1 as 费用类型," & _
                                    int记录性质 & " as 记录性质," & _
                                    "0 as 已收费," & _
                                    "'[未计费]' as NO," & _
                                    "'" & str费别 & "' as 费别," & _
                                    cur应收 & " as 应收金额," & _
                                    cur实收 & " as 实收金额,'" & _
                                    UserInfo.姓名 & "' as 开单人," & _
                                    "Sysdate as 登记时间,'" & _
                                    UserInfo.姓名 & "' as 操作员,'" & rs("NO").Value & "' AS 发送单据 " & _
                            " From Dual"
                        End If
                End If
                
                rs.MoveNext
            Loop
        End If
    End If
    
    
    '已计费的
    strSQL = strSQL & IIF(strSQL <> "", " Union ALL ", "") & _
        " Select 1 as 费用类型,A.记录性质,Decode(B.记录状态,1,1,0) as 已收费," & _
        " A.NO,B.费别,Sum(B.应收金额) as 应收金额,Sum(B.实收金额) as 实收金额," & _
        " B.开单人,B.登记时间,Nvl(B.操作员姓名,B.划价人) as 操作员,'' AS 发送单据 " & _
        " From 病人医嘱发送 A,病人费用记录 B" & _
        " Where NVL(A.计费状态,0)=1 AND A.医嘱ID IN " & strTmp & _
        " And A.NO=B.NO And A.记录性质=B.记录性质 And A.医嘱ID=B.医嘱序号+0 " & _
        " And B.记录状态 in (0,1) " & _
        " Group by A.记录性质,B.记录状态,A.NO,B.费别,B.开单人,B.登记时间,Nvl(B.操作员姓名,B.划价人)"
        
    '附费用部份
    strSQL = strSQL & IIF(strSQL <> "", " Union ALL ", "") & _
            " Select 2 as 费用类型,A.记录性质,Decode(B.记录状态,1,1,0) as 已收费," & _
            " A.NO,B.费别,Sum(B.应收金额) as 应收金额,Sum(B.实收金额) as 实收金额," & _
            " B.开单人,B.登记时间,Nvl(B.操作员姓名,B.划价人) as 操作员,'' as 发送单据 " & _
            " From 病人医嘱附费 A,病人费用记录 B" & _
            " Where A.医嘱ID in (select id from 病人医嘱记录 where " & lng医嘱ID & " in (ID,相关id)) " & _
            " And A.NO=B.NO And A.记录性质=Decode(B.记录性质,0,1,B.记录性质)" & _
            " And B.记录状态 in (0,1) " & _
            " Group by A.记录性质,B.记录状态,A.NO,B.费别,B.开单人,B.登记时间,Nvl(B.操作员姓名,B.划价人)"
    
    '数据转储处理
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人费用记录", "H病人费用记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱附费", "H病人医嘱附费")
    ElseIf mblnChargeDataMoved Then
        strTmp = strSQL
        strTmp = Replace(strTmp, "病人费用记录", "H病人费用记录")
        strSQL = strSQL & " Union All " & strTmp
    End If
    
    strSQL = "Select * From (" & strSQL & ") Order by 费用类型,登记时间 Desc"
    Call zlDatabase.OpenRecordset(rsList, strSQL, Me.Caption)
    With vsMoney
        lngRow = .FixedRows
        strPre = .TextMatrix(.Row, 1) & "_" & .TextMatrix(.Row, 2)
        
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
        vsDetail.Rows = vsDetail.FixedRows
        vsDetail.Rows = vsDetail.FixedRows + 1
        
        mblnCash = True
        If Not rsList.EOF Then
            .Rows = rsList.RecordCount + 1
            For i = 1 To rsList.RecordCount
            
                '是否全部已收费,要管未审核的记帐划价单,不管尚未计费的主费用
                If rsList!已收费 = 0 And Nvl(rsList!实收金额, 0) <> 0 And rsList!NO <> "[未计费]" Then mblnCash = False
                                
                .TextMatrix(i, 0) = IIF(rsList!费用类型 = 1, "主费用", "附加费用")
                .TextMatrix(i, 1) = IIF(rsList!记录性质 = 1, "收费单据" & IIF(rsList!已收费 = 1, "√", ""), "记帐单据")
                If rsList!记录性质 = 1 And rsList!已收费 = 1 Then '已收费蓝色显示
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HC00000 '深蓝
                End If
                
                .TextMatrix(i, 2) = rsList!NO
                .TextMatrix(i, 3) = Nvl(rsList!费别)
                .TextMatrix(i, 4) = Format(Nvl(rsList!应收金额, 0), gstrDec)
                .TextMatrix(i, 5) = Format(Nvl(rsList!实收金额, 0), gstrDec)
                .TextMatrix(i, 6) = Nvl(rsList!开单人)
                .TextMatrix(i, 7) = Format(rsList!登记时间, "MM-dd HH:mm")
                .TextMatrix(i, 8) = Nvl(rsList!操作员)
                .TextMatrix(i, 9) = Nvl(rsList!发送单据)
                '附加数据
                .Cell(flexcpData, i, 1) = CInt(rsList!记录性质)
                .Cell(flexcpData, i, 7) = Format(rsList!登记时间, "yyyy-MM-dd HH:mm:ss")
                                                
                '是否产生冻结行
                If rsList!费用类型 = 1 Then blnMain = True
                If rsList!费用类型 = 2 Then blnSub = True
                
                '定位到原行位置
                If .TextMatrix(i, 1) & "_" & .TextMatrix(i, 2) = strPre Then lngRow = i
                rsList.MoveNext
            Next
        Else
            mblnCash = False
        End If
        If blnMain And blnSub Then
            .FrozenRows = 1
            .Select 1, 0, 1, .Cols - 1
            .CellBorder &HC00000, 0, 0, 0, 1, 0, 0
        End If
        .Row = lngRow: .Col = .FixedCols
        Call .ShowCell(.Row, .Col)
        
        .Redraw = flexRDDirect
        Call vsMoney_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    Call AppendRows(vsMoney, lnX0, lnY0)
        
    LoadMoneyList = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function NewAdvicePrice(ByVal rs As ADODB.Recordset, ByRef cur应收 As Currency, ByRef cur实收 As Currency) As Boolean
    '----------------------------------------------------------------------------------------------------
    '功能：读取指定医嘱的计价关系到临时记录集
    '说明：要计算的项目应该不是叮嘱,院外执行,无需计费
    '----------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    Dim dbl数量 As Double
    Dim bln附加手术 As Boolean
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
        
    On Error GoTo ErrHand
            
    '读取要计算主费用的医嘱记录(包含附加手术,检查部位；手术麻醉单独)
    strSQL = _
            " Select B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID,B.开嘱科室ID," & _
            " Nvl(A.发送数次,Sum(Nvl(C.本次数次,0))) as 数量 " & _
            " From 病人医嘱发送 A,病人医嘱记录 B,病人医嘱执行 C" & _
            " Where NVL(A.计费状态,0)=0 AND A.NO=[1] " & _
                " And A.医嘱ID=B.ID And A.发送号+0=[2] " & _
                " And C.医嘱ID(+)=A.医嘱ID And C.发送号(+)=A.发送号" & _
            " Group by B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID,B.开嘱科室ID,A.发送数次"
            
    '数据转储处理
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱执行", "H病人医嘱执行")
    End If
    
    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(rs("No").Value), Val(rs("发送号").Value))
    For i = 1 To rsAdvice.RecordCount
        dbl数量 = Nvl(rsAdvice!数量, 0)
        
        '读取对应的收费价目:只读取固定对照,且不是变价的对照
        bln附加手术 = (rsAdvice!诊疗类别 = "F" And Not IsNull(rsAdvice!相关ID))
        strSQL = IIF(bln附加手术, "Nvl(B.附术收费率,100)/100", "1") & " as 附术率"
        strSQL = _
                " Select A.收费项目ID,A.收费数量,B.收入项目ID,D.收据费目,C.类别," & _
                " C.计算单位,C.执行科室,Decode(C.是否变价,1,NULL,B.现价) as 单价," & strSQL & _
                " From 诊疗收费关系 A,收费价目 B,收费项目目录 C,收入项目 D" & _
                " Where A.诊疗项目ID=[1] " & _
                    " And A.收费项目ID=B.收费细目ID And A.收费项目ID=C.ID And B.收入项目ID=D.ID" & _
                    " And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " And Nvl(C.是否变价,0)=0"
                    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsAdvice!诊疗项目ID))
        For j = 1 To rsTmp.RecordCount
        
            mrsPrice.AddNew
            
            mrsPrice!医嘱ID = rsAdvice!医嘱ID
            mrsPrice!开嘱科室ID = rsAdvice!开嘱科室ID
            mrsPrice!类别 = rsTmp!类别
            mrsPrice!收费细目ID = rsTmp!收费项目ID
            mrsPrice!计算单位 = Nvl(rsTmp!计算单位)
            mrsPrice!附加手术 = IIF(bln附加手术, 1, 0)
            mrsPrice!执行科室 = Nvl(rsTmp!执行科室, 0)
            mrsPrice!收入项目ID = rsTmp!收入项目ID
            mrsPrice!收据费目 = rsTmp!收据费目
            mrsPrice!单价 = Format(Nvl(rsTmp!单价, 0), "0.00000")
            mrsPrice!数量 = Format(Nvl(rsTmp!收费数量, 0) * dbl数量, "0.00000")
            mrsPrice!应收 = Format(mrsPrice!数量 * mrsPrice!单价 * rsTmp!附术率, gstrDec)
            mrsPrice!发送单据 = rs("NO").Value
            mrsPrice!发送号 = rs("发送号").Value
            
            Select Case mstrSys
            Case "体检"
                'mrsPrice!实收 = Format(msgl体检折扣 * mrsPrice!应收, gstrDec)
                mrsPrice!实收 = mrsPrice!应收
            Case Else
                If str费别 = "" Then
                    mrsPrice!实收 = mrsPrice!应收
                Else
                    mrsPrice!实收 = Format(ActualMoney(str费别, mrsPrice!收入项目ID, mrsPrice!应收), gstrDec)
                End If
            End Select
            
            cur应收 = cur应收 + zlCommFun.Nvl(mrsPrice!应收, 0)
            cur实收 = cur实收 + zlCommFun.Nvl(mrsPrice!实收, 0)
            
            mrsPrice.Update
            rsTmp.MoveNext
        Next
        
        rsAdvice.MoveNext
        
    Next
        
    Dim sgl合计 As Single
    
    If mstrSys = "体检" And msgl体检折扣 > 0 Then
        If mrsPrice.RecordCount > 0 Then mrsPrice.MoveFirst
        For j = 1 To mrsPrice.RecordCount
            mrsPrice!实收 = Format(msgl体检折扣 * mrsPrice!实收 / cur实收, gstrDec)
            sgl合计 = sgl合计 + mrsPrice!实收
            mrsPrice.MoveNext
        Next
        
        If sgl合计 <> msgl体检折扣 Then
            mrsPrice!实收 = mrsPrice!实收 + (msgl体检折扣 - sgl合计)
        End If
        
        cur实收 = msgl体检折扣
        
        mrsPrice.Update
    End If
    If mrsPrice.RecordCount > 0 Then mrsPrice.MoveFirst
    
    NewAdvicePrice = True
    
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Private Function LoadAdvicePrice(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByVal str费别 As String, ByVal strNO As String) As Boolean
    '----------------------------------------------------------------------------------------------------
    '功能：读取指定医嘱的计价关系到临时记录集
    '说明：要计算的项目应该不是叮嘱,院外执行,无需计费
    '----------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    Dim dbl数量 As Double
    Dim bln附加手术 As Boolean
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
        
    On Error GoTo errH
            
    '读取要计算主费用的医嘱记录(包含附加手术,检查部位；手术麻醉单独)
    strSQL = _
            " Select B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID,B.开嘱科室ID," & _
            " Nvl(A.发送数次,Sum(Nvl(C.本次数次,0))) as 数量" & _
            " From 病人医嘱发送 A,病人医嘱记录 B,病人医嘱执行 C" & _
            " Where NVL(A.计费状态,0)=0 AND B.相关ID=[1] " & _
                " And A.医嘱ID=B.ID And A.发送号+0=[2] " & _
                " And C.医嘱ID(+)=A.医嘱ID And C.发送号(+)=A.发送号" & _
            " Group by B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID,B.开嘱科室ID,A.发送数次"
            
    strSQL = strSQL & " Union ALL " & _
            " Select B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID,B.开嘱科室ID," & _
            " Nvl(A.发送数次,Sum(Nvl(C.本次数次,0))) as 数量" & _
            " From 病人医嘱发送 A,病人医嘱记录 B,病人医嘱执行 C" & _
            " Where NVL(A.计费状态,0)=0 AND B.ID=[1] " & _
                " And A.医嘱ID=B.ID And A.发送号+0=[2] " & _
                " And C.医嘱ID(+)=A.医嘱ID And C.发送号(+)=A.发送号" & _
                " Group by B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID,B.开嘱科室ID,A.发送数次" & _
            " Order by 序号"
    
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱执行", "H病人医嘱执行")
    End If
    
    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
    For i = 1 To rsAdvice.RecordCount
        dbl数量 = Nvl(rsAdvice!数量, 0)
        
        '读取对应的收费价目:只读取固定对照,且不是变价的对照
        bln附加手术 = (rsAdvice!诊疗类别 = "F" And Not IsNull(rsAdvice!相关ID))
        strSQL = IIF(bln附加手术, "Nvl(B.附术收费率,100)/100", "1") & " as 附术率"
        strSQL = _
                " Select A.收费项目ID,A.收费数量,B.收入项目ID,D.收据费目,C.类别," & _
                " C.计算单位,C.执行科室,Decode(C.是否变价,1,NULL,B.现价) as 单价," & strSQL & _
                " From 诊疗收费关系 A,收费价目 B,收费项目目录 C,收入项目 D" & _
                " Where A.诊疗项目ID=" & rsAdvice!诊疗项目ID & _
                    " And A.收费项目ID=B.收费细目ID And A.收费项目ID=C.ID And B.收入项目ID=D.ID" & _
                    " And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " And Nvl(C.是否变价,0)=0"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        For j = 1 To rsTmp.RecordCount
        
            mrsPrice.AddNew
            
            mrsPrice!医嘱ID = rsAdvice!医嘱ID
            mrsPrice!开嘱科室ID = rsAdvice!开嘱科室ID
            mrsPrice!类别 = rsTmp!类别
            mrsPrice!收费细目ID = rsTmp!收费项目ID
            mrsPrice!计算单位 = Nvl(rsTmp!计算单位)
            mrsPrice!附加手术 = IIF(bln附加手术, 1, 0)
            mrsPrice!执行科室 = Nvl(rsTmp!执行科室, 0)
            mrsPrice!收入项目ID = rsTmp!收入项目ID
            mrsPrice!收据费目 = rsTmp!收据费目
            mrsPrice!单价 = Format(Nvl(rsTmp!单价, 0), "0.00000")
            mrsPrice!数量 = Format(Nvl(rsTmp!收费数量, 0) * dbl数量, "0.00000")
            mrsPrice!应收 = Format(mrsPrice!数量 * mrsPrice!单价 * rsTmp!附术率, gstrDec)
            
            Select Case mstrSys
            Case "体检"
                mrsPrice!实收 = Format(msgl体检折扣 * mrsPrice!应收, gstrDec)
            Case Else
                If str费别 = "" Then
                    mrsPrice!实收 = mrsPrice!应收
                Else
                    mrsPrice!实收 = Format(ActualMoney(str费别, mrsPrice!收入项目ID, mrsPrice!应收), gstrDec)
                End If
            End Select
            
            mrsPrice.Update
            rsTmp.MoveNext
        Next
        
        rsAdvice.MoveNext
        
    Next
    If mrsPrice.RecordCount > 0 Then mrsPrice.MoveFirst
    LoadAdvicePrice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
'    Set mrsPrice = Nothing
End Function

Private Function LoadBillDetail(ByVal lngRow As Long) As Boolean
    '----------------------------------------------------------------------------------------------------
    '功能：显示单据明细内容
    '参数：lngRow=单据清单行
    '----------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strNO As String, int记录性质 As Integer
    Dim lng病人科室ID As Long, int来源 As Integer
    Dim lng项目ID As Long, int父号 As Integer
    Dim strSQL As String, i As Long
    Dim str登记时间 As String
    Dim bln药房单位 As Boolean, str药房单位 As String, str药房包装 As String

    On Error GoTo errH

    If lngRow < vsMoney.FixedRows Then Exit Function

    vsDetail.Rows = vsDetail.FixedRows
    vsDetail.Rows = vsDetail.FixedRows + 1

    With vsMoney
        strNO = .TextMatrix(lngRow, 2)
        int记录性质 = Val(.Cell(flexcpData, lngRow, 1))
        lng病人科室ID = lngPatientDept
        int来源 = iPatientType
        
        '登记时间是为了区分同时存在审核与未审核的情况
        str登记时间 = "To_Date('" & .Cell(flexcpData, lngRow, 7) & "','YYYY-MM-DD HH24:MI:SS')"
        
        If strNO = "" Then Exit Function
    End With

    '药品单位
    bln药房单位 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "药品单位", 0)) <> 0
    If int来源 = 1 Then
        str药房单位 = "门诊单位": str药房包装 = "门诊包装"
    Else
        str药房单位 = "住院单位": str药房包装 = "住院包装"
    End If

    If strNO = "[未计费]" Then
        If mrsPrice Is Nothing Then Exit Function
        With mrsPrice
            .Filter = "发送单据='" & vsMoney.TextMatrix(lngRow, 9) & "'"
            If .RecordCount > 0 Then
                .MoveFirst
                For i = 1 To .RecordCount
                    If lng项目ID <> !收费细目ID Then int父号 = i
                    strSQL = strSQL & IIF(strSQL <> "", " Union ALL ", "") & _
                        " Select " & i & " as 序号," & IIF(int父号 = i, "-NULL", int父号) & " as 价格父号," & _
                        "'" & strNO & "' as NO," & int记录性质 & " as 记录性质,1 as 记录状态," & _
                        "'" & !类别 & "' as 收费类别," & !收费细目ID & " as 收费细目ID," & _
                        Get收费执行科室ID(lngPatientID, lngPageId, !类别, !收费细目ID, !执行科室, lng病人科室ID, Nvl(!开嘱科室ID, 0), int来源) & " as 执行部门ID," & _
                        !收入项目ID & " as 收入项目ID,1 as 付数," & !数量 & " as 数次," & !单价 & " as 标准单价," & _
                        !应收 & " as 应收金额," & !实收 & " as 实收金额 From Dual"
    
                    lng项目ID = !收费细目ID
                    .MoveNext
                Next
                If strSQL = "" Then Exit Function
                strSQL = "(" & strSQL & ")"
            Else
                strSQL = "病人费用记录"
            End If
            .Filter = ""
        End With
    Else
        If zlDatabase.NOMoved("病人费用记录", strNO) Then
            strSQL = "H病人费用记录"
        Else
            strSQL = "病人费用记录"
        End If
    End If

    strSQL = "Select C.名称 as 类别,Nvl(F.名称,B.名称)||Decode(B.规格,NULL,NULL,' '||B.规格) as 项目," & _
        " Sum(A.标准单价" & IIF(bln药房单位, "*Nvl(E." & str药房包装 & ",1)", "") & ") as 单价," & _
        " Avg(Nvl(A.付数,1)*A.数次" & IIF(bln药房单位, "/Nvl(E." & str药房包装 & ",1)", "") & ") as 数量," & _
        IIF(bln药房单位, "Decode(E.药品ID,NULL,B.计算单位,E." & str药房单位 & ")", "B.计算单位") & " as 计算单位," & _
        " Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额,D.名称 as 执行部门" & _
        " From " & strSQL & " A,收费项目目录 B,收费项目类别 C,部门表 D,药品规格 E,收费项目别名 F" & _
        " Where Decode(A.记录性质,0,1,A.记录性质)=" & int记录性质 & _
        " And A.NO='" & strNO & "' And A.记录状态 in (0,1) And A.收费细目ID=B.ID" & _
        " And A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And B.ID=E.药品ID(+)" & _
        " And A.收费细目ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=" & IIF(gbln商品名, 3, 1) & _
        IIF(strSQL = "病人费用记录", " And A.登记时间=" & str登记时间, "") & _
        " Group by Nvl(A.价格父号,A.序号),C.名称,Nvl(F.名称,B.名称),B.规格,B.计算单位,D.名称,E.药品ID,E." & str药房单位
        
    strSQL = strSQL & " Order by Nvl(A.价格父号,A.序号)"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)

    With vsDetail
        .Redraw = flexRDNone
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + .FixedRows
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, 0) = rsTmp!类别
                .TextMatrix(i, 1) = rsTmp!项目
                .TextMatrix(i, 2) = FormatEx(rsTmp!数量, 5) & " " & Nvl(rsTmp!计算单位)
                .TextMatrix(i, 3) = Format(rsTmp!单价, "0.00000")
                .TextMatrix(i, 4) = Format(rsTmp!应收金额, gstrDec)
                .TextMatrix(i, 5) = Format(rsTmp!实收金额, gstrDec)
                .TextMatrix(i, 6) = Nvl(rsTmp!执行部门)
                rsTmp.MoveNext
            Next
        End If
        .Row = .FixedRows: .Col = .FixedCols
        .Redraw = flexRDDirect
    End With
    
    Call AppendRows(vsDetail, lnX1, lnY1)
    
    LoadBillDetail = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsMoney_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 2 And mfrmParent.mnuCharge.Visible And mfrmParent.mnuCharge.Enabled Then PopupMenu mfrmParent.mnuCharge
    
End Sub

Private Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant, Optional ByVal lngHideRows As Long = 0) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:补充表格控件的空行
    '参数:objVsf 要补充行的表格控件对象
    '返回:若成功返回True,否则返回 False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    
    On Error GoTo ErrHand
    
    If objVsf.Rows = 0 Then Exit Function
    
    For lngLoop = objVsf.Rows - 1 To 1 Step -1
        If objVsf.RowHidden(lngLoop) = False Then
            lngLastRow = lngLoop
            Exit For
        End If
    Next
    
    lngTop = objVsf.Cell(flexcpTop, lngLastRow, 0) + objVsf.RowHeight(lngLastRow)
    
    '1.隐藏所有的线
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    
    '2.重新计算需要的纵线
    For lngLoop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)

        With objLineY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.重新计算需要的横线
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeight(0)) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeight(0) + 15
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    Exit Function
    
ErrHand:
    
End Function


