VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPACSReq 
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
   Begin VB.Frame fraSplit1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   6
      Top             =   3245
      Width           =   7110
   End
   Begin VB.Frame fraFee 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   3360
      Width           =   7935
      Begin VB.Label lblCash 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   7560
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   " 费用"
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   50
         Width           =   450
      End
   End
   Begin VB.PictureBox picFile 
      Height          =   2055
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   7875
      TabIndex        =   3
      Top             =   0
      Width           =   7935
   End
   Begin VB.Frame fraBill 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   0
      Top             =   4855
      Width           =   7110
   End
   Begin VSFlex8Ctl.VSFlexGrid vsMoney 
      Height          =   1140
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   7185
      _cx             =   12674
      _cy             =   2011
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
   Begin VSFlex8Ctl.VSFlexGrid vsDetail 
      Height          =   1380
      Left            =   0
      TabIndex        =   2
      Top             =   5040
      Width           =   7200
      _cx             =   12700
      _cy             =   2434
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
End
Attribute VB_Name = "frmPACSReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COLOR_LOST = &HFFEBD7
Private Const COLOR_FOCUS = &HFFCC99

Private mrsPrice As ADODB.Recordset '未计费医嘱的主费用
Private objBillForm As Object
Private WithEvents frmParent As Form
Attribute frmParent.VB_VarHelpID = -1

Private pgbLoad As Object
Private AdviceID As Long, lngSendNO As Long
Private iPatientType As Integer, lngPatientID As Long, lngPatientDept As Long
Private lngPageId As Long, strCheckNo As String
Private str计费状态 As String, str费别 As String, int记录性质 As Integer
Private int执行状态 As Integer, strNO As String, lng开单科室ID As Long
Private mstrPrivs As String
Private mblnMoved As Boolean

Public Sub zlRefresh(objParent As Object, ByVal lngAdviceID As Long, ByVal SendNO As Long, _
    ByVal strPrivs As String, Optional objpgbLoad As Object, Optional ByVal blnMoved As Boolean = False)
    
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    If objBillForm Is Nothing Then Exit Sub
    On Error GoTo DBError
    mblnMoved = blnMoved
    
    strSQL = _
        " Select X.记录性质 as 费用性质,X.记录状态 as 费用状态," & _
        " A.医嘱ID,A.发送号,B.相关ID,B.序号,B.诊疗类别,B.诊疗项目ID,A.发送时间 as 时间,A.NO," & _
        " A.记录性质,A.执行状态,A.计费状态,B.病人ID,B.主页ID,B.挂号单,B.病人科室ID,E.名称 as 科室,D.姓名," & _
        " Decode(B.病人来源,1,D.门诊号,2,D.住院号,4,D.门诊号,NULL) as 标识号,Nvl(F.费别,D.费别) as 费别," & _
        " Decode(B.病人来源,1,'门诊',2,'住院',3,'外来',4,'体检') as 来源,C.名称 as 内容,A.执行间,A.执行部门ID" & _
        " From 病人医嘱发送 A,病人医嘱记录 B,诊疗项目目录 C,病人信息 D,部门表 E,病案主页 F,病人费用记录 X" & _
        " Where A.医嘱ID=B.ID And B.诊疗项目ID=C.ID And B.病人ID=D.病人ID" & _
        " And B.病人科室ID=E.ID And B.病人ID=F.病人ID(+) And B.主页ID=F.主页ID(+)" & _
        " And A.NO=X.NO(+) And A.记录性质=Decode(X.记录性质(+),0,1,X.记录性质(+))" & _
        " And X.记录状态(+)<>2 And X.医嘱序号(+)=A.医嘱ID And X.序号(+)=1" & _
        " And A.医嘱ID= [1]  And A.发送号= [2] " & _
        " Order by A.发送时间 Desc,B.病人ID,B.序号"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人费用记录", "H病人费用记录")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngAdviceID, SendNO)
   
    Set frmParent = objParent
    Set pgbLoad = objpgbLoad
    AdviceID = lngAdviceID: lngSendNO = SendNO: iPatientType = 1
    lngPatientID = 0: lngPageId = 0: strCheckNo = "": lngPatientDept = 0
    str计费状态 = "": str费别 = "": int记录性质 = 1: mstrPrivs = strPrivs
    int执行状态 = 0: strNO = "": lng开单科室ID = 0
    If Not rsTmp.EOF Then
        iPatientType = Decode(rsTmp("来源"), "门诊", 1, "体检", 1, 2)
        lngPatientID = rsTmp("病人ID"): lngPageId = Nvl(rsTmp("主页ID"), 0): strCheckNo = Nvl(rsTmp("挂号单"), "")
        lngPatientDept = Nvl(rsTmp("病人科室ID"), 0)
        str计费状态 = GetSendMoneyState(lngAdviceID, SendNO): str费别 = Nvl(rsTmp!费别): int记录性质 = Nvl(rsTmp!记录性质, 1)
        int执行状态 = Nvl(rsTmp!执行状态, 0): strNO = Nvl(rsTmp!NO): lng开单科室ID = Nvl(rsTmp!执行部门ID, 0)
    End If
    ShowMenu
    
    If frmParent.Visible Then
        objBillForm.ShowMe AdviceID, pgbLoad
        Call LoadMoneyList(AdviceID, lngSendNO, str计费状态, str费别, int记录性质)
    Else
        Me.Tag = "Loading":
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlMenuClick(objMenu As Menu)
    Dim strMenu As String
    
    If objMenu.Caption Like "*(&*)*" Then
        strMenu = Split(objMenu.Caption, "(&")(0)
    Else
        strMenu = objMenu.Caption
    End If
    Select Case strMenu
        Case "生成主费用"
                MoneyMain
        Case "修改附加费用"
                MoneyModi
        Case "删除附加费用"
                MoneyDel
        Case "收费单据"
                Call MoneyNewBilling(1)
        Case "记帐单据"
                Call MoneyNewBilling(2)
        Case "零费耗用登记"
                Call MoneyNewBilling(2, True)
    End Select
End Sub

Public Sub zlButtonClick(objButton As MSComctlLib.Button)
    Select Case objButton.Key
        Case "主费"
            MoneyMain
        Case "补费"
            frmParent.PopupMenu frmParent.mnuMoneyFunc(2)
        Case "改费"
            MoneyModi
        Case "删费"
            MoneyDel
    End Select
End Sub

Public Sub zlPrint(ByVal bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    Dim strWidth As String
    Dim objGrid As Object
    
    On Error Resume Next
    If frmParent.lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    '表头
    objOut.Title.Text = "病人费用清单"
    Set objGrid = Me.vsMoney
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    '表上
    With frmParent.lvwPati.SelectedItem
        Set objRow = New zlTabAppRow
        objRow.Add "病人：" & .SubItems(2) & " 来源：" & .Text & " 标识号：" & .SubItems(6)
        objRow.Add "单据：" & .SubItems(1) & " 内容：" & .SubItems(3)
        objOut.UnderAppRows.Add objRow
    End With

    '表下
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow

    '表体
    Set objOut.Body = objGrid

    '输出
    objGrid.Redraw = False
    lngRow = objGrid.Row: lngCol = objGrid.Col

    strWidth = ""
    For i = 0 To objGrid.Cols - 1
        strWidth = strWidth & "," & objGrid.ColWidth(i)
        If i <= objGrid.FixedCols - 1 Or objGrid.ColHidden(i) Then
            objGrid.ColWidth(i) = 0
        End If
    Next

    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If

    strWidth = Mid(strWidth, 2)
    For i = 0 To objGrid.Cols - 1
        objGrid.ColWidth(i) = Split(strWidth, ",")(i)
    Next
    objGrid.Row = lngRow: objGrid.Col = lngCol
    objGrid.Redraw = True
End Sub

Private Sub MoneyMain()
    Dim rsPati As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    
    Dim lng病人ID As Long, lng主页ID As Long, lng发送号, lng医嘱ID As Long
    Dim int来源 As Integer
    Dim int父号 As Integer, lng项目ID As Long, lng执行部门ID As Long
    Dim lng病人病区ID As Long, lng病人科室ID As Long, lng类别ID As Long
    Dim arrSQL As Variant, strSQL As String, strDate As String, i As Long
    Dim int保险项目否 As Integer, lng保险大类ID As Long, str保险编码 As String, cur统筹金额 As Currency
    Dim lng开嘱科室ID As Long, str开嘱医生 As String, int序号 As Integer, strMsg As String
    
    If AdviceID = 0 Then Exit Sub
    If InStr(str计费状态, ",-1,") > 0 Then
        MsgBox "该执行项目无需计费。" & vbCrLf & "如果需要，你可以手工补充附加费用。", vbInformation, gstrSysName
        Exit Sub
    ElseIf InStr(str计费状态, ",1,") > 0 Then
        MsgBox "该执行项目的主费用已经计费。" & vbCrLf & "如果需要，你可以手工补充附加费用。", vbInformation, gstrSysName
        Exit Sub
    End If
    If mrsPrice Is Nothing Then Exit Sub
    If mrsPrice.RecordCount = 0 Then
        MsgBox "该执行项目没有可以计费的主费用。" & vbCrLf & "如果需要，你可以手工补充附加费用。", vbInformation, gstrSysName
        Exit Sub
    End If
    If int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If int记录性质 = 1 Then
        If BillExistBalance(strNO) Then
            MsgBox "单据 " & strNO & " 已经收费，不能再生成这张单据的主费用。" & vbCrLf & "如果需要，你可以手工补充附加费用。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If MsgBox("确实要生成该项目的主费用吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
        
    Screen.MousePointer = 11
    
    lng发送号 = lngSendNO
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
        " Where A.病人ID= [1] And A.病人ID=B.病人ID(+)" & _
        " And B.主页ID(+)= [2] And A.医疗付款方式=C.名称(+)"
    Set rsPati = OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    
    '可能对照费用为药品费用
    If int记录性质 = 1 Then
        lng类别ID = ExistIOClass(8) '门诊划价单
    Else
        lng类别ID = ExistIOClass(9) '门诊/住院记帐单
    End If
    
    '可能发送时已自动生成了部份主费用,现在是手工生成剩余部份。
    '1.因为单据号相同,所以要保持序号连续
    '2.如果是生成收费划价单，要保证一张单据中登记时间相同(不然收费无法处理)
    '3.第2点的情况，如果部份主费用已经收费，则不允许再生成主费用
    int序号 = GetBillMax序号(strNO, int记录性质, strDate)
    If int记录性质 = 2 Or strDate = "" Then
        strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    Else
        strDate = "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
    End If
    
    arrSQL = Array()
    With mrsPrice
        .MoveFirst
        For i = 1 To .RecordCount
            '获取对应的医嘱信息
            If lng医嘱ID <> !医嘱ID Then
                strSQL = "Select 医嘱期效,病人科室ID,开嘱科室ID,开嘱医生,婴儿,执行频次,计价特性" & _
                    " From 病人医嘱记录 Where ID= [1] "
                Set rsAdvice = OpenSQLRecord(strSQL, Me.Caption, CLng(!医嘱ID))
                
                '将当前这条计费医嘱标记为已计费
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱发送_计费(" & !医嘱ID & "," & lng发送号 & ")"
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
            
            '开单科室及开单人
            lng开嘱科室ID = rsAdvice!开嘱科室ID
            str开嘱医生 = rsAdvice!开嘱医生
            
            '每个收费项目的处理
            If lng项目ID <> !收费细目ID Then
                int父号 = int序号 '获取价格父号
                lng执行部门ID = Get收费执行科室ID(!类别, !收费细目ID, !执行科室, Nvl(rsAdvice!病人科室ID, 0), int来源)
                            
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
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If int来源 = 1 Then
                If int记录性质 = 1 Then
                    '生成门诊划价单据
                    arrSQL(UBound(arrSQL)) = _
                        "zl_门诊划价记录_Insert('" & strNO & "'," & int序号 & "," & lng病人ID & ",NULL," & _
                        ZVal(Nvl(rsPati!门诊号, 0)) & ",'" & Nvl(rsPati!付款码) & "','" & Nvl(rsPati!姓名) & "'," & _
                        "'" & Nvl(rsPati!性别) & "','" & Nvl(rsPati!年龄) & "','" & Nvl(rsPati!费别) & "',NULL," & _
                        lng病人病区ID & "," & lng病人科室ID & "," & lng开嘱科室ID & ",'" & str开嘱医生 & "'," & _
                        "NULL," & lng项目ID & ",'" & !类别 & "','" & !计算单位 & "',NULL,1," & !数量 & "," & _
                        !附加手术 & "," & ZVal(lng执行部门ID) & "," & IIf(int父号 = int序号, "NULL", int父号) & "," & _
                        !收入项目ID & ",'" & Nvl(!收据费目) & "'," & !单价 & "," & !应收 & "," & !实收 & "," & _
                        strDate & "," & strDate & ",NULL,'" & UserInfo.姓名 & "'," & ZVal(lng类别ID) & ",NULL," & _
                        !医嘱ID & ",'" & Nvl(rsAdvice!执行频次) & "',NULL,NULL," & Nvl(rsAdvice!医嘱期效, 0) & "," & _
                        Nvl(rsAdvice!计价特性, 0) & ",1)"
                Else
                    '生成门诊记帐单据
                    arrSQL(UBound(arrSQL)) = _
                        "zl_门诊记帐记录_Insert('" & strNO & "'," & int序号 & "," & lng病人ID & "," & _
                        ZVal(Nvl(rsPati!门诊号, 0)) & ",'" & Nvl(rsPati!姓名) & "','" & Nvl(rsPati!性别) & "'," & _
                        "'" & Nvl(rsPati!年龄) & "','" & Nvl(rsPati!费别) & "',NULL," & ZVal(Nvl(rsAdvice!婴儿, 0)) & "," & _
                        lng病人病区ID & "," & lng病人科室ID & "," & lng开嘱科室ID & "," & _
                        "'" & str开嘱医生 & "',NULL," & lng项目ID & ",'" & !类别 & "'," & _
                        "'" & !计算单位 & "',1," & !数量 & "," & !附加手术 & "," & ZVal(lng执行部门ID) & "," & _
                        IIf(int父号 = int序号, "NULL", int父号) & "," & !收入项目ID & ",'" & Nvl(!收据费目) & "'," & !单价 & "," & _
                        !应收 & "," & !实收 & "," & strDate & "," & strDate & ",NULL,NULL,'" & UserInfo.编号 & "'," & _
                        "'" & UserInfo.姓名 & "'," & ZVal(lng类别ID) & ",NULL,NULL," & !医嘱ID & "," & _
                        "'" & Nvl(rsAdvice!执行频次) & "',NULL,NULL," & Nvl(rsAdvice!医嘱期效, 0) & "," & _
                        Nvl(rsAdvice!计价特性, 0) & ")"
                End If
            Else
                '生成住院记帐单据
                arrSQL(UBound(arrSQL)) = _
                    "zl_住院记帐记录_Insert('" & strNO & "'," & int序号 & "," & lng病人ID & "," & ZVal(lng主页ID) & "," & _
                    ZVal(Nvl(rsPati!住院号, 0)) & ",'" & Nvl(rsPati!姓名) & "','" & Nvl(rsPati!性别) & "'," & _
                    "'" & Nvl(rsPati!年龄) & "'," & ZVal(Nvl(rsPati!床号)) & ",'" & Nvl(rsPati!费别) & "'," & _
                    lng病人病区ID & "," & lng病人科室ID & ",NULL," & ZVal(Nvl(rsAdvice!婴儿, 0)) & "," & _
                    lng开嘱科室ID & ",'" & str开嘱医生 & "',NULL," & lng项目ID & ",'" & !类别 & "'," & _
                    "'" & !计算单位 & "'," & int保险项目否 & "," & ZVal(lng保险大类ID) & ",'" & str保险编码 & "'," & _
                    "1," & !数量 & "," & !附加手术 & "," & ZVal(lng执行部门ID) & "," & _
                    IIf(int父号 = int序号, "NULL", int父号) & "," & !收入项目ID & ",'" & Nvl(!收据费目) & "'," & !单价 & "," & _
                    !应收 & "," & !实收 & "," & cur统筹金额 & "," & strDate & "," & strDate & ",NULL,NULL," & _
                    "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',NULL," & ZVal(lng类别ID) & ",NULL,NULL,NULL," & _
                    !医嘱ID & ",'" & Nvl(rsAdvice!执行频次) & "',NULL,NULL," & Nvl(rsAdvice!医嘱期效, 0) & "," & _
                    Nvl(rsAdvice!计价特性, 0) & ",NULL)"
            End If
            
            int序号 = int序号 + 1
            
            .MoveNext
        Next
    End With
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call ExecuteProc(arrSQL(i), Me.Caption)
    Next
    
    '在提交前进行医保传输
    If int来源 = 2 And Not IsNull(rsPati!险类) Then
        If gclsInsure.GetCapability(support记帐上传, , rsPati!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, , rsPati!险类) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strMsg, , rsPati!险类) Then
                gcnOracle.RollbackTrans
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                Screen.MousePointer = 0: Exit Sub
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
    Me.Tag = "Loading": Call Form_Activate
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub MoneyModi()
    Dim lng病人ID As Long, lng主页ID As Long
    Dim lng病人科室ID As Long
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim int病人来源 As Integer, int记录性质 As Integer
    Dim strNO As String, bln零耗 As Boolean
    
    If AdviceID = 0 Then Exit Sub
    If vsMoney.TextMatrix(vsMoney.Row, 0) = "主费用" Then
        MsgBox "执行项目的主费用不能修改。如果需要，你可以手工补充附加费用。", vbInformation, gstrSysName
        Exit Sub
    End If
    If int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsMoney
        strNO = .TextMatrix(.Row, 2)
        If strNO = "" Or strNO = "[未计费]" Then Exit Sub
        int记录性质 = .Cell(flexcpData, .Row, 1)
        
        If InStr(.TextMatrix(.Row, 1), "√") > 0 Then
            MsgBox "该单据已经收费，不能再修改。", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    lng医嘱ID = AdviceID
    lng发送号 = lngSendNO
    lng病人ID = lngPatientID
    lng主页ID = lngPageId
    int病人来源 = iPatientType
    lng病人科室ID = lngPatientDept
    
    If EditExpense(Me, 0, int记录性质, mstrPrivs, strNO, lng医嘱ID, lng发送号, lng病人ID, lng主页ID, _
        int病人来源, lng开单科室ID, lng病人科室ID, False) Then
        '刷新
        Me.Tag = "Loading": Call Form_Activate
    End If
End Sub

Private Sub MoneyDel()
    Dim int病人来源 As Integer, int记录性质 As Integer
    Dim strNO As String
    
    If AdviceID = 0 Then Exit Sub
    If vsMoney.TextMatrix(vsMoney.Row, 0) = "主费用" Then
        MsgBox "执行项目的主费用不能删除。", vbInformation, gstrSysName
        Exit Sub
    End If
    If int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsMoney
        strNO = .TextMatrix(.Row, 2)
        If strNO = "" Or strNO = "[未计费]" Then Exit Sub
        int记录性质 = .Cell(flexcpData, .Row, 1)
    
        If InStr(.TextMatrix(.Row, 1), "√") > 0 Then
            MsgBox "该单据已经收费，不能再删除。", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    int病人来源 = iPatientType
    
    If EditExpense(Me, 3, int记录性质, mstrPrivs, strNO, 0, 0, 0, 0, _
        int病人来源, 0, 0) Then
        '刷新
        Me.Tag = "Loading": Call Form_Activate
    End If
End Sub

Private Sub MoneyNewBilling(ByVal iRecordType As Integer, Optional OnlyRecord As Boolean = False)
    Dim lng病人ID As Long, lng主页ID As Long
    Dim lng病人科室ID As Long
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim int病人来源 As Integer
    
    If AdviceID = 0 Then Exit Sub
    
    If int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lng医嘱ID = AdviceID
    lng发送号 = lngSendNO
    lng病人ID = lngPatientID
    lng主页ID = lngPageId
    int病人来源 = iPatientType
    lng病人科室ID = lngPatientDept
    
    If EditExpense(Me, 0, iRecordType, mstrPrivs, "", lng医嘱ID, lng发送号, lng病人ID, lng主页ID, _
        int病人来源, lng开单科室ID, lng病人科室ID, OnlyRecord) Then
        '刷新
        Me.Tag = "Loading": Call Form_Activate
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Me.Tag = "Loading" Then
        frmParent.Refresh: Me.Tag = ""
        pgbLoad.Visible = True
        objBillForm.ShowMe AdviceID, pgbLoad
        Call LoadMoneyList(AdviceID, lngSendNO, str计费状态, str费别, int记录性质)
        pgbLoad.Visible = False
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ShowError
    
    Set mrsPrice = Nothing
    
    Set objBillForm = getRequestForm
    
    SetWindowLong objBillForm.Hwnd, GWL_STYLE, WS_CHILD
    objBillForm.Show , Me
    SetParent objBillForm.Hwnd, picFile.Hwnd
    
    InitMoneyTable
    InitDetailTable
    Exit Sub
ShowError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    On Error Resume Next
    With Me.fraSplit1
        .Left = 0: .Width = Me.ScaleWidth
        .Top = Me.ScaleHeight - (Me.vsDetail.Top + Me.vsDetail.Height - .Top)
    End With
    With Me.fraBill
        .Left = 0: .Width = Me.ScaleWidth
        .Top = Me.ScaleHeight - (Me.vsDetail.Top + Me.vsDetail.Height - .Top)
    End With
    
    With picFile
        .Top = 0: .Left = 0
        .Width = Me.ScaleWidth: .Height = Me.fraSplit1.Top - .Top
    End With
    With fraFee
        .Left = 0: .Top = Me.fraSplit1.Top + Me.fraSplit1.Height
        .Width = Me.ScaleWidth
        
        lblCash.Left = .Width - lblCash.Width
    End With
    With Me.vsMoney
        .Left = 0: .Top = Me.fraFee.Top + Me.fraFee.Height
        .Width = Me.ScaleWidth: .Height = Me.fraBill.Top - .Top
    End With
    With Me.vsDetail
        .Left = 0: .Top = Me.fraBill.Top + Me.fraBill.Height
        .Width = Me.ScaleWidth: .Height = Me.ScaleHeight - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload objBillForm
    Set objBillForm = Nothing
    
    Set mrsPrice = Nothing
End Sub

Private Sub fraBill_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    With fraBill
        .BackColor = RGB(0, 0, 0)
        If .Top + y - Me.vsMoney.Top < 1000 Then
            .Top = Me.vsMoney.Top + 1000
        ElseIf Me.ScaleHeight - .Top - y < 1000 Then
            .Top = Me.ScaleHeight - 1000
        Else
            .Top = .Top + y
        End If
    End With
End Sub

Private Sub fraBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub

    fraBill.BackColor = Me.BackColor
    Form_Resize
End Sub

Private Sub fraSplit1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    With fraSplit1
        .BackColor = RGB(0, 0, 0)
        If .Top + y - Me.picFile.Top < 3000 Then
            .Top = Me.picFile.Top + 3000
        ElseIf Me.ScaleHeight - .Top - y < 2100 Then
            .Top = Me.ScaleHeight - 2100
        Else
            .Top = .Top + y
        End If
    End With
End Sub

Private Sub fraSplit1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub

    fraSplit1.BackColor = Me.BackColor
    Form_Resize
End Sub

Private Sub frmParent_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub picFile_Resize()
    Dim vRect As RECT
    
    On Error Resume Next
    If Not objBillForm Is Nothing Then
        MoveWindow objBillForm.Hwnd, 0, 0, picFile.ScaleWidth / Screen.TwipsPerPixelX, picFile.ScaleHeight / Screen.TwipsPerPixelY, 1
        Call GetWindowRect(objBillForm.Hwnd, vRect)
        SetWindowPos objBillForm.Hwnd, 0, 0, 0, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED
    End If
End Sub

Private Sub vsMoney_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = OldRow Then Exit Sub
    If NewCol >= vsMoney.FixedCols And NewRow >= vsMoney.FixedRows Then
        vsMoney.ForeColorSel = vsMoney.Cell(flexcpForeColor, NewRow, 0)
        Call LoadBillDetail(NewRow)
    End If
End Sub

Private Sub vsMoney_DblClick()
    If vsMoney.MouseRow >= vsMoney.FixedRows Then
        Call vsMoney_KeyPress(13)
    End If
End Sub

Private Sub vsMoney_GotFocus()
    vsMoney.BackColorSel = COLOR_FOCUS
End Sub

Private Sub vsMoney_KeyPress(KeyAscii As Integer)
    Dim int病人来源 As Integer, int记录性质 As Integer
    Dim strNO As String
    
    If KeyAscii = 13 Or KeyAscii = 32 Then
        KeyAscii = 0
        With vsMoney
            strNO = .TextMatrix(.Row, 2)
            If strNO = "" Or strNO = "[未计费]" Then Exit Sub
            int记录性质 = .Cell(flexcpData, .Row, 1)
        End With
        int病人来源 = iPatientType
        
        Call EditExpense(Me, 1, int记录性质, mstrPrivs, strNO, 0, 0, 0, 0, _
            int病人来源, 0, 0, False)
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
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "费用类型,900,1;单据类型,1000,1;单据号,900,1;费别,900,1;应收金额,1000,7;实收金额,1000,7;开单科室,1000,1;开单人,750,1;登记时间,1080,1;登记人,750,1"
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
    End With
End Sub

Private Sub InitDetailTable()
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "类别,650,1;项目,2000,1;数量,1000,1;单价,1000,7;应收金额,1000,7;实收金额,1000,7;执行科室,1000,1"
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
    End With
End Sub

Private Function LoadMoneyList(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, _
    ByVal str计费状态 As String, ByVal str费别 As String, ByVal int记录性质 As Integer) As Boolean
'功能：读取指定医嘱的主要费用及附加费用
'说明：1.包含医嘱本身的主费用及附加费用,主费用可能尚未产生
'      2.目前单据暂不支持部份退费,所以清单中只需简单显示
    Dim rsTmp As New ADODB.Recordset
    Dim rsList As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim blnMain As Boolean, blnSub As Boolean
    Dim strPre As String, lngRow As Long
    Dim cur应收 As Currency, cur实收 As Currency
    Dim str医嘱ID As String, lngMain As Long
    Dim blnCash As Boolean  '是已经收费，只要存在记帐记录或一条为收费的记录，则认为未收费；
    
    On Error GoTo errH

    Set mrsPrice = Nothing
    '存在未计费状态,直接读取收费关系显示
    If InStr(str计费状态, ",0,") > 0 Then
        Call LoadAdvicePrice(lng医嘱ID, lng发送号, str费别)
        If mrsPrice.RecordCount > 0 Then
            For i = 1 To mrsPrice.RecordCount
                cur应收 = cur应收 + Nvl(mrsPrice!应收, 0)
                cur实收 = cur实收 + Nvl(mrsPrice!实收, 0)
                mrsPrice.MoveNext
            Next
            
            strSQL = "Select B.名称 as 开嘱科室,开嘱医生 From 病人医嘱记录 A,部门表 B Where A.开嘱科室ID=B.ID And A.ID= [1] "
            If mblnMoved Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
            If Not rsTmp.EOF Then
                strSQL = "Select " & _
                    " 1 as 费用类型," & int记录性质 & " as 记录性质,0 as 已收费,'[未计费]' as NO," & _
                    "'" & str费别 & "' as 费别," & cur应收 & " as 应收金额," & cur实收 & " as 实收金额," & _
                    "'" & Nvl(rsTmp!开嘱科室) & "' as 开单科室,'" & Nvl(rsTmp!开嘱医生) & "' as 开单人," & _
                    " Sysdate as 登记时间,'" & UserInfo.姓名 & "' as 操作员" & _
                    " From Dual"
            Else
                strSQL = ""
            End If
        End If
    End If
    
    '存在已计费状态,应该可以直接读取主费用部份(只有一张单据,可能含其它医嘱费用;已删除或退费销帐,则不显示)
    If InStr(str计费状态, ",1,") > 0 Then
        '包含检查部位，附加手术，检验组合的费用
        str医嘱ID = _
            " Select ID From 病人医嘱记录 Where ID= [1] Union All" & _
            " Select ID From 病人医嘱记录 Where 相关ID= [1] And 诊疗类别 IN('F','D')"
        strSQL = strSQL & IIf(strSQL <> "", " Union ALL ", "") & _
            " Select 1 as 费用类型,A.记录性质,Decode(B.记录状态,1,1,0) as 已收费," & _
            " A.NO,B.费别,Sum(B.应收金额) as 应收金额,Sum(B.实收金额) as 实收金额," & _
            " C.名称 as 开单科室,B.开单人,B.登记时间,Nvl(B.操作员姓名,B.划价人) as 操作员" & _
            " From 病人医嘱发送 A,病人费用记录 B,部门表 C" & _
            " Where A.医嘱ID IN(" & str医嘱ID & ") And A.发送号= [2] " & _
            " And A.NO=B.NO And A.记录性质=B.记录性质 And A.医嘱ID=B.医嘱序号+0" & _
            " And B.记录状态 IN(0,1) And B.开单部门ID=C.ID" & _
            " Group by A.记录性质,B.记录状态,A.NO,B.费别,C.名称,B.开单人,B.登记时间,Nvl(B.操作员姓名,B.划价人)"
    End If
    
    '附费用部份(已删除或退费销帐,则不显示)
    '医嘱ID对主费用有用,对附加费用都是相同的这个
    strSQL = strSQL & IIf(strSQL <> "", " Union ALL ", "") & _
        " Select 2 as 费用类型,A.记录性质,Decode(B.记录状态,1,1,0) as 已收费," & _
        " A.NO,B.费别,Sum(B.应收金额) as 应收金额,Sum(B.实收金额) as 实收金额," & _
        " C.名称 as 开单科室,B.开单人,B.登记时间,Nvl(B.操作员姓名,B.划价人) as 操作员" & _
        " From 病人医嘱附费 A,病人费用记录 B,部门表 C" & _
        " Where A.医嘱ID= [1]  And A.发送号= [2] " & _
        " And A.NO=B.NO And A.记录性质=B.记录性质 And A.医嘱ID=B.医嘱序号+0" & _
        " And B.记录状态 IN(0,1) And B.开单部门ID=C.ID" & _
        " Group by A.记录性质,B.记录状态,A.NO,B.费别,C.名称,B.开单人,B.登记时间,Nvl(B.操作员姓名,B.划价人)"
        
    strSQL = "Select * From (" & strSQL & ") Order by 费用类型,登记时间 Desc"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人费用记录", "H病人费用记录")
        strSQL = Replace(strSQL, "病人医嘱附费", "H病人医嘱附费")
    End If
    Set rsList = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
    With vsMoney
        lngRow = .FixedRows
        strPre = .TextMatrix(.Row, 1) & "_" & .TextMatrix(.Row, 2)
        
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
        vsDetail.Rows = vsDetail.FixedRows
        vsDetail.Rows = vsDetail.FixedRows + 1
        
        blnCash = True
        If Not rsList.EOF Then
            .Rows = rsList.RecordCount + 1
            For i = 1 To rsList.RecordCount
                If rsList!已收费 = 0 And Nvl(rsList!实收金额, 0) <> 0 And rsList!NO <> "[未计费]" Then
                    blnCash = False '要管未审核的记帐划价单,不管尚未计费的主费用
                End If
                
                .TextMatrix(i, 0) = IIf(rsList!费用类型 = 1, "主费用", "附加费用")
                
                '是否该条已收费(仅突出收费单据)
                .TextMatrix(i, 1) = IIf(rsList!记录性质 = 1, "收费单据" & IIf(rsList!已收费 = 1, "√", ""), "记帐单据")
                If rsList!记录性质 = 1 And rsList!已收费 = 1 Then '已收费蓝色显示
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HC00000 '深蓝
                End If
                .TextMatrix(i, 2) = rsList!NO
                .TextMatrix(i, 3) = Nvl(rsList!费别)
                .TextMatrix(i, 4) = Format(Nvl(rsList!应收金额, 0), gstrDec)
                .TextMatrix(i, 5) = Format(Nvl(rsList!实收金额, 0), gstrDec)
                .TextMatrix(i, 6) = Nvl(rsList!开单科室)
                .TextMatrix(i, 7) = Nvl(rsList!开单人)
                .TextMatrix(i, 8) = Format(rsList!登记时间, "MM-dd HH:mm")
                .TextMatrix(i, 9) = Nvl(rsList!操作员)
                                                
                '附加数据
                .Cell(flexcpData, i, 1) = CInt(rsList!记录性质)
                .Cell(flexcpData, i, 8) = Format(rsList!登记时间, "yyyy-MM-dd HH:mm:ss")
                                                
                '是否产生冻结行
                If rsList!费用类型 = 1 Then blnMain = True
                If rsList!费用类型 = 2 Then
                    If blnMain Then lngMain = i - 1
                    blnSub = True
                End If
                
                '定位到原行位置
                If .TextMatrix(i, 1) & "_" & .TextMatrix(i, 2) = strPre Then lngRow = i
                rsList.MoveNext
            Next
        Else
            blnCash = False
        End If
        If blnMain And blnSub Then
            .FrozenRows = lngMain
            .Select lngMain, 0, lngMain, .Cols - 1
            .CellBorder &HC00000, 0, 0, 0, 1, 0, 0
        End If
        .Row = lngRow: .Col = .FixedCols
        Call .ShowCell(.Row, .Col)
        
        .Redraw = flexRDDirect
        Call vsMoney_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    LoadMoneyList = True
    
    Me.lblCash.Caption = IIf(blnCash, "收", "")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadAdvicePrice(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByVal str费别 As String) As Boolean
'功能：读取指定医嘱的计价关系到临时记录集
'说明：要计算的项目应该不是叮嘱,院外执行,无需计费
    Dim rsTmp As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    Dim dbl数量 As Double, bln附加手术 As Boolean
    Dim strSQL As String, str医嘱ID As String
    Dim i As Long, j As Long
    
    Set mrsPrice = New ADODB.Recordset
    mrsPrice.Fields.Append "医嘱ID", adBigInt
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
    
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
    
    On Error GoTo errH
            
    '读取要计算主费用的医嘱记录(包含附加手术,检查部位；手术麻醉单独)
    '包含检查部位，附加手术，检验组合的费用
    str医嘱ID = _
        " Select ID From 病人医嘱记录 Where ID= [1] Union All" & _
        " Select ID From 病人医嘱记录 Where 相关ID= [1] And 诊疗类别 IN('F','D')"
    strSQL = _
        " Select B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID," & _
        " Nvl(A.发送数次,Sum(Nvl(C.本次数次,0))) as 数量" & _
        " From 病人医嘱发送 A,病人医嘱记录 B,病人医嘱执行 C" & _
        " Where Nvl(A.计费状态,0)=0 And B.ID IN(" & str医嘱ID & ")" & _
        " And A.医嘱ID=B.ID And A.发送号+0= [2] " & _
        " And C.医嘱ID(+)=A.医嘱ID And C.发送号(+)=A.发送号" & _
        " Group by B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID,A.发送数次" & _
        " Order by 序号"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱执行", "H病人医嘱执行")
    End If
    Set rsAdvice = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
    
    For i = 1 To rsAdvice.RecordCount
        dbl数量 = Nvl(rsAdvice!数量, 0)
        
        '读取对应的收费价目:只读取固定对照,且不是变价的对照
        bln附加手术 = (rsAdvice!诊疗类别 = "F" And Not IsNull(rsAdvice!相关ID))
        strSQL = IIf(bln附加手术, "Nvl(B.附术收费率,100)/100", "1") & " as 附术率"
        strSQL = _
            " Select A.收费项目ID,A.收费数量,B.收入项目ID,D.收据费目,C.类别," & _
            " C.计算单位,C.执行科室,Decode(C.是否变价,1,NULL,B.现价) as 单价," & strSQL & _
            " From 诊疗收费关系 A,收费价目 B,收费项目目录 C,收入项目 D" & _
            " Where A.诊疗项目ID= [1] " & _
            " And A.收费项目ID=B.收费细目ID And A.收费项目ID=C.ID And B.收入项目ID=D.ID" & _
            " And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
            " And Nvl(A.固有对照,0)=1 And Nvl(C.是否变价,0)=0"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, CLng(rsAdvice!诊疗项目ID))
        
        For j = 1 To rsTmp.RecordCount
            mrsPrice.AddNew
            mrsPrice!医嘱ID = rsAdvice!医嘱ID
            mrsPrice!类别 = rsTmp!类别
            mrsPrice!收费细目ID = rsTmp!收费项目ID
            mrsPrice!计算单位 = Nvl(rsTmp!计算单位)
            mrsPrice!附加手术 = IIf(bln附加手术, 1, 0)
            mrsPrice!执行科室 = Nvl(rsTmp!执行科室, 0)
            mrsPrice!收入项目ID = rsTmp!收入项目ID
            mrsPrice!收据费目 = rsTmp!收据费目
            mrsPrice!单价 = Format(Nvl(rsTmp!单价, 0), "0.00000")
            mrsPrice!数量 = Format(Nvl(rsTmp!收费数量, 0) * dbl数量, "0.00000")
            mrsPrice!应收 = Format(mrsPrice!数量 * mrsPrice!单价 * rsTmp!附术率, gstrDec)
            If str费别 = "" Then
                mrsPrice!实收 = mrsPrice!应收
            Else
                mrsPrice!实收 = Format(ActualMoney(str费别, mrsPrice!收入项目ID, mrsPrice!应收), gstrDec)
            End If
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
    Set mrsPrice = Nothing
End Function

Private Function LoadBillDetail(ByVal lngRow As Long) As Boolean
'功能：显示单据明细内容
'参数：lngRow=单据清单行
    Dim rsTmp As New ADODB.Recordset
    Dim strNO As String, int记录性质 As Integer
    Dim lng病人科室ID As Long, int来源 As Integer
    Dim lng项目ID As Long, int父号 As Integer
    Dim lng医嘱ID As Long, str登记时间 As String
    Dim strSQL As String, strIDs As String, strIDsSql As String, i As Long
    
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
        If strNO = "" Then Exit Function
    
        '登记时间是为了区分同时存在审核与未审核的情况
        str登记时间 = .Cell(flexcpData, lngRow, 8)
    End With
    
    '医嘱ID对主费用有用,对附加费用都是相同的这个
    lng医嘱ID = AdviceID
    
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
            .MoveFirst
            For i = 1 To .RecordCount
                If lng项目ID <> !收费细目ID Then int父号 = i
                strSQL = strSQL & IIf(strSQL <> "", " Union ALL ", "") & _
                    " Select " & i & " as 序号," & IIf(int父号 = i, "-NULL", int父号) & " as 价格父号," & _
                    "'" & strNO & "' as NO," & int记录性质 & " as 记录性质,1 as 记录状态," & _
                    !医嘱ID & " as 医嘱序号,'" & !类别 & "' as 收费类别," & !收费细目ID & " as 收费细目ID," & _
                    Get收费执行科室ID(!类别, !收费细目ID, !执行科室, lng病人科室ID, int来源) & " as 执行部门ID," & _
                    !收入项目ID & " as 收入项目ID,1 as 付数," & !数量 & " as 数次," & !单价 & " as 标准单价," & _
                    !应收 & " as 应收金额," & !实收 & " as 实收金额,To_Date('" & str登记时间 & "','YYYY-MM-DD HH24:MI:SS') as 登记时间 From Dual"
                    
                lng项目ID = !收费细目ID
                strIDs = strIDs & "," & !医嘱ID
                .MoveNext
            Next
            If strSQL = "" Then Exit Function
            strSQL = "(" & strSQL & ")"
            strIDs = Mid(strIDs, 2)  '取检验组合中涉及的医嘱ID
        End With
    Else
        strSQL = "病人费用记录"
        '包含检查部位，附加手术，检验组合的费用
        strIDsSql = _
            " Select ID From 病人医嘱记录 Where ID=" & lng医嘱ID & " Union All" & _
            " Select ID From 病人医嘱记录 Where 相关ID=" & lng医嘱ID & " And 诊疗类别 IN('F','D') "
    End If
    
    strSQL = "Select C.名称 as 类别,Nvl(F.名称,B.名称)||Decode(B.规格,NULL,NULL,' '||B.规格) as 项目," & _
        " Sum(A.标准单价" & IIf(bln药房单位, "*Nvl(E." & str药房包装 & ",1)", "") & ") as 单价," & _
        " Avg(Nvl(A.付数,1)*A.数次" & IIf(bln药房单位, "/Nvl(E." & str药房包装 & ",1)", "") & ") as 数量," & _
        IIf(bln药房单位, "Decode(E.药品ID,NULL,B.计算单位,E." & str药房单位 & ")", "B.计算单位") & " as 计算单位," & _
        " Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额,D.名称 as 执行部门" & _
        " From " & strSQL & " A,收费项目目录 B,收费项目类别 C,部门表 D,药品规格 E,收费项目别名 F" & _
        " Where A.记录性质= [1] And A.记录状态 IN(0,1)" & _
        " And A.收费细目ID=B.ID And A.收费类别=C.编码 And A.执行部门ID=D.ID(+)" & _
        " And B.ID=E.药品ID(+) And A.收费细目ID=F.收费细目ID(+)" & _
        " And F.码类(+)=1 And F.性质(+)=" & IIf(gbln商品名, 3, 1) & _
        " And A.NO= [3] And " & _
        IIf(strIDsSql <> "", " a.医嘱序号 in (" & strIDsSql & ")", " instr([4],','||A.医嘱序号||',')> 0 ") & _
        " And A.登记时间= [2] " & _
        " Group by Nvl(A.价格父号,A.序号),C.名称,Nvl(F.名称,B.名称),B.规格,B.计算单位,D.名称,E.药品ID,E." & str药房单位 & _
        " Order by Nvl(A.价格父号,A.序号)"
        
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人费用记录", "H病人费用记录")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, int记录性质, CDate(Format(str登记时间, "yyyy-MM-dd hh:mm:ss")), strNO, "," & strIDs & ",")
    
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
    
    LoadBillDetail = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowMenu()
    frmParent.mnuMoneyAdd(0).Visible = iPatientType = 1
End Sub

Private Sub vsMoney_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 2 And frmParent.mnuMoney.Visible And frmParent.mnuMoney.Enabled Then PopupMenu frmParent.mnuMoney, 2
End Sub
