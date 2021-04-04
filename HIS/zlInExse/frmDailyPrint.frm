VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDailyPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "一日清单打印"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdPrintAll 
      Caption         =   "一次打印所有病人(&M)"
      Height          =   350
      Left            =   4110
      TabIndex        =   18
      Top             =   4635
      Width           =   1905
   End
   Begin VB.CommandButton cmdPreviewAll 
      Caption         =   "一次预览所有病人(&A)"
      Height          =   350
      Left            =   4110
      TabIndex        =   16
      Top             =   4215
      Width           =   1905
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "打印设置(&S)"
      Height          =   350
      Left            =   120
      TabIndex        =   21
      Top             =   4635
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "按病人分开打印(&P)"
      Height          =   350
      Left            =   2220
      TabIndex        =   17
      Top             =   4635
      Width           =   1680
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   6195
      TabIndex        =   19
      Top             =   4635
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   20
      Top             =   4215
      Width           =   1275
   End
   Begin VB.Frame fraDetail 
      Height          =   4095
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   7305
      Begin VB.CheckBox chkReCharge 
         Caption         =   "包含退费(&R)"
         Height          =   195
         Left            =   5640
         TabIndex        =   5
         Top             =   308
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.CheckBox chkZeroFee 
         Caption         =   "包含零费用(&Z)"
         Height          =   195
         Left            =   5640
         TabIndex        =   6
         Top             =   653
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.OptionButton opttime 
         Caption         =   "发生时间(&H)"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   3600
         TabIndex        =   4
         Top             =   660
         Width           =   1620
      End
      Begin VB.OptionButton opttime 
         Caption         =   "登记时间(&D)"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   3600
         TabIndex        =   3
         Top             =   315
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1230
         Width           =   2070
      End
      Begin VB.CommandButton cmdNoPati 
         Caption         =   "全清(&R)"
         Height          =   330
         Left            =   270
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   3555
         Width           =   870
      End
      Begin VB.CommandButton cmdAllPati 
         Caption         =   "全选(&A)"
         Height          =   330
         Left            =   270
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   3180
         Width           =   870
      End
      Begin VB.CheckBox chkPatiType 
         Caption         =   "医保病人(&M)"
         Height          =   195
         Index           =   0
         Left            =   3720
         TabIndex        =   9
         Top             =   1290
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.CheckBox chkPatiType 
         Caption         =   "非医保病人(&N)"
         Height          =   195
         Index           =   1
         Left            =   5640
         TabIndex        =   10
         Top             =   1290
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1215
         TabIndex        =   1
         Top             =   255
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   84279299
         CurrentDate     =   37953
      End
      Begin MSComctlLib.ListView lvwPati 
         Height          =   2340
         Left            =   1215
         TabIndex        =   12
         Top             =   1590
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   4128
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "姓名"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "住院号"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "床号"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "住院医师"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "费别"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "护理等级"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "科室"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "入院日期"
            Object.Width           =   2857
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1215
         TabIndex        =   2
         Top             =   600
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   84279299
         CurrentDate     =   37953
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院病人(&I)"
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   1665
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费用时间(&T)"
         Height          =   180
         Left            =   180
         TabIndex        =   0
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院病区(&U)"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   1290
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   7200
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   105
         X2              =   7200
         Y1              =   1095
         Y2              =   1095
      End
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "按病人分开预览(&V)"
      Height          =   350
      Left            =   2220
      TabIndex        =   15
      Top             =   4215
      Width           =   1680
   End
End
Attribute VB_Name = "frmDailyPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mlng病区ID As Long 'IN
Public mlng病人ID As Long 'IN
Public mstrPrivs As String

Private mlngModul As Long

Private Sub cboUnit_Click()
'功能：读取指定范围内的病人列表
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem, strSql As String
    Dim i As Integer, j As Integer
    Dim intBedLen As Integer, str病人IDs As String, lng病区ID As Long
        
    lvwPati.ListItems.Clear
    
    On Error GoTo errH
    If cboUnit.ListIndex <> -1 Then
        lng病区ID = cboUnit.ItemData(cboUnit.ListIndex)
        intBedLen = GetMaxBedLen(lng病区ID, False)
    End If
    strSql = _
        "Select A.病人ID,B.主页ID,Nvl(b.姓名, a.姓名) As 姓名,B.住院号,LPAD(B.出院病床," & intBedLen & ",' ') as 床号," & _
        "       B.住院医师,B.费别,D.名称 as 护理等级,C.名称 as 科室,B.入院日期,B.险类,B.病人类型" & _
        " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 D,在院病人　E" & _
        " Where A.病人ID=B.病人ID And B.主页ID=A.主页ID And B.出院科室ID=C.ID" & _
        " And A.病人ID=E.病人ID   And B.护理等级ID=D.ID(+)" & _
        IIf(chkPatiType(0).Value = 0, " And B.险类 Is Null", "") & _
        IIf(chkPatiType(1).Value = 0, " And B.险类 Is Not Null", "") & _
        IIf(lng病区ID > 0, " And B.当前病区ID=[1] And E.病区ID=[1] ", "") & _
        IIf(lng病区ID = 0, " Order by B.住院号 Desc", " Order by LPAD(床号,10,' ')")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病区ID)
  
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID, rsTmp!姓名)
        objItem.SubItems(1) = IIf(IsNull(rsTmp!住院号), "", rsTmp!住院号)
        objItem.SubItems(2) = IIf(IsNull(rsTmp!床号), "", rsTmp!床号)
        objItem.SubItems(3) = IIf(IsNull(rsTmp!住院医师), "", rsTmp!住院医师)
        objItem.SubItems(4) = IIf(IsNull(rsTmp!费别), "", rsTmp!费别)
        objItem.SubItems(5) = IIf(IsNull(rsTmp!护理等级), "", rsTmp!护理等级)
        objItem.SubItems(6) = IIf(IsNull(rsTmp!科室), "", rsTmp!科室)
        objItem.SubItems(7) = Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm")
        objItem.Tag = rsTmp!主页ID
        
        objItem.ForeColor = zlDatabase.GetPatiColor(NVL(rsTmp!病人类型))
        For j = 1 To objItem.ListSubItems.Count
            objItem.ListSubItems(j).ForeColor = zlDatabase.GetPatiColor(NVL(rsTmp!病人类型))
        Next
        
        If rsTmp!病人ID = mlng病人ID Then
            objItem.Checked = True '缺省只选择当前病人
            objItem.EnsureVisible
            objItem.Selected = True
        End If
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitUnits() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSql As String
    
    On Error GoTo errH
    
    '包含门诊观察室
    If InStr(mstrPrivs, ";所有病区;") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " Order by A.编码"
    Else
        '求有权病区：直接所在病区+所在科室所属病区
        strSql = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSql = strSql & " Union " & _
            " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
            " From 病区科室对应 A,部门人员 B,部门表 C" & _
            " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & vbNewLine & _
            " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSql = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSql & ") Group by ID,编码,名称 Order by 编码"
    End If
    
    cboUnit.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng病区ID Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
    End If
    If cboUnit.ListCount > 0 And cboUnit.ListIndex = -1 Then cboUnit.ListIndex = 0
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub chkPatiType_Click(Index As Integer)
    
    If chkPatiType(0).Tag = "1" Then chkPatiType(0).Tag = "": Exit Sub
    If chkPatiType(0).Value = 0 And chkPatiType(1).Value = 0 Then
        chkPatiType(0).Tag = "1"
        chkPatiType(Index).Value = 1
    Else
        Call cboUnit_Click
    End If
End Sub
Private Sub cmdPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141", Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdAllPati_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdNoPati_Click
    End If
End Sub


Private Sub cmdAllPati_Click()
    Call SelectLVW(lvwPati, True)
    lvwPati.SetFocus
End Sub

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
End Sub


Private Sub cmdNoPati_Click()
    Call SelectLVW(lvwPati, False)
    lvwPati.SetFocus
End Sub

Private Sub cmdPreview_Click()
    Call PrintOrPreview(1) '预览
End Sub

Private Sub cmdPrint_Click()
    Call PrintOrPreview(2) '打印
End Sub

Private Sub cmdPreviewAll_Click()
    Call PrintOrPreviewAll(1) '预览
End Sub

Private Sub cmdPrintAll_Click()
    Call PrintOrPreviewAll(2) '打印
End Sub

Private Sub PrintOrPreview(bytMode As Byte)
    Dim blnNOSelect As Boolean, Item As ListItem
    
    For Each Item In lvwPati.ListItems
        If Item.Checked Then
            blnNOSelect = False
            
            Item.Selected = True
            Item.EnsureVisible
            Me.Refresh
            Call PrintContent(bytMode, Split(Item.Key, "_")(1))
        End If
    Next
    If blnNOSelect Then MsgBox "没有选择要打印清单的病人！", vbInformation, gstrSysName
End Sub

Private Sub PrintOrPreviewAll(bytMode As Byte)
    Dim blnNOSelect As Boolean, Item As ListItem
    Dim str病人ID As String
    blnNOSelect = True
    For Each Item In lvwPati.ListItems
        If Item.Checked Then
            blnNOSelect = False
            str病人ID = str病人ID & "," & Split(Item.Key, "_")(1)
        End If
    Next
    If blnNOSelect Then
        MsgBox "没有选择要打印清单的病人！", vbInformation, gstrSysName
        Exit Sub
    End If
    If str病人ID <> "" Then
        str病人ID = Mid(str病人ID, 2)
        ReportOpen gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141_1", Me, "病人ID=" & str病人ID, _
        "开始时间=" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss"), _
        "结束时间=" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss"), _
        "显示退费=" & chkReCharge.Value, _
        "显示零费用=" & chkZeroFee.Value, _
        "病人病区=" & cboUnit.ItemData(cboUnit.ListIndex), _
        "费用时间=" & IIf(opttime(0).Value = True, "登记时间", "发生时间"), bytMode
    End If
End Sub

Private Sub PrintContent(ByVal bytMode As Byte, ByVal str病人ID As String)
    ReportOpen gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141", Me, "病人ID=" & str病人ID, _
        "开始时间=" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss"), _
        "结束时间=" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss"), _
        "显示退费=" & chkReCharge.Value, _
        "显示零费用=" & chkZeroFee.Value, _
        "病人病区=" & cboUnit.ItemData(cboUnit.ListIndex), _
        "主页ID=0", _
        "费用时间=" & IIf(opttime(0).Value = True, "登记时间", "发生时间"), bytMode
End Sub

Private Sub cmdCancel_Click()
    Dim lngTmp As Long
    Dim blnHavePara As Boolean
    
    blnHavePara = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    zlDatabase.SetPara "开始时间", Format(Me.dtpBegin.Value, "hh:mm:ss"), glngSys, mlngModul, blnHavePara
    zlDatabase.SetPara "结束时间", Format(Me.dtpEnd.Value, "hh:mm:ss"), glngSys, mlngModul, blnHavePara

    lngTmp = DateDiff("d", Me.dtpEnd.Value, zlDatabase.Currentdate)
    zlDatabase.SetPara "结束间隔", lngTmp, glngSys, mlngModul, blnHavePara
    lngTmp = DateDiff("d", Me.dtpBegin.Value, Me.dtpEnd.Value)
    zlDatabase.SetPara "开始间隔", lngTmp, glngSys, mlngModul, blnHavePara
            
    
    zlDatabase.SetPara "费用时间", IIf(opttime(1).Value, 1, 0), glngSys, mlngModul, blnHavePara
    If InStr(mstrPrivs, ";参数设置;") > 0 Then
        zlDatabase.SetPara "显示退费", chkReCharge.Value, glngSys, mlngModul, blnHavePara
        zlDatabase.SetPara "显示零费用", chkZeroFee.Value, glngSys, mlngModul, blnHavePara
    End If
    
    zlDatabase.SetPara "非医保病人", chkPatiType(0).Value, glngSys, mlngModul, blnHavePara
    zlDatabase.SetPara "医保病人", chkPatiType(1).Value, glngSys, mlngModul, blnHavePara
        
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long, lngTmp As Long, strStartTime As String, strEndTime As String
    Dim blnParSet As Boolean
    
    blnParSet = InStr(mstrPrivs, ";参数设置;") > 0
    mlngModul = 1141
       
    strEndTime = zlDatabase.GetPara("结束时间", glngSys, mlngModul, "23:59:59", Array(dtpEnd), blnParSet)
    lngTmp = Val(zlDatabase.GetPara("结束间隔", glngSys, mlngModul, 0, Array(dtpEnd), blnParSet))
    If lngTmp > 7 Then lngTmp = 7
    Me.dtpEnd.Value = CDate(Format(zlDatabase.Currentdate() - lngTmp, "yyyy-MM-dd") & " " & strEndTime)
    
    strStartTime = zlDatabase.GetPara("开始时间", glngSys, mlngModul, "00:00:00", Array(dtpBegin), blnParSet)
    lngTmp = Val(zlDatabase.GetPara("开始间隔", glngSys, mlngModul, 0, Array(dtpBegin), blnParSet))
    If lngTmp > 7 Then lngTmp = 7
    Me.dtpBegin.Value = CDate(Format(Me.dtpEnd.Value - lngTmp, "yyyy-MM-dd") & " " & strStartTime)
    
    
    i = IIf(IIf(zlDatabase.GetPara("费用时间", glngSys, mlngModul, "0", Array(opttime(0), opttime(1)), blnParSet) = "1", "发生时间", "登记时间") = "登记时间", 0, 1) '注册表值为1表示按发生时间
    opttime(i).Value = True
    chkReCharge.Value = IIf(zlDatabase.GetPara("显示退费", glngSys, mlngModul, "0", Array(chkReCharge), blnParSet) = "1", 1, 0)
    chkZeroFee.Value = IIf(zlDatabase.GetPara("显示零费用", glngSys, mlngModul, "0", Array(chkZeroFee), blnParSet) = "1", 1, 0)
    
    
    chkPatiType(0).Value = IIf(zlDatabase.GetPara("非医保病人", glngSys, mlngModul, "1", Array(chkPatiType(0)), blnParSet) = "1", 1, 0)
    chkPatiType(1).Value = IIf(zlDatabase.GetPara("医保病人", glngSys, mlngModul, "1", Array(chkPatiType(1)), blnParSet) = "1", 1, 0)

    Call InitUnits '读取病区/病人
    Call zlControl.LvwFlatColumnHeader(lvwPati)
End Sub
