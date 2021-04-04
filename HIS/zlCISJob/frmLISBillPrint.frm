VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLISBillPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "标本条码打印"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboCapture 
      Height          =   300
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1500
      Width           =   2865
   End
   Begin VB.CommandButton cmdInter 
      Caption         =   "中断(F9)"
      Height          =   350
      Left            =   90
      TabIndex        =   11
      Top             =   3105
      Width           =   1100
   End
   Begin VB.CheckBox chkMachine 
      Caption         =   "按仪器分别打印(&S)"
      Height          =   225
      Left            =   405
      TabIndex        =   5
      ToolTipText     =   "选中此选项，在时间段内只要是同一执行科室、同种标本只打印一张条码。否则每一个采集将分别打印。"
      Top             =   2250
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkRetry 
      Caption         =   "已打印的重新打印(&R)"
      Height          =   225
      Left            =   135
      TabIndex        =   6
      Top             =   2610
      Width           =   3315
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3840
      TabIndex        =   8
      Top             =   3105
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   2700
      TabIndex        =   7
      Top             =   3105
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   2955
      Width           =   4965
   End
   Begin VB.CheckBox chkOnly 
      Caption         =   "同一病人的同样标本合并打印(&O)"
      Height          =   225
      Left            =   135
      TabIndex        =   4
      ToolTipText     =   "选中此选项，在时间段内只要是同一执行科室、同种标本只打印一张条码。否则每一个采集将分别打印。"
      Top             =   1920
      Value           =   1  'Checked
      Width           =   3645
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   97386499
      CurrentDate     =   38082
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   97386499
      CurrentDate     =   38082
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "打印条码的采集方式"
      Height          =   180
      Left            =   150
      TabIndex        =   12
      Top             =   1560
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发送时间                      ～"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   2880
   End
   Begin VB.Label lblDesc 
      Appearance      =   0  'Flat
      Caption         =   $"frmLISBillPrint.frx":0000
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   750
      TabIndex        =   9
      Top             =   120
      Width           =   4170
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmLISBillPrint.frx":008C
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frmLISBillPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strRecentDate As String '最近打印的医嘱发送时间
Private strPatiSource As String '病人来源
Private lngDeptID As Long
Private blnCancel As Boolean '是否取消打印作业

Public Sub ShowMe(objParent As Object, ByVal PatiSource As String, DeptID As Long)
    strPatiSource = PatiSource: lngDeptID = DeptID
    blnCancel = False
    
    Me.Show vbModal, objParent
    Unload Me
End Sub

Private Sub chkOnly_Click()
    Me.chkMachine.Enabled = (Me.chkOnly.Value = 1)
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdInter_Click()
    blnCancel = True
End Sub

Private Sub cmdOK_Click()
    '保存打印参数
        
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\条码打印", "最近医嘱时间", Format(dtpEnd.Value, "yyyy-MM-dd HH:mm")
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\条码打印", "采集方式", cboCapture.Text
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\条码打印", "打印方式", IIf(Me.chkMachine, "按仪器", IIf(Me.chkOnly, "按标本", ""))
    If PrintBill Then Me.Hide
End Sub

Private Function PrintBill() As Boolean
    Dim strSQL As String
    Dim strDateFilter As String
    Dim rsTmp As New ADODB.Recordset
    Dim strNO As String, int性质 As Integer
    
    PrintBill = False
    On Error GoTo DataError
    Me.MousePointer = vbHourglass
    
    If Format(dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
        strDateFilter = " And A.发送时间 Between [2] And Sysdate"
    Else
        strDateFilter = " And A.发送时间 Between [2] And [3]"
    End If

    If chkOnly.Value = 1 Then
        If chkMachine.Value = 0 Then
            '同一标本只打一张
            strSQL = "Select 病人ID,标本,执行部门,NO," & _
                " Trim(内容1||' '||内容2||' '||内容3||' '||内容4||' '||内容5) As 项目,编号" & _
                " From" & _
                " (Select B.病人ID,B.标本部位 As 标本,F.名称 As 执行部门,S.编号," & _
                "  Max(Decode(Mod(Rownum,5),0,B.医嘱内容,'')) As 内容1," & _
                "  Max(Decode(Mod(Rownum,5),1,B.医嘱内容,'')) As 内容2," & _
                "  Max(Decode(Mod(Rownum,5),2,B.医嘱内容,'')) As 内容3," & _
                "  Max(Decode(Mod(Rownum,5),3,B.医嘱内容,'')) As 内容4," & _
                "  Max(Decode(Mod(Rownum,5),4,B.医嘱内容,'')) As 内容5," & _
                "  Max(S.NO||','||S.记录性质) As NO" & _
                "  From 病人医嘱记录 B,部门表 F," & _
                "   (Select A.医嘱ID,A.NO,A.记录性质,B.诊疗项目ID," & _
                "    'ZLCISBILL'||trim(to_Char(F.编号, '00000'))||'-1' AS 编号" & _
                "    From 病人医嘱发送 A,病人医嘱记录 B,诊疗项目目录 C, 病历单据应用 E,病历文件列表 F" & _
                "    Where a.医嘱ID = B.ID And B.诊疗项目ID = C.ID AND B.诊疗项目ID=E.诊疗项目ID AND B.病人来源=E.应用场合 AND E.病历文件ID=F.ID" & _
                "     And instr([4],','||B.病人来源||',') > 0  And A.执行部门ID+0= [1] " & _
                "     And C.类别='E' And Nvl(C.操作类型,'0')='6'" & IIf(cboCapture.ItemData(cboCapture.ListIndex) = 0, "", " And C.ID+0=[5]") & _
                strDateFilter & " And Nvl(A.执行状态,0)=0" & IIf(chkRetry, "", " And A.采样人 Is Null") & ") S" & _
                "  Where B.执行科室ID = F.ID And B.相关ID = S.医嘱ID" & _
                "  Group By B.病人ID, B.标本部位,F.名称,S.诊疗项目ID,S.编号)" & _
                " Order By 病人ID"
        Else
            '同一标本再按仪器分别打印
            strSQL = "Select 病人ID,标本,执行部门,NO," & _
                " Trim(内容1||' '||内容2||' '||内容3||' '||内容4||' '||内容5) As 项目,仪器,编号" & _
                " From" & _
                " (Select B.病人ID,B.标本部位 As 标本,F.名称 As 执行部门,S.仪器,S.编号," & _
                "  Max(Decode(Mod(Rownum,5),0,B.医嘱内容,'')) As 内容1," & _
                "  Max(Decode(Mod(Rownum,5),1,B.医嘱内容,'')) As 内容2," & _
                "  Max(Decode(Mod(Rownum,5),2,B.医嘱内容,'')) As 内容3," & _
                "  Max(Decode(Mod(Rownum,5),3,B.医嘱内容,'')) As 内容4," & _
                "  Max(Decode(Mod(Rownum,5),4,B.医嘱内容,'')) As 内容5," & _
                "  Max(S.NO||','||S.记录性质) As NO" & _
                "  From 病人医嘱记录 B,部门表 F," & _
                "   (Select DISTINCT 医嘱ID,NO,记录性质,仪器,诊疗项目ID,编号 FROM " & _
                "    (Select A.医嘱ID,A.NO,A.记录性质,B.诊疗项目ID,I.报告项目ID," & _
                "     'ZLCISBILL'||trim(to_Char(F.编号, '00000'))||'-1' AS 编号,MAX(Decode(M.名称,NULL,'手工',M.名称)) AS 仪器 " & _
                "     From 病人医嘱发送 A,病人医嘱记录 B,病人医嘱记录 D,诊疗项目目录 C,检验报告项目 I,检验仪器项目 J,检验仪器 M, 病历单据应用 E,病历文件列表 F" & _
                "     Where a.医嘱ID = B.ID And B.诊疗项目ID = C.ID" & _
                "      AND D.相关ID = B.ID AND D.诊疗项目ID=I.诊疗项目ID(+) AND I.报告项目ID=J.项目ID(+) AND J.仪器ID=M.ID(+) AND B.诊疗项目ID=E.诊疗项目ID AND B.病人来源=E.应用场合 AND E.病历文件id=F.ID" & _
                "      And instr([4],','||B.病人来源||',') > 0 And A.执行部门ID+0= [1] " & _
                "      And C.类别='E' And Nvl(C.操作类型,'0')='6'" & IIf(cboCapture.ItemData(cboCapture.ListIndex) = 0, "", " And C.ID+0=[5]") & _
                strDateFilter & " And Nvl(A.执行状态,0)=0" & IIf(chkRetry, "", " And A.采样人 Is Null") & _
                "     GROUP BY A.医嘱ID,A.NO,A.记录性质,B.诊疗项目ID,I.报告项目ID,F.编号)" & _
                "   ) S" & _
                "  Where B.执行科室ID = F.ID And B.相关ID = S.医嘱ID" & _
                "  Group By B.病人ID, B.标本部位,F.名称,S.仪器,S.诊疗项目ID,S.编号)" & _
                " Order By 病人ID"
        End If
    Else
        '分别打印
        strSQL = "Select B.病人ID,B.标本部位 as 标本,F.名称 As 执行部门, B.医嘱内容 As 项目,S.NO||','||S.记录性质 As NO,编号" & _
            " From 病人医嘱记录 B,部门表 F," & _
            " (Select A.医嘱ID,A.NO,A.记录性质,'ZLCISBILL'||trim(to_Char(F.编号, '00000'))||'-1' AS 编号 From 病人医嘱发送 A,病人医嘱记录 B,诊疗项目目录 C, 病历单据应用 E,病历文件列表 F" & _
            "  Where a.医嘱ID = B.ID And B.诊疗项目ID = C.ID AND B.诊疗项目ID=E.诊疗项目ID AND B.病人来源=E.应用场合 AND E.病历文件id=F.ID" & _
            "   And instr([4],','||B.病人来源||',') > 0 And A.执行部门ID+0= [1] " & _
            "   And C.类别='E' And Nvl(C.操作类型,'0')='6'" & IIf(cboCapture.ItemData(cboCapture.ListIndex) = 0, "", " And C.ID+0=[5]") & _
            strDateFilter & " And Nvl(A.执行状态,0)=0" & IIf(chkRetry, "", " And A.采样人 Is Null") & ") S" & _
            " Where B.执行科室ID = F.ID And B.相关ID = S.医嘱ID" & _
            " Order By 病人ID"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngDeptID, CDate(Format(dtpBegin.Value, "yyyy-MM-dd hh:mm:ss")), _
                    CDate(Format(dtpEnd.Value, "yyyy-MM-dd hh:mm:ss")), "," & strPatiSource & ",", cboCapture.ItemData(cboCapture.ListIndex))
    
    If rsTmp.EOF Then
        Me.MousePointer = vbDefault
        MsgBox "在该时段内没有需要打印的标本条码。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, Nvl(rsTmp("编号")), Me) Then
        cmdInter.Enabled = True
        Do While Not rsTmp.EOF
            If blnCancel Then PrintBill = True: Exit Function
            strNO = Split(rsTmp("NO"), ",")(0)
            int性质 = Split(rsTmp("NO"), ",")(1)
            DoEvents
            Call ReportOpen(gcnOracle, glngSys, Nvl(rsTmp("编号")), Me, "NO=" & strNO, "性质=" & int性质, "项目=" & Nvl(rsTmp("项目")), 2)
            
            rsTmp.MoveNext
        Loop
        cmdInter.Enabled = False
        '填写采样人、采样时间，表示已经打印
        strSQL = "ZL_病人医嘱执行_批量采样('" & strPatiSource & "'," & lngDeptID & ",'" & _
            Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss") & "','" & _
            IIf(Format(dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm"), "", _
                Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm")) & "','" & UserInfo.姓名 & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Name
    End If
    PrintBill = True
    
    Me.MousePointer = vbDefault
    Exit Function
DataError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
        
    Me.MousePointer = vbDefault
End Function

Private Sub Form_Activate()
    Dim curDate As Date
    
    cmdInter.Enabled = False
    On Error GoTo DataError
    
    curDate = zlDatabase.Currentdate
    dtpEnd.MaxDate = curDate: dtpBegin.MaxDate = curDate
    dtpEnd.Value = Format(curDate, "yyyy-MM-dd HH:mm")
    dtpEnd.Tag = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm")
    dtpBegin.Value = Format(strRecentDate, "yyyy-MM-dd HH:mm")
        
    dtpBegin.SetFocus
    Exit Sub
DataError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call ZLCommFun.PressKey(vbKeyTab)
    End If
    'F9中断
    If KeyCode = 120 And cmdInter.Enabled Then cmdInter_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String
    On Error GoTo DataError
    
    '读取参数
    strRecentDate = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\条码打印", "最近医嘱时间", Format(Date, "yyyy-MM-dd HH:mm"))
    If Not IsDate(strRecentDate) Then strRecentDate = Format(Date, "yyyy-MM-dd HH:mm")
    
    '初始采集方式
    strSQL = "Select Distinct A.ID,A.名称" & _
        " From 诊疗项目目录 A,诊疗执行科室 B " & _
        " Where A.类别='E' AND A.操作类型='6'" & _
        " And (A.撤档时间 IS NULL Or A.撤档时间=To_Date('3000-01-01','yyyy-mm-dd')) " & _
        " And A.ID=B.诊疗项目ID And B.执行科室ID=[1]" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngDeptID)
    With cboCapture
        .AddItem "所有方式"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0
        Do While Not rsTmp.EOF
            .AddItem rsTmp("名称")
            .ItemData(.NewIndex) = rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
        On Error Resume Next
        strTmp = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\条码打印", "采集方式", "所有方式")
        .Text = strTmp
    End With
    strTmp = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\条码打印", "打印方式", "")
    If strTmp = "按仪器" Then
        Me.chkOnly.Value = 1
        Me.chkMachine.Value = 1
    ElseIf strTmp = "按标本" Then
        Me.chkOnly.Value = 1
        Me.chkMachine.Value = 0
    Else
        Me.chkOnly.Value = 0
        Me.chkMachine.Value = 0
    End If
    
    Exit Sub
DataError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
