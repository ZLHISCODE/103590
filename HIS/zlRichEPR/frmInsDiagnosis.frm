VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmInsDiagnosis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "诊断编辑"
   ClientHeight    =   4350
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6180
   Icon            =   "frmInsDiagnosis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vfgSelect 
      Height          =   2175
      Left            =   -4080
      TabIndex        =   19
      Top             =   1395
      Visible         =   0   'False
      Width           =   4680
      _cx             =   8255
      _cy             =   3836
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
   Begin VB.CommandButton cmdRef 
      Caption         =   "查阅诊断参考(&R)…"
      Height          =   350
      Left            =   135
      TabIndex        =   18
      Top             =   3825
      Width           =   1800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4815
      TabIndex        =   11
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3630
      TabIndex        =   10
      Top             =   3825
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -15
      TabIndex        =   17
      Top             =   3630
      Width           =   6345
   End
   Begin VB.Frame fraHint 
      Height          =   1215
      Left            =   1350
      TabIndex        =   12
      Top             =   2235
      Width           =   4575
      Begin VB.OptionButton optHint 
         Caption         =   "按疾病诊断目录检索输入(F4)"
         Height          =   180
         Index           =   2
         Left            =   165
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   840
         Width           =   3360
      End
      Begin VB.OptionButton optHint 
         Caption         =   "按标准疾病编码检索输入(F3)"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   562
         Width           =   3360
      End
      Begin VB.OptionButton optHint 
         Caption         =   "自由输入(F2)"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   285
         Value           =   -1  'True
         Width           =   2190
      End
      Begin VB.Label lblHint 
         AutoSize        =   -1  'True
         Caption         =   "输入方法提示:"
         Height          =   180
         Left            =   135
         TabIndex        =   13
         Top             =   15
         Width           =   1170
      End
   End
   Begin VB.CheckBox chkDoubt 
      Caption         =   "疑诊(&U)"
      Height          =   195
      Left            =   1350
      TabIndex        =   7
      Top             =   1485
      Width           =   945
   End
   Begin VB.TextBox txtSymptom 
      Height          =   300
      Left            =   1350
      TabIndex        =   9
      Top             =   1800
      Width           =   4575
   End
   Begin VB.TextBox txtDisease 
      Height          =   300
      Left            =   1350
      TabIndex        =   6
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -15
      TabIndex        =   4
      Top             =   900
      Width           =   6345
   End
   Begin VB.OptionButton optType 
      Caption         =   "中医诊断(&H)"
      Height          =   180
      Index           =   1
      Left            =   4710
      TabIndex        =   3
      Top             =   510
      Width           =   1335
   End
   Begin VB.OptionButton optType 
      Caption         =   "西医诊断(&W)"
      Height          =   180
      Index           =   0
      Left            =   3345
      TabIndex        =   2
      Top             =   510
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.ComboBox cboKind 
      Height          =   300
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   450
      Width           =   1725
   End
   Begin VB.ComboBox cboIn 
      Height          =   300
      Left            =   3210
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1432
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cboOut 
      Height          =   300
      Left            =   5055
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1432
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblSymptom 
      AutoSize        =   -1  'True
      Caption         =   "证候(&S)"
      Height          =   180
      Left            =   690
      TabIndex        =   8
      Top             =   1860
      Width           =   630
   End
   Begin VB.Label lblDisease 
      AutoSize        =   -1  'True
      Caption         =   "疾病(&D)"
      Height          =   180
      Left            =   690
      TabIndex        =   5
      Top             =   1140
      Width           =   630
   End
   Begin VB.Label lblKind 
      AutoSize        =   -1  'True
      Caption         =   "在住院病历修订过程中，可选择插入的以下类型的诊断："
      Height          =   180
      Left            =   690
      TabIndex        =   0
      Top             =   135
      Width           =   4500
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   135
      Picture         =   "frmInsDiagnosis.frx":038A
      Top             =   195
      Width           =   480
   End
   Begin VB.Label LabOut 
      Caption         =   "出院情况"
      Height          =   210
      Left            =   4260
      TabIndex        =   23
      Top             =   1485
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LabIn 
      Caption         =   "入院病情"
      Height          =   210
      Left            =   2400
      TabIndex        =   22
      Top             =   1500
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "frmInsDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOk As Boolean
Private mobjDoc As cEPRDocument
Private mblnSyncPage As Boolean

Public Function ShowMe(ByRef edtThis As Editor, ByRef frmParent As frmMain) As cEPRDiagnosis
    '功能：显示诊断编辑窗体，并返回编辑的诊断
    '参数： frmParent-父窗体
    
Dim intFileKind As Integer  '病历文件种类
Dim strFileName As String   '病历文件名称
Dim lngFileID As Long       '病历文件定义的Id
Dim lngDeptId As Long       '书写病历的当前科室
Dim blnEmend As Boolean     '是否修订状态
Dim strCurTime As String
Dim rsTemp As New ADODB.Recordset
    
    '------------------------------------
    Set mobjDoc = frmParent.Document
    Select Case mobjDoc.EditType
    Case cprET_病历文件定义
        intFileKind = mobjDoc.EPRFileInfo.种类
        strFileName = mobjDoc.EPRFileInfo.名称
        lngFileID = mobjDoc.EPRFileInfo.ID
        lngDeptId = 0
        blnEmend = False
    Case cprET_全文示范编辑
        intFileKind = 0
        strFileName = mobjDoc.EPRDemoInfo.名称
        lngFileID = mobjDoc.EPRDemoInfo.文件ID
        lngDeptId = 0
        blnEmend = False
    Case cprET_单病历编辑, cprET_单病历审核
        intFileKind = mobjDoc.EPRPatiRecInfo.病历种类
        strFileName = mobjDoc.EPRPatiRecInfo.病历名称
        lngFileID = mobjDoc.EPRPatiRecInfo.文件ID
        lngDeptId = mobjDoc.EPRPatiRecInfo.科室ID
        blnEmend = (mobjDoc.EditType = cprET_单病历审核)
    End Select
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select l.种类, q.事件, q.唯一, q.书写时限, h.中医, Sysdate As 日期 " & _
            " From 病历文件列表 l, 病历时限要求 q," & _
            "      (Select Sign(Nvl(Count(部门ID), 0)) As 中医 From 部门性质说明 Where 部门id = [2] And 工作性质 = '中医科') h" & _
            " Where l.Id = q.文件id(+) And l.Id = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID, lngDeptId)
    If rsTemp.RecordCount <= 0 Then MsgBox "该病历定义设置丢失，不能插入诊断！", vbExclamation, gstrSysName: Exit Function
    intFileKind = rsTemp!种类: strCurTime = Format(rsTemp!日期, "yyyy-mm-dd hh:mm:ss")
    
    Me.cboKind.Clear
    If intFileKind = 1 Then
        Me.lblKind.Caption = "在" & strFileName & IIf(blnEmend, "修订", "编辑") & "过程中，可选择插入的以下类型的诊断："
        Me.cboKind.AddItem "11-门诊诊断"
        Me.cboKind.ListIndex = 0
        Me.cboKind.Enabled = False
        Me.optType(1).Enabled = (rsTemp!中医 = 1)
    ElseIf intFileKind = 2 Then
        If (rsTemp!事件 = "入院" Or rsTemp!事件 = "首次入院" Or rsTemp!事件 = "再次入院") And rsTemp!唯一 = 1 Then
            Me.lblKind.Caption = "在" & strFileName & IIf(blnEmend, "修订", "编辑") & "过程中，可选择插入的以下类型的诊断："
            If blnEmend = False Then
                Me.cboKind.AddItem "21-初步诊断"
                Me.cboKind.ListIndex = 0
                Me.cboKind.Enabled = False
            Else
                Me.cboKind.AddItem "22-确诊诊断"
                Me.cboKind.AddItem "23-修正诊断"
                Me.cboKind.AddItem "24-补充诊断"
                Me.cboKind.ListIndex = 0
            End If
            Me.optType(1).Enabled = (rsTemp!中医 = 1)
        ElseIf rsTemp!事件 = "24小时出院" Or rsTemp!事件 = "24小时死亡" Then
            Me.lblKind.Caption = "在" & strFileName & IIf(blnEmend, "修订", "编辑") & "过程中，可选择插入的以下类型的诊断："
            If blnEmend = False Then
                Me.cboKind.AddItem "21-初步诊断"
            Else
                Me.cboKind.AddItem "22-确诊诊断"
                Me.cboKind.AddItem "23-修正诊断"
                Me.cboKind.AddItem "24-补充诊断"
            End If
            Me.cboKind.AddItem "31-出院诊断"
            Me.cboKind.ListIndex = 0
            Me.optType(1).Enabled = (rsTemp!中医 = 1)
        ElseIf rsTemp!事件 = "出院" Or rsTemp!事件 = "死亡" Then
            Me.lblKind.Caption = "在" & strFileName & IIf(blnEmend, "修订", "编辑") & "过程中，可选择插入的以下类型的诊断："
            Me.cboKind.AddItem "31-出院诊断"
            Me.cboKind.ListIndex = 0
            Me.cboKind.Enabled = False
            Me.optType(1).Enabled = (rsTemp!中医 = 1)
            mblnSyncPage = (zldatabase.GetPara("SyncPage", glngSys, 1070, 0) = 1)
            If mblnSyncPage Then
                Call optType_Click(0)
            End If
        ElseIf rsTemp!事件 = "手术" And rsTemp!唯一 = 1 Then
            Me.lblKind.Caption = "在" & strFileName & IIf(blnEmend, "修订", "编辑") & "过程中，可选择插入的以下类型的诊断："
            Me.cboKind.AddItem "41-术前诊断"
            Me.cboKind.AddItem "42-术后诊断"
            Me.cboKind.ListIndex = 0
            Me.optType(1).Value = False: Me.optType(1).Enabled = False
        Else
            MsgBox "该病历不能插入诊断！", vbExclamation, gstrSysName: Exit Function
        End If
    ElseIf intFileKind = 7 Then     '诊疗报告
        gstrSQL = "Select Nvl(Instr(i.操作类型, '病理'), 0) As 病理" & vbNewLine & _
                "From 病人医嘱记录 l, 诊疗项目目录 i" & vbNewLine & _
                "Where l.诊疗项目id = i.Id And l.诊疗类别 = 'D' And l.Id = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDoc.EPRPatiRecInfo.医嘱id)
        If rsTemp.RecordCount <= 0 Then Exit Function   '非检查报告，不能插入诊断
        Me.lblKind.Caption = "在" & strFileName & IIf(blnEmend, "修订", "编辑") & "过程中，可选择插入的以下类型的诊断："
        If rsTemp.Fields(0).Value > 0 Then
            Me.cboKind.AddItem "51-病理诊断"
        Else
            Me.cboKind.AddItem "52-影像诊断"
        End If
        Me.cboKind.ListIndex = 0
        Me.optType(1).Value = False: Me.optType(1).Enabled = False
        Me.optType(0).Visible = False: Me.optType(1).Visible = False
    Else
        MsgBox "该病历不能插入诊断！", vbExclamation, gstrSysName: Exit Function
    End If
    
    '记录当前的西医中医标志，以便判断改变时清除诊断：
    If Me.optType(0).Value Then
        Me.lblKind.Tag = 0
    Else
        Me.lblKind.Tag = 1
    End If
    
    '------------------------------------
    If Me.optType(0).Value Then
        Me.lblSymptom.Enabled = False: Me.txtSymptom.Enabled = False
    Else
        Me.lblSymptom.Enabled = True: Me.txtSymptom.Enabled = True
    End If
    
    '输入方式通系统参数控制
    '是否允许自由录入
    If Mid(zldatabase.GetPara("诊断输入方式", glngSys, , "11"), IIf(mobjDoc.EPRFileInfo.种类 = cpr门诊病历, 1, 2), 1) = 1 Then
        optHint(0).Enabled = True
    Else
        optHint(0).Value = False
        optHint(0).Enabled = False
    End If
    
    Select Case zldatabase.GetPara("诊断输入来源", glngSys, , "1")
        Case 1 '医生可以选择
            optHint(1).Enabled = True
            optHint(2).Enabled = True
            If optHint(0).Enabled = False Then optHint(1).Value = True
        Case 2 '按诊断
            optHint(1).Enabled = False
            optHint(2).Enabled = True
            If optHint(0).Enabled = False Then optHint(2).Value = True
        Case 3 '按ICD10
            optHint(1).Enabled = True
            optHint(2).Enabled = False
            If optHint(0).Enabled = False Then optHint(1).Value = True
    End Select
    
    
    '显示窗体
    Me.Show vbModal, frmParent
    If mblnOk = False Then Set ShowMe = Nothing: Unload Me: Exit Function
    
    '------------------------------------
    '构造返回对象
    Dim rs As New ADODB.Recordset
    Dim oDiagnosis As cEPRDiagnosis
    Dim strTmp As String
    Dim aryDisease() As String
    
    Set oDiagnosis = New cEPRDiagnosis
    aryDisease = Split(Me.lblDisease.Tag, ",")
    
    
    '检查对应的疾病报告是否书写
    
    If mobjDoc.EPRFileInfo.种类 = cpr门诊病历 Or mobjDoc.EPRFileInfo.种类 = cpr住院病历 Then
        Select Case mobjDoc.EditType
        Case cprET_单病历编辑, cprET_单病历审核
            If UBound(aryDisease) >= 1 And mobjDoc.EPRPatiRecInfo.病人ID > 0 Then
                If Val(aryDisease(1)) > 0 Then
    
                    gstrSQL = "Select Distinct b.名称,c.病人id From 疾病报告前提 a,病历文件列表 b,电子病历记录 c Where a.诊断id=[1] And a.文件id=b.Id And a.文件id=c.文件id(+) And c.病人id(+)=1 And c.主页id(+)=1"
                    Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(aryDisease(1)), mobjDoc.EPRPatiRecInfo.病人ID, mobjDoc.EPRPatiRecInfo.主页ID)
                    If rs.BOF = False Then
                        strTmp = ""
                        Do While Not rs.EOF
                            If zlCommFun.NVL(rs("病人id").Value, 0) = 0 Then
                                strTmp = strTmp & vbCrLf & Space(4) & rs("名称").Value
                            End If
                            rs.MoveNext
                        Loop
                        If strTmp <> "" Then
                            MsgBox "警告：当前病人的如下疾病证明报告还没有书写：" & strTmp, vbInformation, gstrSysName
                        End If
                    End If
    
                ElseIf Val(aryDisease(0)) > 0 Then
                
                    gstrSQL = "Select Distinct b.名称,c.病人id From 疾病报告前提 a,病历文件列表 b,电子病历记录 c Where a.疾病id=[1] And a.文件id=b.Id And a.文件id=c.文件id(+) And c.病人id(+)=1 And c.主页id(+)=1"
                    Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(aryDisease(0)), mobjDoc.EPRPatiRecInfo.病人ID, mobjDoc.EPRPatiRecInfo.主页ID)
                    If rs.BOF = False Then
                        strTmp = ""
                        Do While Not rs.EOF
                            If zlCommFun.NVL(rs("病人id").Value, 0) = 0 Then
                                strTmp = strTmp & vbCrLf & Space(4) & rs("名称").Value
                            End If
                            rs.MoveNext
                        Loop
                        If strTmp <> "" Then
                            MsgBox "警告：当前病人的如下疾病证明报告还没有书写：" & strTmp, vbInformation, gstrSysName
                        End If
                    End If
                End If
            End If
            
        End Select
    End If
    
    Err = 0: On Error GoTo 0
    With oDiagnosis
        .文件ID = lngFileID
        .类型 = Val(Me.cboKind.Text)
        If UBound(aryDisease) < 1 Then
            .疾病id = 0: .诊断id = 0
        Else
            .疾病id = Val(aryDisease(0)): .诊断id = Val(aryDisease(1))
        End If
        .证候id = Val(Me.lblSymptom.Tag)
        If Me.optType(0).Value Then
            .描述 = Trim(Me.txtDisease.Text)
        Else
            .描述 = Trim(Me.txtDisease.Text) & "(" & Trim(Me.txtSymptom.Text) & ")"
        End If
        If Me.chkDoubt.Value = vbChecked Then
            .描述 = .描述 & "(?)"
            .疑诊 = 1
        Else
            .疑诊 = 0
        End If
        If optType(1).Value Then .中医 = 1
        .日期 = strCurTime
        If mblnSyncPage Then
            .入院病情 = cboIn.Text
            .出院情况 = Mid(cboOut.Text, 3)
        End If
    End With
    Set ShowMe = oDiagnosis
    Unload Me: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set ShowMe = Nothing
    Unload Me
End Function

Private Sub cboKind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkDoubt_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chkDoubt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False: Me.Hide: Exit Sub
End Sub

Private Sub cmdOK_Click()
    If (optHint(0).Value = False And txtDisease.Tag = "") Then MsgBox "请输入疾病诊断并回车提取编码！", vbExclamation, gstrSysName: Me.txtDisease.SetFocus: Exit Sub
    If Trim(Me.txtDisease.Text) = "" Then MsgBox "没有输入疾病诊断！", vbExclamation, gstrSysName: Me.txtDisease.SetFocus: Exit Sub
    If Me.optType(1).Value Then
        If (optHint(0).Value = False And txtSymptom.Tag = "") Then MsgBox "请输入证候并回车提取编码！", vbExclamation, gstrSysName: Me.txtDisease.SetFocus: Exit Sub
        If Trim(Me.txtSymptom.Text) = "" Then MsgBox "没有输入证候！", vbExclamation, gstrSysName: Me.txtSymptom.SetFocus:: Exit Sub
    End If
    mblnOk = True: Me.Hide: Exit Sub
End Sub

Private Sub cmdRef_Click()
    Dim aryDisease() As String, lngId As Long
    aryDisease = Split(Me.lblDisease.Tag, ",")
    If UBound(aryDisease) < 1 Then
        lngId = 0
    Else
        lngId = Val(aryDisease(1))
    End If
    Call mobjDoc.Event_ClickDiagRef(lngId, vbModal)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2
        If Me.vfgSelect.Visible = False And optHint(0).Enabled Then Me.optHint(0).Value = True
    Case vbKeyF3
        If Me.vfgSelect.Visible = False And optHint(1).Enabled Then Me.optHint(1).Value = True
    Case vbKeyF4
        If Me.vfgSelect.Visible = False And optHint(2).Enabled Then Me.optHint(2).Value = True
    Case vbKeyEscape
        If Me.vfgSelect.Visible Then
            Me.vfgSelect.Visible = False
        Else
            Call cmdCancel_Click
        End If
    Case Else
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub optHint_Click(Index As Integer)
    If Me.txtDisease.Visible Then Me.txtDisease.SetFocus
End Sub

Private Sub optHint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optType_Click(Index As Integer)
Dim rsTemp As ADODB.Recordset, lCount As Long, i As Integer
    On Error GoTo errHand
    If Me.optType(0).Value Then
        Me.lblSymptom.Enabled = False: Me.txtSymptom.Enabled = False
    Else
        Me.lblSymptom.Enabled = True: Me.txtSymptom.Enabled = True
    End If
    
    If mblnSyncPage And InStr(cboKind.Text, "出院诊断") > 0 Then '病历诊断同步首页诊断，提取首页诊断、入院病情、出院情况
        For i = 1 To mobjDoc.Diagnosises.Count
            If mobjDoc.Diagnosises(i).中医 = Index And mobjDoc.Diagnosises(i).终止版 = 0 Then
                lCount = lCount + 1
            End If
        Next
        
        With cboIn
            .Clear
            .AddItem "有"
            .AddItem "临床未确定"
            .AddItem "情况不明"
            .AddItem "无"
        End With
        gstrSQL = "Select 编码 || '-' || 名称 As 出院情况 From 治疗结果 Order By 编码"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        cboOut.Clear
        Do Until rsTemp.EOF
            cboOut.AddItem rsTemp!出院情况
            rsTemp.MoveNext
        Loop
            
        gstrSQL = "Select 疾病id, 诊断id, 证候id, 诊断描述, 入院病情, 出院情况, 是否未治, 是否疑诊" & vbNewLine & _
                    "From 病人诊断记录" & vbNewLine & _
                    "Where 病人id = [1] And 主页id = [2] And 记录来源=3 And 诊断类型=[3] And 编码序号 = 1 And 诊断次序 = [4]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDoc.EPRPatiRecInfo.病人ID, mobjDoc.EPRPatiRecInfo.主页ID, IIf(Index = 0, 3, 13), lCount + 1)
        If Not rsTemp.EOF Then
            chkDoubt.Value = NVL(rsTemp!是否疑诊, 0)
            Call zlControl.CboSetText(cboIn, NVL(rsTemp!入院病情))
            Call zlControl.CboSetText(cboOut, NVL(rsTemp!出院情况))
            
            Me.lblDisease.Tag = NVL(rsTemp!疾病id, 0) & "," & NVL(rsTemp!诊断id, 0)
            If Index = 0 Then '西医诊断
                Me.lblKind.Tag = 0
                If InStr(NVL(rsTemp!诊断描述), "(") > 0 Then '首页保存的诊断描述是 (编码)名称
                    Me.txtDisease.Tag = Split(Split(NVL(rsTemp!诊断描述), "(")(1), ")")(1): Me.txtDisease.Text = txtDisease.Tag
                Else '病历保存的只有名称
                    Me.txtDisease.Tag = NVL(rsTemp!诊断描述): Me.txtDisease.Text = txtDisease.Tag
                End If
            Else '中医诊断
                Me.lblKind.Tag = 1
                If UBound(Split(NVL(rsTemp!诊断描述), "(")) > 1 Then
                    Me.txtDisease.Tag = Split(Split(NVL(rsTemp!诊断描述), "(")(1), ")")(1)
                    Me.txtDisease.Text = Me.txtDisease.Tag

                    Me.lblSymptom.Tag = "" & NVL(rsTemp!证候id, 0)
                    Me.txtSymptom.Tag = Split(Split(NVL(rsTemp!诊断描述), "(")(2), ")")(0): Me.txtSymptom.Text = Me.txtSymptom.Tag
                Else
                    Me.txtDisease.Tag = Split(NVL(rsTemp!诊断描述), "(")(0)
                    Me.txtDisease.Text = Me.txtDisease.Tag

                    Me.lblSymptom.Tag = "" & NVL(rsTemp!证候id, 0)
                    Me.txtSymptom.Tag = Split(Split(NVL(rsTemp!诊断描述), "(")(1), ")")(0): Me.txtSymptom.Text = Me.txtSymptom.Tag
                End If
            End If
        End If
        
        Me.LabIn.Visible = True: Me.cboIn.Visible = True
        Me.LabOut.Visible = True: Me.cboOut.Visible = True
    End If
    
    If Val(Me.lblKind.Tag) = 0 And Me.optType(0).Value = False Or Val(Me.lblKind.Tag) <> 0 And Me.optType(0).Value Then
        Me.lblDisease.Tag = "": Me.txtDisease.Tag = "": Me.txtDisease.Text = ""
        Me.lblSymptom.Tag = "": Me.txtSymptom.Tag = "": Me.txtSymptom.Text = ""
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub optType_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub optType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtDisease_Change()
    ValidControlText txtDisease
End Sub

Private Sub txtDisease_GotFocus()
    Me.txtDisease.SelStart = 0: Me.txtDisease.SelLength = 4000
    If Me.optHint(0).Value Then
        Call zlCommFun.OpenIme(True)
    Else
        Call zlCommFun.OpenIme(False)
    End If
End Sub

Private Sub txtDisease_KeyPress(KeyAscii As Integer)
Dim rsTemp As New ADODB.Recordset

    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If Me.optHint(0).Value Then
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    ElseIf Me.optHint(1).Value Then
        If Me.txtDisease.Tag = Trim(Me.txtDisease.Text) Or Trim(Me.txtDisease.Text) = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        gstrSQL = "Select Id As 疾病id, r.诊断id, l.编码, l.名称, l.简码" & _
                " From 疾病编码目录 l, (Select 疾病id, Min(诊断id) As 诊断id From 疾病诊断对照 Group By 疾病id) r" & _
                " Where l.类别 = [1] And l.Id = r.疾病id(+) And (l.编码 Like [2] Or l.名称 Like [3] Or l.简码 Like [3])" & _
                " And (l.撤档时间 is Null Or l.撤档时间>=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Order By l.编码"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            IIf(Me.optType(0).Value, "D", "B"), _
            UCase(Trim(Me.txtDisease.Text)) & "%", _
            gstrMatch & UCase(Trim(Me.txtDisease.Text)) & "%")
        If rsTemp.RecordCount <= 0 Then
            MsgBox "未找到要求的标准疾病编码！", vbExclamation, gstrSysName: Exit Sub
        ElseIf rsTemp.RecordCount = 1 Then
            Me.lblDisease.Tag = rsTemp!疾病id & "," & rsTemp!诊断id
            Me.txtDisease.Tag = rsTemp!名称: Me.txtDisease.Text = rsTemp!名称
        Else
            With Me.vfgSelect
                .Tag = "D"
                .Clear
                Set .DataSource = rsTemp
                .ColHidden(0) = True: .ColHidden(1) = True
                .Row = .FixedRows
                .Move Me.txtDisease.Left, Me.txtDisease.Top + Me.txtDisease.Height, Me.txtDisease.Width
                .Visible = True
                .SetFocus
            End With
        End If
    
    ElseIf Me.optHint(2).Value Then
        If Me.txtDisease.Tag = Trim(Me.txtDisease.Text) Or Trim(Me.txtDisease.Text) = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        gstrSQL = "Select r.疾病id, n.诊断id, l.编码, n.名称, n.简码" & _
                " From 疾病诊断目录 l, 疾病诊断别名 n, (Select 诊断id, Min(疾病id) As 疾病id From 疾病诊断对照 Group By 诊断id) r" & _
                " Where l.Id = n.诊断id And l.类别 = [1] And l.Id = r.诊断id(+) And" & _
                "       (l.编码 Like [2] Or n.名称 Like [3] Or n.简码 Like [3])" & _
                " And (l.撤档时间 is Null Or l.撤档时间>=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Order By l.编码"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            IIf(Me.optType(0).Value, 1, 2), _
            UCase(Trim(Me.txtDisease.Text)) & "%", _
            gstrMatch & UCase(Trim(Me.txtDisease.Text)) & "%")
        If rsTemp.RecordCount <= 0 Then
            MsgBox "未找到要求的疾病诊断条目！", vbExclamation, gstrSysName: Exit Sub
        ElseIf rsTemp.RecordCount = 1 Then
            Me.lblDisease.Tag = rsTemp!疾病id & "," & rsTemp!诊断id
            Me.txtDisease.Tag = rsTemp!名称: Me.txtDisease.Text = rsTemp!名称
        Else
            With Me.vfgSelect
                .Tag = "D"
                .Clear
                Set .DataSource = rsTemp
                .ColHidden(0) = True: .ColHidden(1) = True
                .Row = .FixedRows
                .Move Me.txtDisease.Left, Me.txtDisease.Top + Me.txtDisease.Height, Me.txtDisease.Width
                .Visible = True
                .SetFocus
            End With
        End If
    End If
End Sub

Private Sub txtSymptom_Change()
    ValidControlText txtSymptom
End Sub

Private Sub txtSymptom_GotFocus()
    Me.txtSymptom.SelStart = 0: Me.txtSymptom.SelLength = 4000
    If Me.optHint(0).Value Then
        Call zlCommFun.OpenIme(True)
    Else
        Call zlCommFun.OpenIme(False)
    End If
End Sub

Private Sub txtSymptom_KeyPress(KeyAscii As Integer)
Dim rsTemp As New ADODB.Recordset
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If Me.optHint(0).Value Then
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    ElseIf Me.optHint(1).Value Then
        If Me.txtSymptom.Tag = Trim(Me.txtSymptom.Text) Or Trim(Me.txtSymptom.Text) = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        gstrSQL = "Select Id As 证候id, 编码, 名称, 简码" & _
                " From 疾病编码目录" & _
                " Where 类别 = 'Z' And (编码 Like [1] Or 名称 Like [2] Or 简码 Like [2])" & _
                " And (撤档时间 is Null Or 撤档时间>=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Order By 编码"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(Trim(Me.txtSymptom.Text)) & "%", gstrMatch & UCase(Trim(Me.txtSymptom.Text)) & "%")
        If rsTemp.RecordCount <= 0 Then
            MsgBox "未找到要求的标准中医证候！", vbExclamation, gstrSysName: Exit Sub
        ElseIf rsTemp.RecordCount = 1 Then
            Me.lblSymptom.Tag = "" & rsTemp!证候id
            Me.txtSymptom.Tag = rsTemp!名称: Me.txtSymptom.Text = rsTemp!名称
        Else
            With Me.vfgSelect
                .Tag = "S"
                .Clear
                Set .DataSource = rsTemp
                .ColHidden(0) = True
                .Row = .FixedRows
                .Move Me.txtSymptom.Left, Me.txtSymptom.Top + Me.txtSymptom.Height, Me.txtSymptom.Width
                .Visible = True
                .SetFocus
            End With
        End If
    Else
        If Me.txtSymptom.Tag = Trim(Me.txtSymptom.Text) And Trim(Me.txtSymptom.Text) <> "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        
        Dim aryDisease() As String, lngDisease As Long
        aryDisease = Split(Me.lblDisease.Tag, ",")
        If UBound(aryDisease) < 1 Then
            lngDisease = 0
        Else
            lngDisease = Val(aryDisease(1))
        End If
        gstrSQL = "Select Distinct 证候id, 证候序号 As 序号, 证候名称 As 名称, Zlspellcode(证候名称) As 简码" & _
                " From 疾病诊断参考" & _
                " Where 诊断id = [1] And 证候序号 Is Not Null And (证候名称 Like [2] Or Zlspellcode(证候名称) Like [2])" & _
                " Order By 证候序号"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngDisease, gstrMatch & UCase(Trim(Me.txtSymptom.Text)) & "%")
        If rsTemp.RecordCount <= 0 Then
            MsgBox "未找到要求的当前中医证候！", vbExclamation, gstrSysName: Exit Sub
        ElseIf rsTemp.RecordCount = 1 Then
            Me.lblSymptom.Tag = "" & rsTemp!证候id
            Me.txtSymptom.Tag = rsTemp!名称: Me.txtSymptom.Text = rsTemp!名称
        Else
            With Me.vfgSelect
                .Tag = "S"
                .Clear
                Set .DataSource = rsTemp
                .ColHidden(0) = True
                .Row = .FixedRows
                .Move Me.txtSymptom.Left, Me.txtSymptom.Top + Me.txtSymptom.Height, Me.txtSymptom.Width
                .Visible = True
                .SetFocus
            End With
        End If
    End If
End Sub

Private Sub vfgSelect_DblClick()
    With Me.vfgSelect
        Select Case .Tag
        Case "D"
            Me.lblDisease.Tag = Val(.TextMatrix(.Row, 0)) & "," & Val(.TextMatrix(.Row, 1))
            Me.txtDisease.Tag = .TextMatrix(.Row, 3): Me.txtDisease.Text = .TextMatrix(.Row, 3)
            '为保证中医病证吻合，清除证候等待重新输入
            Me.lblSymptom.Tag = "": Me.txtSymptom.Tag = "": Me.txtSymptom.Text = ""
        Case "S"
            Me.lblSymptom.Tag = Val(.TextMatrix(.Row, 0))
            Me.txtSymptom.Tag = .TextMatrix(.Row, 2): Me.txtSymptom.Text = .TextMatrix(.Row, 2)
        End Select
        .Visible = False
    End With
End Sub

Private Sub vfgSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vfgSelect_DblClick
    End If
End Sub

Private Sub vfgSelect_LostFocus()
    With Me.vfgSelect
        .Visible = False
        Select Case .Tag
        Case "D": Me.txtDisease.SetFocus
        Case "S": Me.txtSymptom.SetFocus
        End Select
    End With
End Sub
