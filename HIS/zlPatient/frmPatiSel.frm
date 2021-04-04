VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiSel 
   Caption         =   "病人选择"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   Icon            =   "frmPatiSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6945
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   4350
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   529
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1875
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSel.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3540
      Left            =   2100
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3540
      ScaleWidth      =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   90
      Width           =   45
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   3600
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   6350
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   476
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   6945
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3750
      Width           =   6945
      Begin VB.CommandButton cmdFilter 
         Caption         =   "过滤(&F)"
         Height          =   350
         Left            =   195
         TabIndex        =   4
         ToolTipText     =   "筛选满足条件的病人(Ctrl+F)"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "定位(&G)"
         Height          =   350
         Left            =   1410
         TabIndex        =   5
         ToolTipText     =   "定位到满足条件的病人上(Ctrl+G)"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5445
         TabIndex        =   3
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4215
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPati 
      Height          =   3585
      Left            =   2265
      TabIndex        =   1
      Top             =   75
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   6324
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmPatiSel.frx":06E4
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmPatiSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mlng病人ID As Long
Public mstrPrivs As String

Private mrsPati As ADODB.Recordset
Private mlngCurRow As Long, mlngTopRow As Long
Private mlngGo As Long, mblnDown As Boolean, mblnGo As Boolean
Private mstrFilter As String
Private mfrmFilter As frmPatiFilter
Private mfrmFind As frmPatiFind
Private mstrUnitIDs As String '操作员所在病区或科室所属病区

Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    登记时间B As Date
    登记时间E As Date
    出生时间B As Date
    出生时间E As Date
    入院时间B As Date
    入院时间E As Date
    出院时间B As Date
    出院时间E As Date
    住院号 As String
    性别 As String
    费别 As String
    区域 As String
    Patient As String
End Type
Private SQLCondition As Type_SQLCondition

Private Sub cmdCancel_Click()
    mlng病人ID = 0
    
    SaveWinState Me, App.ProductName
    
    Hide
End Sub

Private Sub cmdFilter_Click()
    Dim blnOK As Boolean
    
    blnOK = gblnOK
    mfrmFilter.mbytType = Val(mshPati.Tag)
    mfrmFilter.Show 1, Me
    If gblnOK Then
        With mfrmFilter
            mstrFilter = .mstrFilter
            SQLCondition.登记时间B = .dtp登记B
            SQLCondition.登记时间E = .dtp登记E
            SQLCondition.出生时间B = .dtp出生B
            SQLCondition.出生时间E = .dtp出生E
            
            SQLCondition.入院时间B = .dtp入院B
            SQLCondition.入院时间E = .dtp入院E
            SQLCondition.出院时间B = .dtp出院B
            SQLCondition.出院时间E = .dtp出院E
            
            SQLCondition.住院号 = Trim(.txt住院号.Text)
            SQLCondition.性别 = zlCommFun.GetNeedName(.cbo性别.Text)
            SQLCondition.费别 = zlCommFun.GetNeedName(.cbo费别.Text)
            SQLCondition.区域 = zlCommFun.GetNeedName(.txt区域.Text)
            
            If .PatiIdentify.GetCurCard.名称 = "姓名" And .mlngPatiId = 0 And (.chk登记.Value = 1 Or .chk入院.Value = 1 Or .chk出院.Value = 1) Then       '姓名
                SQLCondition.Patient = Trim(.PatiIdentify.Text) & "%"
            Else
                SQLCondition.Patient = IIf(.mlngPatiId <> 0, .mlngPatiId, .PatiIdentify.Text)
            End If
        End With
        
        Call ShowPatis(mstrFilter)
    End If
    gblnOK = blnOK
End Sub

Private Sub cmdFind_Click()
    Dim blnOK As Boolean
    blnOK = gblnOK
    mfrmFind.mbytType = Val(mshPati.Tag)
    mfrmFind.Show 1, Me
    If gblnOK Then Call SeekPati(mfrmFind.optHead)
    gblnOK = blnOK
End Sub

Private Sub cmdOK_Click()
    If Val(mshPati.TextMatrix(mshPati.Row, 0)) = 0 Then
        If glngSys Like "8??" Then
            MsgBox "没有客户可以选择！", vbInformation, gstrSysName
        Else
            MsgBox "没有病人可以选择！", vbInformation, gstrSysName
        End If
        Exit Sub
    End If
    mlng病人ID = Val(mshPati.TextMatrix(mshPati.Row, 0))
    
    SaveWinState Me, App.ProductName
    
    Hide
End Sub

Private Sub Form_Activate()
    mshPati.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Call SeekPati(False)
        Case vbKeyReturn
            cmdOK_Click
        Case vbKeyEscape
            mblnGo = False
        Case vbKeyF
            If Shift = 2 Then cmdFilter_Click
        Case vbKeyG
            If Shift = 2 Then cmdFind_Click
    End Select
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    mlng病人ID = 0
    If glngSys Like "8??" Then
        Caption = "客户选择"
        tvw_s.Visible = False
        pic.Visible = False
    End If
    
    Set mfrmFilter = New frmPatiFilter
    Set mfrmFind = New frmPatiFind
        
    mstrUnitIDs = GetUserUnits
    Call InitUnits
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tvw_s.Left = 0
    tvw_s.Top = 0
    tvw_s.Height = Me.ScaleHeight - picCmd.Height - sta.Height
    
    pic.Top = 0
    pic.Left = tvw_s.Width
    pic.Height = tvw_s.Height
    
    mshPati.Top = 0
    mshPati.Left = IIf(pic.Visible, pic.Width, 0) + IIf(tvw_s.Visible, tvw_s.Width, 0)
    mshPati.Width = Me.ScaleWidth - IIf(pic.Visible, pic.Width, 0) - IIf(tvw_s.Visible, tvw_s.Width, 0)
    mshPati.Height = tvw_s.Height
    
    If ScaleWidth - cmdCancel.Width - 300 > 4000 Then
        cmdCancel.Left = ScaleWidth - cmdCancel.Width - 300
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmFind
    Unload mfrmFilter
    mstrFilter = ""
    mlng病人ID = 0
End Sub

Private Sub mshPati_DblClick()
    cmdOK_Click
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or mshPati.Width - X < 1000 Then Exit Sub
        pic.Left = pic.Left + X
        tvw_s.Width = tvw_s.Width + X
        mshPati.Left = mshPati.Left + X
        mshPati.Width = mshPati.Width - X
    End If
End Sub

Private Function InitUnits() As Boolean
'功能：初始化病人病区分布列表
'说明：以病区分层
    Dim rsTmp As New ADODB.Recordset
    Dim objNode As Node, i As Integer
    Dim strPreKey  As String
    
    On Error GoTo errH
    
    strPreKey = ""
    If Not tvw_s.SelectedItem Is Nothing Then strPreKey = tvw_s.SelectedItem.Key
    
    If glngSys Like "8??" Then
        tvw_s.Nodes.Clear
        Set objNode = tvw_s.Nodes.Add(, , "Clinic", "所有客户", 1)
        objNode.Expanded = True
        objNode.Selected = True
    Else
        tvw_s.Nodes.Clear
        Set objNode = tvw_s.Nodes.Add(, , "All", "所有病人", 1)
        objNode.Expanded = True
        
        Set objNode = tvw_s.Nodes.Add("All", 4, "In", "在院病人", 1)
        Set objNode = tvw_s.Nodes.Add("All", 4, "Out", "出院病人", 1)
        Set objNode = tvw_s.Nodes.Add("All", 4, "Clinic", "门诊病人", 1)
        Set objNode = tvw_s.Nodes.Add("All", 4, "Temp", "留观病人", 1)
        objNode.Expanded = True
        If objNode.Key = strPreKey Then objNode.Selected = True
                
        Set rsTmp = GetUnit(InStr(mstrPrivs, "所有病区") = 0, "1,2,3", "护理")
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                Set objNode = tvw_s.Nodes.Add("In", 4, "D" & rsTmp!ID, "[" & rsTmp!编码 & "]" & rsTmp!名称, 1)
                
                If rsTmp!ID = UserInfo.部门ID Then objNode.Selected = True
                If objNode.Key = strPreKey Then objNode.Selected = True
                objNode.Expanded = True
                
                rsTmp.MoveNext
            Next
        End If
        If tvw_s.SelectedItem Is Nothing Then tvw_s.Nodes("In").Selected = True
    End If
    
    InitUnits = True
    
    Call tvw_s_NodeClick(tvw_s.SelectedItem)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If tvw_s.Tag = Node.Key Then Exit Sub
    tvw_s.Tag = Node.Key
    
    SaveFlexState mshPati, App.ProductName & "\" & Me.Name
    Call ShowPatis("", , True)  '切换病人类型时,条件清空,使用缺省条件
End Sub

Private Sub ShowPatis(Optional ByVal strIF As String, Optional blnSort As Boolean, Optional blnSet As Boolean)
'功能：根据当前菜单浏览要求(自动生成条件),读取病人信息
'参数：strIF=" And ...."形式的过滤条件
    Dim i As Integer, strSQL As String
    Dim strInfo As String, Curdate As Date
    Dim strCard As String, lngUnitID As String
    Dim blnLimitUnit As Boolean, blnFirst As Boolean
    
    On Error GoTo errH
    
    If Not blnSort Then
        blnLimitUnit = InStr(mstrPrivs, "所有病区") = 0
        
        If strIF = "" Then
            blnFirst = True
            If InStr(1, ",All,Clinic,Temp,", "," & tvw_s.SelectedItem.Key & ",") > 0 Then
                strIF = " And A.登记时间 Between trunc(Sysdate) And Sysdate"
            ElseIf tvw_s.SelectedItem.Key = "Out" Then
                strIF = " And P.出院日期 Between trunc(Sysdate) And Sysdate"
            ElseIf tvw_s.SelectedItem.Key = "In" Then
                strIF = " And P.入院日期 Between trunc(Sysdate) And Sysdate"
            End If
        End If
        strIF = strIF & " And A.停用时间 is NULL"
        '就诊卡号显示
        strCard = "Decode(" & IIf(gblnShowCard, 1, 0) & ",1,A.就诊卡号,LPAD('*',Length(A.就诊卡号),'*')) as 就诊卡,"
        
        
        If tvw_s.SelectedItem.Key = "All" Then '所有病人
            strIF = strIF & IIf(blnLimitUnit, " And (A.当前病区ID Is NULL Or Instr(','||[2]||',',','||A.当前病区ID||',')>0)", "")
            
            '问题25886 by lesfeng 2009-10-28 处理表列头含病区，而SQL不含病区，以致科室之后移位 b
'            strSQL = "Select A.病人ID,A.门诊号,A.住院号," & strCard & "A.姓名,A.性别,A.年龄,A.费别 as 门诊费别," & _
'            " C.名称 as 科室,A.当前床号 as 床号,To_Char(A.入院时间,'YYYY-MM-DD') as 入院时间," & _
'            " To_Char(A.出院时间,'YYYY-MM-DD') as 出院时间,A.住院次数,To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期," & _
'            " A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份,A.身份证号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间," & _
'            " Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型" & _
'            " From 病案主页 P,病人信息 A,部门表 C" & _
'            " Where A.当前科室ID=C.ID(+) And A.病人ID=P.病人ID(+) And A.主页ID=P.主页ID(+) " & strIF & _
'            " Order by A.登记时间 Desc"
            strSQL = "Select A.病人ID,A.门诊号,A.住院号," & strCard & "A.姓名,A.性别,A.年龄,A.费别 as 门诊费别," & _
            " B.名称 as 病区,C.名称 as 科室,A.当前床号 as 床号,To_Char(A.入院时间,'YYYY-MM-DD') as 入院时间," & _
            " To_Char(A.出院时间,'YYYY-MM-DD') as 出院时间,A.住院次数,To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期," & _
            " A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份,A.身份证号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间," & _
            " Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型" & _
            " From 病案主页 P,病人信息 A,部门表 B,部门表 C" & _
            " Where A.当前病区ID=B.ID(+) And A.当前科室ID=C.ID(+) And A.病人ID=P.病人ID(+) And A.主页ID=P.主页ID(+) " & strIF & _
            " Order by A.登记时间 Desc"
            '问题25886 by lesfeng 2009-10-28 处理表列头含病区，而SQL不含病区，以致科室之后移位 b
            strInfo = "正在读取所有病人清单,请稍候 ..."
            If Val(mshPati.Tag) <> 0 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 0
        ElseIf tvw_s.SelectedItem.Key = "In" Or Mid(tvw_s.SelectedItem.Key, 1, 1) = "D" Then  '在院病人
            '58842,刘鹏飞,2013-02-25,在院病人读取(从在院病人中读取)
            If Mid(tvw_s.SelectedItem.Key, 1, 1) = "D" Then
                lngUnitID = Mid(tvw_s.SelectedItem.Key, 2)
                strIF = strIF & " And E.病区ID= [1] "
            Else
                If blnLimitUnit Then
                    strIF = strIF & " And Instr(','||[2]||',',','||E.病区ID||',')>0"
                End If
            End If
            
            strSQL = "Select A.病人ID,A.住院号," & strCard & "NVL(P.姓名,A.姓名) 姓名,NVL(P.性别,A.性别) 性别,NVL(P.年龄,A.年龄) 年龄,P.费别 as 住院费别," & _
                " B.名称 as 病区,C.名称 as 科室,A.当前床号 as 床号,To_Char(A.入院时间,'YYYY-MM-DD') as 入院时间," & _
                " A.住院次数,To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
                " A.身份证号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间,Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型" & _
                " From 病案主页 P,病人信息 A,部门表 B,部门表 C,在院病人 E" & _
                " Where A.当前病区ID=B.ID And A.当前科室ID=C.ID" & strIF & _
                " And A.病人ID=P.病人ID And A.主页ID=P.主页ID And A.病人ID=E.病人ID And Nvl(P.主页ID,0)<>0 " & _
                " Order by A.入院时间 Desc,A.住院号 Desc"
            
            strInfo = "正在读取在院病人清单,请稍候 ..."
            If Val(mshPati.Tag) <> 1 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 1
        ElseIf tvw_s.SelectedItem.Key = "Out" Then '出院病人
            strIF = strIF & IIf(blnLimitUnit, " And Instr(','||[2]||',',','||P.当前病区ID||',')>0", "")
                    
            strSQL = "Select A.病人ID,A.住院号," & strCard & "NVL(P.姓名,A.姓名) 姓名,NVL(P.性别,A.性别) 性别,NVL(P.年龄,A.年龄) 年龄,P.费别 as 住院费别," & _
                " To_Char(A.入院时间,'YYYY-MM-DD') as 入院时间,To_Char(A.出院时间,'YYYY-MM-DD') as 出院时间," & _
                " A.住院次数,To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
                " A.身份证号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间,Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型" & _
                " From 病案主页 P,病人信息 A" & _
                " Where A.病人ID=P.病人ID And A.主页ID=P.主页ID" & _
                " And Nvl(P.主页ID,0)<>0 And P.出院日期 Is Not NULL " & strIF & _
                " Order by A.出院时间 Desc,A.住院号"
            
            strInfo = "正在读取出院病人清单,请稍候 ..."
            If Val(mshPati.Tag) <> 2 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 2
        ElseIf tvw_s.SelectedItem.Key = "Clinic" Then '门诊病人
            strSQL = "Select A.病人ID,A.门诊号," & strCard & "A.姓名,A.性别,A.年龄," & _
                " A.费别 as " & IIf(glngSys Like "8??", "会员等级", "门诊费别") & "," & _
                " To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
                " A.身份证号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间,Decode(A.险类,Null,'普通病人','医保病人') 病人类型" & _
                " From 病人信息 A " & _
                " Where A.当前病区ID is NULL And A.当前科室ID is NULL And A.主页ID is NULL" & strIF & _
                " Order by A.登记时间 Desc,A.门诊号 Desc"
            
            If glngSys Like "8??" Then
                strInfo = "正在读取客户清单,请稍候 ..."
            Else
                strInfo = "正在读取门诊病人清单,请稍候 ..."
            End If
            
            If Val(mshPati.Tag) <> 3 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 3
        ElseIf tvw_s.SelectedItem.Key = "Temp" Then
            '门诊留观和住院留观病人
            strSQL = "Select Distinct A.病人ID,Decode(P.病人性质,1,'门诊留观','住院留观') as 性质, A.门诊号," & strCard & "NVL(P.姓名,A.姓名) 姓名,NVL(P.性别,A.性别) 性别,NVL(P.年龄,A.年龄) 年龄," & _
                " A.费别 as " & IIf(glngSys Like "8??", "会员等级", "门诊费别") & "," & _
                " To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
                " A.身份证号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间,Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型" & _
                " From 病案主页 P,病人信息 A " & _
                " Where A.病人ID=P.病人ID And P.病人性质<>0 And A.住院号 is Null " & strIF & _
                " Order by 性质,登记时间 Desc"
            
            strInfo = "正在读取留观病人清单,请稍候 ..."
            
            If Val(mshPati.Tag) <> 4 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 4
        End If
        
        tvw_s.Tag = tvw_s.SelectedItem.Key
        sta.SimpleText = strInfo
        Screen.MousePointer = 11
        DoEvents
        Me.Refresh
        
        With SQLCondition
            Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUnitID, mstrUnitIDs, .登记时间B, .登记时间E, .出生时间B, .出生时间E, _
            .入院时间B, .入院时间E, .出院时间B, .出院时间E, .住院号, .性别, .区域, .费别, .Patient)
        End With
    End If
    
    mshPati.Clear
    mshPati.Rows = 2
    
    If mrsPati.EOF Then
        Call SetHeader(blnSet)
        If glngSys Like "8??" Then
            sta.SimpleText = IIf(blnFirst, "当天", "") & "没有找到符合条件的客户,请点击[筛选],选择查询条件."
        Else
            sta.SimpleText = IIf(blnFirst, "当天", "") & "没有找到符合条件的病人,请点击[筛选],选择查询条件."
        End If
    Else
        Set mshPati.DataSource = mrsPati
        Call SetHeader(blnSet)
        If glngSys Like "8??" Then
            sta.SimpleText = IIf(blnFirst, "当天", "") & "共找到 " & mrsPati.RecordCount & " 位符合条件的客户"
        Else
            sta.SimpleText = IIf(blnFirst, "当天", "") & "共找到 " & mrsPati.RecordCount & " 位符合条件的病人."
        End If
    End If
    
    Screen.MousePointer = 0
    
    Me.Refresh
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetHeader(Optional blnSet As Boolean)
    Dim strHead As String
    Dim i As Integer
    
    If tvw_s.SelectedItem.Key = "All" Then '所有病人
        strHead = "病人ID,1,750|门诊号,1,750|住院号,1,750|就诊卡,4,850|姓名,1,800|性别,4,500|年龄,4,800|门诊费别,4,850|" & _
            "病区,1,850|科室,1,850|床号,4,500|入院时间,4,1000|出院时间,4,1000|住院次数,4,850|出生日期,4,1000|" & _
            "国籍,4,500|民族,4,800|区域,1,600|学历,4,500|职业,1,1000|身份,1,750|身份证号,1,2000|家庭地址,1,2000|工作单位,1,2000|登记时间,4,1000|病人类型,1,800"
    ElseIf tvw_s.SelectedItem.Key = "Clinic" Then '门诊病人
        If glngSys Like "8??" Then
            strHead = "客户ID,1,750|客户号,1,750|会员卡,4,850|姓名,1,800|性别,4,500|年龄,4,800|会员等级,4,850|" & _
                "出生日期,4,1000|国籍,4,500|民族,4,800|区域,1,600|学历,4,500|职业,1,1000|身份,1,750|身份证号,1,2000|" & _
                "家庭地址,1,2000|工作单位,1,2000|登记时间,4,1000|病人类型,1,800"
        Else
            strHead = "病人ID,1,750|门诊号,1,750|就诊卡,4,850|姓名,1,800|性别,4,500|年龄,4,800|门诊费别,4,850|" & _
                "出生日期,4,1000|国籍,4,500|民族,4,800|区域,1,600|学历,4,500|职业,1,1000|身份,1,750|身份证号,1,2000|" & _
                "家庭地址,1,2000|工作单位,1,2000|登记时间,4,1000|病人类型,1,800"
        End If
    ElseIf tvw_s.SelectedItem.Key = "Temp" Then  '留观病人
         strHead = "病人ID,1,750|性质,1,1000|门诊号,1,750|就诊卡,4,850|姓名,1,800|性别,4,500|年龄,4,800|门诊费别,4,850|" & _
                "出生日期,4,1000|国籍,4,500|民族,4,800|区域,1,600|学历,4,500|职业,1,1000|身份,1,750|身份证号,1,2000|" & _
                "家庭地址,1,2000|工作单位,1,2000|登记时间,4,1000|病人类型,1,800"
    ElseIf tvw_s.SelectedItem.Key = "Out" Then '出院病人
        strHead = "病人ID,1,750|住院号,1,750|就诊卡,4,850|姓名,1,800|性别,4,500|年龄,4,800|住院费别,4,850|" & _
            "入院时间,4,1000|出院时间,4,1000|住院次数,4,850|出生日期,4,1000|国籍,4,500|民族,4,800|区域,1,600|" & _
            "学历,4,500|职业,1,1000|身份,1,750|身份证号,1,2000|家庭地址,1,2000|工作单位,1,2000|登记时间,4,1000|病人类型,1,800"
    ElseIf tvw_s.SelectedItem.Key = "In" Or InStr("D", Left(tvw_s.SelectedItem.Key, 1)) > 0 Then '在院病人
        strHead = "病人ID,1,750|住院号,1,750|就诊卡,4,850|姓名,1,800|性别,4,500|年龄,4,800|住院费别,4,850|" & _
            "病区,1,850|科室,1,850|床号,4,500|入院时间,4,1000|住院次数,4,850|出生日期,4,1000|" & _
            "国籍,4,500|民族,4,800|区域,1,600|学历,4,500|职业,1,1000|身份,1,750|身份证号,1,2000|家庭地址,1,2000|工作单位,1,2000|登记时间,4,1000|病人类型,1,800"
    End If
    
    With mshPati
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Or blnSet Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Or blnSet Then Call RestoreFlexState(mshPati, App.ProductName & "\" & Me.Name)
        
        If glngSys Like "8??" Then .ColWidth(1) = 0
        .RowHeight(0) = 320
        
        '恢复上次行
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .Cols - 1
        Call mshPati_EnterCell

        .Redraw = True
    End With
End Sub

Private Sub mshPati_EnterCell()
    If glngSys Like "8??" Then
        If mshPati.Row = 0 Or mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID")) = "" Then Exit Sub
    Else
        If mshPati.Row = 0 Or mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")) = "" Then Exit Sub
    End If
    mlngGo = mshPati.Row
    mlngCurRow = mshPati.Row: mlngTopRow = mshPati.TopRow
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To mshPati.Cols - 1
        If mshPati.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub mshPati_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshPati.MouseRow = 0 Then
        mshPati.MousePointer = 99
    Else
        mshPati.MousePointer = Screen.MousePointer
    End If
End Sub

Private Sub mshPati_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshPati.MouseCol
    
    If Button = 1 And mshPati.MousePointer = 99 And mblnDown Then '双击最大化时会执行
        mblnDown = False
        
        If mshPati.TextMatrix(0, lngCol) = "" Then Exit Sub
        If glngSys Like "8??" Then
            If mshPati.TextMatrix(1, GetColNum("客户ID")) = "" Then Exit Sub
        Else
            If mshPati.TextMatrix(1, GetColNum("病人ID")) = "" Then Exit Sub
        End If
        
        Set mshPati.DataSource = Nothing
        
        Select Case mshPati.TextMatrix(0, lngCol)
            Case "客户ID"
                mrsPati.Sort = "病人ID" & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
            Case "会员卡"
                mrsPati.Sort = "就诊卡" & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
            Case Else
                mrsPati.Sort = mshPati.TextMatrix(0, lngCol) & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
        End Select
        
        mshPati.ColData(lngCol) = (mshPati.ColData(lngCol) + 1) Mod 2
        
        Call ShowPatis(, True)
    End If
End Sub

Private Sub mshPati_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mblnDown = True
End Sub

Private Sub SeekPati(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    If glngSys Like "8??" Then
        sta.SimpleText = "正在定位满足条件的客户,按ESC终止 ..."
    Else
        sta.SimpleText = "正在定位满足条件的病人,按ESC终止 ..."
    End If
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshPati.Rows - 1
        DoEvents
        
        '比较条件
        blnFill = True
        With mfrmFind
            If .txt病人ID.Text <> "" Then
                If glngSys Like "8??" Then
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("客户ID")) = .txt病人ID.Text
                Else
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("病人ID")) = .txt病人ID.Text
                End If
            End If
            If .txt就诊卡.Text <> "" Then
                If glngSys Like "8??" Then
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("会员卡")) = .txt就诊卡.Text
                Else
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("就诊卡")) = .txt就诊卡.Text
                End If
            End If
            If .txt门诊号.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("门诊号")) = .txt门诊号.Text
            End If
            If .txt住院号.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("住院号")) = .txt住院号.Text
            End If
            If .txt床号.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("床号")) = .txt床号.Text
            End If
            If .txt姓名.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("姓名")) Like "*" & .txt姓名.Text & "*"
            End If
            If .txt身份证.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("身份证号")) = .txt身份证.Text
            End If
        End With
        
        '满足则退出
        If blnFill Then
            mlngGo = i + 1
            If i <= mshPati.Rows - 1 Then mshPati.Row = i: mshPati.TopRow = i
            mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
            sta.SimpleText = "找到一条记录"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '按ESC取消
        If mblnGo = False Then
            sta.SimpleText = "用户取消定位操作"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    sta.SimpleText = "已定位到清单尾部"
    Screen.MousePointer = 0
End Sub
