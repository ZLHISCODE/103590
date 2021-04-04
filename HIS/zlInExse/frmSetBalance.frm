VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetBalance 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "结帐设置"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   ControlBox      =   0   'False
   Icon            =   "frmSetBalance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSyn 
      Caption         =   "同步"
      Height          =   375
      Left            =   4095
      TabIndex        =   23
      ToolTipText     =   "按指定住院次数的入出院时间同步费用起止时间"
      Top             =   210
      Width           =   510
   End
   Begin VB.ComboBox cboDiag 
      Height          =   300
      Left            =   675
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   6862
      Width           =   3030
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7935
      Left            =   6675
      TabIndex        =   20
      Top             =   -210
      Width           =   15
   End
   Begin MSComctlLib.TreeView tvwTime 
      Height          =   2790
      Left            =   225
      TabIndex        =   7
      Top             =   900
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   4921
      _Version        =   393217
      Indentation     =   529
      Style           =   6
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.CheckBox chkKind 
      Caption         =   "体检费用"
      Height          =   255
      Index           =   1
      Left            =   5070
      TabIndex        =   4
      Top             =   6885
      Width           =   1095
   End
   Begin VB.CheckBox chkKind 
      Caption         =   "普通费用"
      Height          =   255
      Index           =   0
      Left            =   3990
      TabIndex        =   3
      Top             =   6885
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6945
      TabIndex        =   2
      Top             =   6837
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6945
      TabIndex        =   1
      Top             =   735
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6945
      TabIndex        =   0
      Top             =   230
      Width           =   1100
   End
   Begin zl9InExse.ctlDate dtpBegin 
      Height          =   300
      Left            =   1005
      TabIndex        =   5
      Top             =   255
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   529
      Value           =   40212
      MaxDate         =   2958101
      MinDate         =   36526
   End
   Begin zl9InExse.ctlDate dtpEnd 
      Height          =   300
      Left            =   2715
      TabIndex        =   6
      Top             =   255
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   529
      Value           =   40212
      MaxDate         =   2958101
      MinDate         =   36526
   End
   Begin MSComctlLib.TreeView tvwChargeType 
      Height          =   2790
      Left            =   2370
      TabIndex        =   11
      Top             =   930
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   4921
      _Version        =   393217
      Indentation     =   529
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvwDept 
      Height          =   2790
      Left            =   4500
      TabIndex        =   13
      Top             =   930
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   4921
      _Version        =   393217
      Indentation     =   529
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvwFeeType 
      Height          =   2790
      Left            =   210
      TabIndex        =   15
      Top             =   4035
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   4921
      _Version        =   393217
      Indentation     =   529
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvwBaby 
      Height          =   2790
      Left            =   2370
      TabIndex        =   18
      Top             =   4035
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   4921
      _Version        =   393217
      Indentation     =   529
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvwItem 
      Height          =   2790
      Left            =   4500
      TabIndex        =   19
      Top             =   4035
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   4921
      _Version        =   393217
      Indentation     =   529
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.Label lblDiag 
      AutoSize        =   -1  'True
      Caption         =   "诊断"
      Height          =   180
      Left            =   210
      TabIndex        =   21
      Top             =   6922
      Width           =   360
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "收费项目"
      Height          =   180
      Left            =   4500
      TabIndex        =   17
      Top             =   3795
      Width           =   720
   End
   Begin VB.Label lblBaby 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "婴儿费"
      Height          =   180
      Left            =   2370
      TabIndex        =   16
      Top             =   3795
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "费用类型"
      Height          =   180
      Left            =   210
      TabIndex        =   14
      Top             =   3795
      Width           =   720
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "收费科室"
      Height          =   180
      Left            =   4500
      TabIndex        =   12
      Top             =   675
      Width           =   720
   End
   Begin VB.Label lbl收费类别 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "收费类别"
      Height          =   180
      Left            =   2370
      TabIndex        =   10
      Top             =   675
      Width           =   720
   End
   Begin VB.Label lbl次数 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "住院次数"
      Height          =   180
      Left            =   210
      TabIndex        =   9
      Top             =   675
      Width           =   720
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "费用时间                 至"
      Height          =   180
      Left            =   210
      TabIndex        =   8
      Top             =   315
      Width           =   2430
   End
End
Attribute VB_Name = "frmSetBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'入口参数
Private mlngInsure As Long '是否医保病人设置
Private mbytFuns As Byte '0-门诊病人;1-住院病人
Private mlng病人ID As Long '病人ID
Private mstrALLChargeType As String '收费类别
Private mstrAllTime As String
Private mstrAllUnit As String
Private mstrALLItem As String
Private mstrAllClass As String
Private mstrAllDiag As String
Private mMinDate As Date, mMaxDate As Date
Private mblnEditFee As Boolean   '是否可修改收费类别
Private mblnOK As Boolean
Private mblnNOCancel As Boolean
Private mstrUnAuditTime As String '未审核的住院次数,全部未审核时不会进入结帐设置,有“对未审核病人结帐”权限时，传入空
Private mbln门诊记帐结帐 As Boolean  'True

Private mintInsure As Integer
Private mblnDBegin As Boolean   '医保病人是否允许修改时间范围
Private mblnDEnd As Boolean
Private mblnNotClick As Boolean
Private mobjBalanceAllCon As clsBalanceAllCon
Private mobjBalanceCon As clsBalanceCon
Private mblnChange As Boolean
Private mblnNodeCheck As Boolean

Public Function ShowMe(frmMain As Object, ByVal bytFunc As Byte, ByVal lng病人ID As Long, ByVal intInsure As Integer, _
                        ByRef objBalanceAllCon As clsBalanceAllCon, ByRef objBalanceCon As clsBalanceCon) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:frmMain-调用的主窗体
    '出参:
    '返回:条件设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-13 17:31:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjBalanceAllCon = objBalanceAllCon
    Set mobjBalanceCon = objBalanceCon
    mbytFuns = bytFunc
    mlng病人ID = lng病人ID: mintInsure = intInsure
    Me.Show 1, frmMain
    ShowMe = mblnOK
End Function

Private Sub chkKind_Click(Index As Integer)
    If Visible And chkKind(0).Value = 0 And chkKind(1).Value = 0 Then
        chkKind(Index).Value = 1
    End If
    
    '仅结体检费用时,不管期间
    If chkKind(0).Value = 0 And chkKind(1).Value = 1 Then
        dtpBegin.Enabled = False
        dtpEnd.Enabled = False
    Else
        dtpBegin.Enabled = mblnDBegin
        dtpEnd.Enabled = mblnDEnd
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    mblnOK = True
    If mblnChange Then
        
        If UpdateCons = False Then Exit Sub
    End If
    Unload Me
End Sub
Private Function CheckTimeValied(ByRef strSelTimes As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前选择的住院次数的数据合法性
    '返回:数据合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-13 17:49:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, objNode As Node, objNodeTemp As Node, blnFirst As Boolean
    Dim int主页ID As Integer, int主页ID1 As Integer, intInsure1 As Integer
    Dim bln允许结多次住院费用 As Boolean, strInsureName As String, strInsureName1 As String
    Dim blnSeled As Boolean, strTimes As String
    Dim blnAll As Boolean
    
    On Error GoTo errHandle
    
    If mbytFuns = 0 Then CheckTimeValied = True: Exit Function
    
    bln允许结多次住院费用 = False
    With tvwTime
        blnFirst = True: blnSeled = False
        blnAll = True
        For Each objNode In .Nodes
            If objNode.Checked And objNode.Key <> "Root" Then
                blnSeled = True
                If zlGetTimeDataFromTimes(objNode.Tag, int主页ID1, intInsure1, strInsureName1) = False Then Exit Function
            
                If blnFirst Then
                    intInsure = intInsure1: int主页ID = int主页ID1: strInsureName = strInsureName1
                    If intInsure <> 0 Then
                        bln允许结多次住院费用 = gclsInsure.GetCapability(support允许一次结多次住院费用, mlng病人ID, intInsure)
                    End If
                Else
                    If intInsure <> 0 And bln允许结多次住院费用 = False Then
                        MsgBox "第" & int主页ID & "次住院为医保(" & strInsureName & ")住院，不允许一次结多次住院费用!", vbInformation + vbDefaultButton1, gstrSysName
                        For Each objNodeTemp In .Nodes
                            If zlGetTimeDataFromTimes(objNodeTemp.Tag, int主页ID1, intInsure1, strInsureName1) = False Then Exit Function
                            If int主页ID1 <> int主页ID Then objNodeTemp.Checked = False
                        Next
                        Exit Function
                    End If
                End If
                
                strTimes = strTimes & "," & int主页ID1
                blnFirst = False
            End If
            If objNode.Checked = False Then blnAll = False
        Next
        If Not blnSeled Then
            MsgBox "必须选择住院次数!", vbInformation + vbDefaultButton1, gstrSysName
           If tvwTime.Enabled And tvwTime.Visible Then tvwTime.SetFocus
           Exit Function
        End If
        If strTimes <> "" Then strTimes = Mid(strTimes, 2)
        strSelTimes = strTimes
    End With
    CheckTimeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Function UpdateCons() As Boolean
    On Error GoTo errH
    Dim strValue As String
    Dim i As Integer, j As Integer
    Dim strSelTimes As String
    
    If CheckTimeValied(strSelTimes) = False Then Exit Function
    
    With mobjBalanceCon
        .dtBeginDate = Format(dtpBegin.Value, "yyyy-mm-dd")
        .dtEndDate = Format(dtpEnd.Value, "yyyy-mm-dd")
    
        .strBaby = ""
        For i = 1 To tvwBaby.Nodes.Count
            If tvwBaby.Nodes(i).Checked = True Then
                If tvwBaby.Nodes(i).Key <> "Root" Then
                    .strBaby = .strBaby & "," & Mid(tvwBaby.Nodes(i).Key, 2)
                End If
            End If
        Next i
        If .strBaby <> "" Then
            .strBaby = Mid(.strBaby, 2)
        Else
            If mobjBalanceAllCon.strAllBabys <> "" Then
                MsgBox "请至少选择一项婴儿或本人费用!", vbInformation, gstrSysName
                If tvwBaby.Visible And tvwBaby.Enabled Then tvwBaby.SetFocus
                Exit Function
            End If
        End If
        
        If CheckTimeValied(strSelTimes) = False Then Exit Function
        .strTime = strSelTimes
        If .strTime = "" And mbytFuns <> 0 Then
            If mobjBalanceAllCon.strAllTime <> "" Then
                MsgBox "请至少选择一项次数!", vbInformation, gstrSysName
                If tvwTime.Visible And tvwTime.Enabled Then tvwTime.SetFocus
                Exit Function
            End If
        End If
        
        .strDeptIDs = ""
        For i = 1 To tvwDept.Nodes.Count
            If tvwDept.Nodes(i).Checked = True Then
                If tvwDept.Nodes(i).Key <> "Root" Then
                    .strDeptIDs = .strDeptIDs & "," & Mid(tvwDept.Nodes(i).Key, 2)
                End If
            End If
        Next i
        If .strDeptIDs <> "" Then
            .strDeptIDs = Mid(.strDeptIDs, 2)
        Else
            If mobjBalanceAllCon.strAllDeptIDs <> "" Then
                MsgBox "请至少选择一个科室!", vbInformation, gstrSysName
                If tvwDept.Visible And tvwDept.Enabled Then tvwDept.SetFocus
                Exit Function
            End If
        End If
        
        .strChargeType = ""
        For i = 1 To tvwChargeType.Nodes.Count
            If tvwChargeType.Nodes(i).Checked = True Then
                If tvwChargeType.Nodes(i).Key <> "Normal" And tvwChargeType.Nodes(i).Key <> "Owner" And tvwChargeType.Nodes(i).Key <> "Root" And tvwChargeType.Nodes(i).Key <> "Blood" Then
                    .strChargeType = .strChargeType & ",'" & Mid(tvwChargeType.Nodes(i).Key, 2) & "'"
                End If
            End If
        Next i
        If .strChargeType <> "" Then
            .strChargeType = Mid(.strChargeType, 2)
        Else
            If mobjBalanceAllCon.strAllChargeType <> "" Then
                MsgBox "请至少选择一项收费类别!", vbInformation, gstrSysName
                If tvwChargeType.Visible And tvwChargeType.Enabled Then tvwChargeType.SetFocus
                Exit Function
            End If
        End If
        
        .strItem = ""
        For i = 1 To tvwItem.Nodes.Count
            If tvwItem.Nodes(i).Checked = True Then
                If tvwItem.Nodes(i).Key <> "Root" Then
                    .strItem = .strItem & ",'" & tvwItem.Nodes(i).Key & "'"
                End If
            End If
        Next i
        If .strItem <> "" Then
            .strItem = Mid(.strItem, 2)
        Else
            If mobjBalanceAllCon.strAllItem <> "" Then
                MsgBox "请至少选择一项收费项目!", vbInformation, gstrSysName
                If tvwItem.Visible And tvwItem.Enabled Then tvwItem.SetFocus
                Exit Function
            End If
        End If
        
        .strDiag = cboDiag.Text
        
        .strClass = ""
        For i = 1 To tvwFeeType.Nodes.Count
            If tvwFeeType.Nodes(i).Checked = True Then
                If tvwFeeType.Nodes(i).Key <> "Root" Then
                    If tvwFeeType.Nodes(i).Key = "未知" Then
                        .strClass = .strClass & ",'无'"
                        .strClass = .strClass & ",'未知'"
                    Else
                        .strClass = .strClass & ",'" & tvwFeeType.Nodes(i).Key & "'"
                    End If
                End If
            End If
        Next i
        If .strClass <> "" Then
            .strClass = Mid(.strClass, 2)
        Else
            If mobjBalanceAllCon.strAllClass <> "" Then
                MsgBox "请至少选择一项收费项目!", vbInformation, gstrSysName
                If tvwFeeType.Visible And tvwFeeType.Enabled Then tvwFeeType.SetFocus
                Exit Function
            End If
        End If
        
        If mbytFuns = 0 Then
            .blnHealthCheckFee = chkKind(1).Value = 1
            .blnNormalFee = chkKind(0).Value = 1
            If .blnHealthCheckFee And .blnNormalFee Then
                .bytKind = 2
            Else
                If .blnHealthCheckFee Then .bytKind = 1
                If .blnNormalFee Then .bytKind = 0
            End If
        End If
    End With
    UpdateCons = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

 
Private Sub cmdSyn_Click()
    Call ResetFeeTime
End Sub
Private Sub dtpBegin_Change()
    mblnChange = True
End Sub

Private Sub dtpBegin_LastDayInput()
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpEnd_Change()
    mblnChange = True
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub dtpBegin_CmdDownClick()
    Dim dtDate  As Date
    dtDate = dtpBegin.Value
    If frmDownDate.ShowDate(dtpBegin, dtpBegin.MaxDate, dtpBegin.MinDate, dtDate) = False Then Exit Sub
    dtpBegin.Value = dtDate
    If dtpBegin.Enabled Then dtpBegin.SetFocus
End Sub

Private Sub dtpEnd_CmdDownClick()
    Dim dtDate As Date
    dtDate = dtpEnd.Value
    If frmDownDate.ShowDate(dtpEnd, dtpEnd.MaxDate, dtpEnd.MinDate, dtDate) = False Then Exit Sub
    dtpEnd.Value = dtDate
     If dtpEnd.Enabled Then dtpEnd.SetFocus
End Sub

Private Sub dtpEnd_LastDayInput()
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub InitTree()
    Dim i As Integer
    Dim intInsure As Integer, blnAll As Boolean, strInsureName As String
    Dim strTmp As String, blnAdded As Boolean
    Dim strSQL As String, rsDept As ADODB.Recordset
    Dim rsType As ADODB.Recordset, rsDay As ADODB.Recordset
    '初始化树形列表
    On Error GoTo errH
    tvwBaby.Nodes.Clear
    tvwChargeType.Nodes.Clear
    tvwDept.Nodes.Clear
    tvwFeeType.Nodes.Clear
    tvwItem.Nodes.Clear
    tvwTime.Nodes.Clear
    
    tvwBaby.Nodes.Add , , "Root", "所有费用"
    tvwChargeType.Nodes.Add , , "Owner", "所有自费类别"
    tvwChargeType.Nodes.Add , , "Normal", "所有收费类别"
    tvwDept.Nodes.Add , , "Root", "所有科室"
    tvwFeeType.Nodes.Add , , "Root", "所有类型"
    tvwItem.Nodes.Add , , "Root", "所有项目"
    
    If mbytFuns = 0 Then
        '门诊
        tvwTime.Nodes.Add , , "Root", "所有门诊"
        lbl次数.Caption = "门诊次数"
    Else
        '住院
        tvwTime.Nodes.Add , , "Root", "所有住院"
        lbl次数.Caption = "住院次数"
    End If
    
    If LoadDataPatiNumsToComBox(mlng病人ID, mobjBalanceAllCon.strAllTime, blnAll, mobjBalanceAllCon.rsAllTime, intInsure, strInsureName) = False Then Exit Sub
    
    With mobjBalanceAllCon
        dtpBegin.Value = .MinDate
        dtpEnd.Value = .MaxDate
        blnAdded = False
        For i = 0 To UBound(Split(.strAllClass, ","))
            strTmp = Replace(Split(.strAllClass, ",")(i), "'", "")
            If strTmp = "无" Or strTmp = "未知" Then
                If blnAdded = False Then
                    blnAdded = True
                    strTmp = "未知"
                    tvwFeeType.Nodes.Add "Root", tvwChild, strTmp, strTmp
                End If
            Else
                If strTmp <> "" Then
                    tvwFeeType.Nodes.Add "Root", tvwChild, strTmp, strTmp
                End If
            End If
        Next
        
        If .strAllDeptIDs <> "" Then
            strSQL = "Select A.名称,A.ID From 部门表 A,Table(f_str2list([1])) B Where A.ID=B.Column_Value"
            Set rsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .strAllDeptIDs)
            For i = 0 To UBound(Split(.strAllDeptIDs, ","))
                strTmp = Split(.strAllDeptIDs, ",")(i)
                If Val(strTmp) <> 0 Then
                    rsDept.Filter = "ID=" & Val(strTmp)
                    If Not rsDept.EOF Then tvwDept.Nodes.Add "Root", tvwChild, "K" & strTmp, Nvl(rsDept!名称)
                End If
            Next
        End If
        For i = 0 To UBound(Split(.strAllItem, ","))
            strTmp = Replace(Split(.strAllItem, ",")(i), "'", "")
            If strTmp <> "" Then
                tvwItem.Nodes.Add "Root", tvwChild, strTmp, strTmp
            End If
        Next
        strSQL = "Select 编码,类别 From 收费类别"
        Set rsType = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        For i = 0 To UBound(Split(.strAllChargeType, ","))
            If InStr("," & .strAllOwnerFeeType & ",", "," & Split(.strAllChargeType, ",")(i) & ",") = 0 Then
                strTmp = Replace(Split(.strAllChargeType, ",")(i), "'", "")
                If strTmp <> "" Then
                    rsType.Filter = "编码=" & "'" & strTmp & "'"
                    If Not rsType.EOF Then tvwChargeType.Nodes.Add "Normal", tvwChild, "K" & strTmp, Nvl(rsType!类别)
                End If
            End If
        Next
        
        For i = 0 To UBound(Split(.strAllOwnerFeeType, ","))
            strTmp = Replace(Split(.strAllOwnerFeeType, ",")(i), "'", "")
            If strTmp <> "" Then
                rsType.Filter = "编码=" & "'" & strTmp & "'"
                If Not rsType.EOF Then tvwChargeType.Nodes.Add "Owner", tvwChild, "K" & strTmp, Nvl(rsType!类别)
            End If
        Next
        
        tvwBaby.Nodes.Add "Root", tvwChild, "K0", "仅病人费用"
        For i = 0 To UBound(Split(.strAllBabys, ","))
            strTmp = Split(.strAllBabys, ",")(i)
             If Val(strTmp) <> 0 Then
                tvwBaby.Nodes.Add "Root", tvwChild, "K" & strTmp, "第" & Val(strTmp) & "个婴儿"
             End If
        Next
        
        cboDiag.Clear
        cboDiag.AddItem "所有诊断"
        cboDiag.ListIndex = cboDiag.NewIndex
        For i = 0 To UBound(Split(.strAllDiag, ","))
            strTmp = Replace(Split(.strAllDiag, ",")(i), "'", "")
            If strTmp <> "" Then
                cboDiag.AddItem strTmp
            End If
        Next
    End With
    
    For i = 1 To tvwTime.Nodes.Count
        tvwTime.Nodes.Item(i).Expanded = True
    Next i
    For i = 1 To tvwBaby.Nodes.Count
        tvwBaby.Nodes.Item(i).Expanded = True
    Next i
    For i = 1 To tvwChargeType.Nodes.Count
        tvwChargeType.Nodes.Item(i).Expanded = True
    Next i
    For i = 1 To tvwDept.Nodes.Count
        tvwDept.Nodes.Item(i).Expanded = True
    Next i
    For i = 1 To tvwFeeType.Nodes.Count
        tvwFeeType.Nodes.Item(i).Expanded = True
    Next i
    For i = 1 To tvwItem.Nodes.Count
        tvwItem.Nodes.Item(i).Expanded = True
    Next i
        
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadData()
    Dim strTmp As String, i As Integer
    Dim blnAll As Boolean, lngRootIndex As Long
    Dim blnSelf As Boolean, blnNormal As Boolean
    Dim lngSelfIndex As Long, lngNormalIndex As Long
            
    On Error GoTo errH
    If mbytFuns = 0 Then
        chkKind(0).Visible = True
        chkKind(1).Visible = True
    Else
        chkKind(0).Visible = False
        chkKind(1).Visible = False
    End If

    With mobjBalanceCon
        If .strBaby = "" Then
            For i = 1 To tvwBaby.Nodes.Count
                tvwBaby.Nodes.Item(i).Checked = True
            Next i
        Else
            For i = 0 To UBound(Split(.strBaby, ","))
                strTmp = Split(.strBaby, ",")(i)
                'If Val(strTmp) <> 0 Then
                    tvwBaby.Nodes.Item("K" & strTmp).Checked = True
                'End If
            Next
        End If
        
        blnAll = True
        For i = 1 To tvwBaby.Nodes.Count
            If tvwBaby.Nodes.Item(i).Key = "Root" Then
                lngRootIndex = i
            End If
            If Not tvwBaby.Nodes.Item(i).Checked And tvwBaby.Nodes.Item(i).Key <> "Root" Then
                blnAll = False
            End If
        Next i
        If blnAll Then tvwBaby.Nodes.Item(lngRootIndex).Checked = True
        
        If .strTime = "" Then
            For i = 1 To tvwTime.Nodes.Count
                tvwTime.Nodes.Item(i).Checked = True
            Next i
        Else
            For i = 0 To UBound(Split(.strTime, ","))
                strTmp = Split(.strTime, ",")(i)
                tvwTime.Nodes.Item("K" & strTmp).Checked = True
            Next
        End If
        If mbytFuns = 1 Then
            Call ResetFeeTime
        End If
        
        '加载住院时间
        
        
        blnAll = True
        For i = 1 To tvwTime.Nodes.Count
            If tvwTime.Nodes.Item(i).Key = "Root" Then
                lngRootIndex = i
            End If
            If Not tvwTime.Nodes.Item(i).Checked And tvwTime.Nodes.Item(i).Key <> "Root" Then
                blnAll = False
            End If
        Next i
        If blnAll Then tvwTime.Nodes.Item(lngRootIndex).Checked = True
        
        If .strDeptIDs = "" Then
            For i = 1 To tvwDept.Nodes.Count
                tvwDept.Nodes.Item(i).Checked = True
            Next i
        Else
            For i = 0 To UBound(Split(.strDeptIDs, ","))
                strTmp = Split(.strDeptIDs, ",")(i)
                If Val(strTmp) <> 0 Then
                    tvwDept.Nodes.Item("K" & strTmp).Checked = True
                End If
            Next
        End If
        
        blnAll = True
        For i = 1 To tvwDept.Nodes.Count
            If tvwDept.Nodes.Item(i).Key = "Root" Then
                lngRootIndex = i
            End If
            If Not tvwDept.Nodes.Item(i).Checked And tvwDept.Nodes.Item(i).Key <> "Root" Then
                blnAll = False
            End If
        Next i
        If blnAll Then tvwDept.Nodes.Item(lngRootIndex).Checked = True
        
        If .strItem = "" Then
            For i = 1 To tvwItem.Nodes.Count
                tvwItem.Nodes.Item(i).Checked = True
            Next i
        Else
            For i = 0 To UBound(Split(.strItem, ","))
                strTmp = Replace(Split(.strItem, ",")(i), "'", "")
                If strTmp <> "" Then
                    tvwItem.Nodes.Item(strTmp).Checked = True
                End If
            Next
        End If
        
        If .strDiag <> "" Then
            For i = 0 To cboDiag.ListCount - 1
                If cboDiag.List(i) = .strDiag Then
                    cboDiag.ListIndex = i
                    Exit For
                End If
            Next i
        End If
        
        blnAll = True
        For i = 1 To tvwItem.Nodes.Count
            If tvwItem.Nodes.Item(i).Key = "Root" Then
                lngRootIndex = i
            End If
            If Not tvwItem.Nodes.Item(i).Checked And tvwItem.Nodes.Item(i).Key <> "Root" Then
                blnAll = False
            End If
        Next i
        If blnAll Then tvwItem.Nodes.Item(lngRootIndex).Checked = True
        
        If .strClass = "" Then
            For i = 1 To tvwFeeType.Nodes.Count
                tvwFeeType.Nodes.Item(i).Checked = True
            Next i
        Else
            For i = 0 To UBound(Split(.strClass, ","))
                strTmp = Replace(Split(.strClass, ",")(i), "'", "")
                If strTmp = "无" Then strTmp = "未知"
                If strTmp <> "" Then
                    tvwFeeType.Nodes.Item(strTmp).Checked = True
                End If
            Next
        End If
        
        blnAll = True
        For i = 1 To tvwFeeType.Nodes.Count
            If tvwFeeType.Nodes.Item(i).Key = "Root" Then
                lngRootIndex = i
            End If
            If Not tvwFeeType.Nodes.Item(i).Checked And tvwFeeType.Nodes.Item(i).Key <> "Root" Then
                blnAll = False
            End If
        Next i
        If blnAll Then tvwFeeType.Nodes.Item(lngRootIndex).Checked = True
        
        If .strChargeType = "" Then
            For i = 1 To tvwChargeType.Nodes.Count
                tvwChargeType.Nodes.Item(i).Checked = True
            Next i
        Else
            For i = 0 To UBound(Split(.strChargeType, ","))
                strTmp = Replace(Split(.strChargeType, ",")(i), "'", "")
                If strTmp <> "" Then
                    tvwChargeType.Nodes.Item("K" & strTmp).Checked = True
                End If
            Next
        End If
        
        blnSelf = True
        blnNormal = True
        For i = 1 To tvwChargeType.Nodes.Count
            If tvwChargeType.Nodes.Item(i).Key = "Owner" Then
                lngSelfIndex = i
            End If
            If tvwChargeType.Nodes.Item(i).Key = "Normal" Then
                lngNormalIndex = i
            End If
            If Not tvwChargeType.Nodes.Item(i).Parent Is Nothing Then
                If Not tvwChargeType.Nodes.Item(i).Checked And tvwChargeType.Nodes.Item(i).Parent.Key = "Owner" Then
                    blnSelf = False
                End If
                If Not tvwChargeType.Nodes.Item(i).Checked And tvwChargeType.Nodes.Item(i).Parent.Key = "Normal" Then
                    blnNormal = False
                End If
            End If
        Next i
        If blnSelf Then tvwChargeType.Nodes.Item(lngSelfIndex).Checked = True
        If blnNormal Then tvwChargeType.Nodes.Item(lngNormalIndex).Checked = True
        
   
        '0-仅普通费用,1-仅体检费用,2-普通费用和体检费用
        Select Case .bytKind
        Case 0
            chkKind(0).Value = 1
            chkKind(1).Value = 0
        Case 1
            chkKind(0).Value = 0
            chkKind(1).Value = 1
        Case 2
            chkKind(0).Value = 1
            chkKind(1).Value = 1
        End Select
        
    End With
    

    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long, rsTemp As ADODB.Recordset
    Dim strTmp As String
    Dim j As Long
    mblnChange = False
    mblnOK = False
    
    If mintInsure <> 0 Then
        dtpBegin.Enabled = False
        dtpEnd.Enabled = False
    Else
        dtpBegin.Enabled = True
        dtpEnd.Enabled = True
    End If
   
    '住院次数范围
    Me.Caption = IIf(mbytFuns = 0, "门诊结帐设置", "住院结帐设置")
    Call SetControlEanbled
    Call InitTree
    Call LoadData

     
End Sub

Private Sub tvwBaby_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    mblnChange = True
    If Node.Key = "Root" Then
        With tvwBaby
            For i = 1 To .Nodes.Count
                If .Nodes.Item(i).Key <> "Root" Then
                    .Nodes.Item(i).Checked = Node.Checked
                End If
            Next i
        End With
    End If
End Sub

Private Sub tvwChargeType_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    mblnChange = True
    
    If Node.Key = "Owner" Then
        With tvwChargeType
            For i = 1 To .Nodes.Count
                If Not .Nodes.Item(i).Parent Is Nothing Then
                    If .Nodes.Item(i).Parent.Key = "Owner" Then
                        .Nodes.Item(i).Checked = Node.Checked
                    End If
                    If .Nodes.Item(i).Parent.Key = "Normal" And Node.Checked = True Then
                        .Nodes.Item(i).Checked = Not Node.Checked
                    End If
                End If
                If .Nodes.Item(i).Key = "Normal" And Node.Checked = True Then
                    .Nodes.Item(i).Checked = Not Node.Checked
                End If
            Next i
        End With
    End If
    If Node.Key = "Normal" Then
        With tvwChargeType
            For i = 1 To .Nodes.Count
                If Not .Nodes.Item(i).Parent Is Nothing Then
                    If .Nodes.Item(i).Parent.Key = "Normal" Then
                        .Nodes.Item(i).Checked = Node.Checked
                    End If
                    If .Nodes.Item(i).Parent.Key = "Owner" And Node.Checked = True Then
                        .Nodes.Item(i).Checked = Not Node.Checked
                    End If
                End If
                If .Nodes.Item(i).Key = "Owner" And Node.Checked = True Then
                    .Nodes.Item(i).Checked = Not Node.Checked
                End If
            Next i
        End With
    End If
    If Not Node.Parent Is Nothing And Node.Checked = True Then
        If Node.Parent.Key = "Normal" Then
            With tvwChargeType
                For i = 1 To .Nodes.Count
                    If .Nodes.Item(i).Key = "Owner" Then
                        .Nodes.Item(i).Checked = Not Node.Checked
                    End If
                    If Not .Nodes.Item(i).Parent Is Nothing Then
                        If .Nodes.Item(i).Parent.Key = "Owner" Then
                            .Nodes.Item(i).Checked = Not Node.Checked
                        End If
                    End If
                Next i
            End With
        End If
        
        If Node.Parent.Key = "Owner" Then
            With tvwChargeType
                For i = 1 To .Nodes.Count
                    If .Nodes.Item(i).Key = "Normal" Then
                        .Nodes.Item(i).Checked = Not Node.Checked
                    End If
                    If Not .Nodes.Item(i).Parent Is Nothing Then
                        If .Nodes.Item(i).Parent.Key = "Normal" Then
                            .Nodes.Item(i).Checked = Not Node.Checked
                        End If
                    End If
                Next i
            End With
        End If
    End If
End Sub

Private Sub tvwDept_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    mblnChange = True
    If Node.Key = "Root" Then
        With tvwDept
            For i = 1 To .Nodes.Count
                If .Nodes.Item(i).Key <> "Root" Then
                    .Nodes.Item(i).Checked = Node.Checked
                End If
            Next i
        End With
    End If
End Sub

Private Sub tvwFeeType_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    mblnChange = True
    If Node.Key = "Root" Then
        With tvwFeeType
            For i = 1 To .Nodes.Count
                If .Nodes.Item(i).Key <> "Root" Then
                    .Nodes.Item(i).Checked = Node.Checked
                End If
            Next i
        End With
    End If
End Sub

Private Sub tvwItem_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    mblnChange = True
    If Node.Key = "Root" Then
        With tvwItem
            For i = 1 To .Nodes.Count
                If .Nodes.Item(i).Key <> "Root" Then
                    .Nodes.Item(i).Checked = Node.Checked
                End If
            Next i
        End With
    End If
End Sub

Private Sub tvwTime_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    mblnChange = True
    If Node.Key = "Root" Then
        With tvwTime
            For i = 1 To .Nodes.Count
                If .Nodes.Item(i).Key <> "Root" Then
                    .Nodes.Item(i).Checked = Node.Checked
                End If
            Next i
        End With
    End If
    
    If mbytFuns <> 1 Then Exit Sub
        
    Call ResetFeeTime
End Sub

Private Sub ResetFeeTime()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新刷新费用时间
    '入参:
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-11-26 11:35:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngMax As Long, lngMin As Long
    Dim lngCurrent As Long, str主页Ids As String
    Dim strStartDate As String, strEndDate As String
    
    For i = 1 To tvwTime.Nodes.Count
        If tvwTime.Nodes.Item(i).Checked = True Then
            lngCurrent = Val(Mid(tvwTime.Nodes(i).Key, 2))
            If lngMax = 0 Then lngMax = lngCurrent
            If lngMin = 0 Then lngMin = lngCurrent
            
            If lngMax < lngCurrent Then
                lngMax = lngCurrent
            End If
            If lngMin > lngCurrent Then
                lngMin = lngCurrent
            End If
        End If
    Next
    
    If lngMin = 0 And lngMax = 0 Then
        MsgBox "请先选择住院次数!", vbInformation, Me.Caption
        Exit Sub
    End If
    str主页Ids = IIf(lngMin = lngMax, lngMax, lngMin & "," & lngMax)
    If mobjBalanceAllCon.GetPatiFeeDateRang(mlng病人ID, str主页Ids, strStartDate, strEndDate, gint费用时间 = 0) = False Then
        strStartDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        dtpBegin.Value = Format(CDate(strStartDate), "yyyy-mm-dd")
        dtpEnd.Value = strStartDate
        Exit Sub
    End If
    dtpBegin.Value = CDate(strStartDate)
    dtpEnd.Value = CDate(strEndDate)
End Sub

Private Sub tvwTime_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnChange = True
End Sub

Private Sub tvwItem_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnChange = True
End Sub

Private Sub tvwFeeType_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnChange = True
End Sub

Private Sub tvwChargeType_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnChange = True
End Sub

Private Sub tvwBaby_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnChange = True
End Sub

Private Sub tvwDept_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnChange = True
End Sub

Private Function LoadDataPatiNumsToComBox(ByVal lng病人ID As Long, ByVal str主页Ids As String, ByRef blnAllSel As Boolean, _
    ByRef rsTimeAll As ADODB.Recordset, ByRef intInsure As Integer, Optional strInsureName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载住院次数，给下拉列表框
    '入参: str主页IDs-所有住院次数,用逗号分隔
    '出参:blnAllSel-当前是否选择了所有住院次数
    '     intInsure-返回第一个选择的医保序号
    '     strInsureName-返回第一个选择的医保名称
    '返回:加载成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-13 11:23:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, int主页ID As Long, strTag As String
    Dim i As Integer, intInsure1 As Integer, strInsureName1 As String
    Dim objNode As Node
    
    On Error GoTo errHandle
    
    tvwTime.Nodes.Clear
    'mbytFuns As Byte '0-门诊病人;1-住院病人
    If mbytFuns <> 1 Then
        tvwTime.Nodes.Add , , "Root", "所有门诊"
        varTemp = Split(str主页Ids, ",")
        blnAllSel = True
        For i = 0 To UBound(varTemp)
            If Val(varTemp(i)) = 0 Then
                tvwTime.Nodes.Add "Root", tvwChild, "K0", "普通门诊"
            Else
                Set objNode = tvwTime.Nodes.Add("Root", tvwChild, "K" & i, "第" & Val(varTemp(i)) & "次留观")
                objNode.Tag = i
            End If
        Next
        Call tvwTime.Refresh
        LoadDataPatiNumsToComBox = True
        Exit Function
    End If
    
    tvwTime.Nodes.Add , , "Root", "所有住院"
    '获取当前未结住院次所涉及的医保数据集
    If rsTimeAll Is Nothing Then
        Call mobjBalanceAllCon.zlGetTimeRecordFromTimeString(lng病人ID, str主页Ids, rsTimeAll)
    End If
    
    '加载住院次数文本框
    Dim blnSelect As Boolean
    
    rsTimeAll.Filter = 0
    
    With rsTimeAll
        intInsure = 0
        If .RecordCount <> 0 Then
            .MoveFirst:  intInsure = Val(Nvl(!险类)): strInsureName = Nvl(!保险名称)
        End If
        
        i = 1: blnAllSel = True
        Do While Not .EOF
            '自费的，先缺省全选,最后一次住院为医保的，则先结医保的
            blnSelect = mobjBalanceAllCon.strAllOwnerFeeType <> "" Or (intInsure <> 0 And i = 1) Or intInsure = 0
            If blnAllSel And Not blnSelect Then blnAllSel = False
            int主页ID = Val(Nvl(!主页ID)): intInsure1 = Val(Nvl(!险类)): strInsureName1 = Nvl(!保险名称)
            strTag = int主页ID & "|" & Val(Nvl(!险类)) & "|" & Nvl(!保险名称)
            Set objNode = tvwTime.Nodes.Add("Root", tvwChild, "K" & int主页ID, "第" & int主页ID & "次住院" & IIf(Val(Nvl(!险类)) <> 0, "(医保)", ""))
            objNode.Tag = strTag
            .MoveNext
        Loop
     End With
    LoadDataPatiNumsToComBox = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetControlEanbled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件Eanbled属性
    '编制:刘兴洪
    '日期:2017-12-29 10:07:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEnabled As Boolean
    
    On Error GoTo errHandle
    '医保病人只能设置住院次数和费用期间
    If mintInsure > 0 Then
        If mbytFuns = 0 Then   '刘兴洪:25435
            dtpBegin.Enabled = False
            mblnDBegin = dtpBegin.Enabled
            tvwTime.Enabled = False
            tvwDept.Enabled = True
            tvwItem.Enabled = True
            tvwFeeType.Enabled = True
            dtpBegin.Enabled = True
            tvwBaby.Enabled = False
        Else
            tvwBaby.Enabled = gclsInsure.GetCapability(support结帐_设置婴儿费条件, mlng病人ID, mintInsure)
            tvwDept.Enabled = gclsInsure.GetCapability(support结帐_指定科室, mlng病人ID, mintInsure)
            tvwItem.Enabled = gclsInsure.GetCapability(support结帐_指定费用项目, mlng病人ID, mintInsure)
            tvwFeeType.Enabled = gclsInsure.GetCapability(support结帐_指定费用类型, mlng病人ID, mintInsure)
            tvwTime.Enabled = gclsInsure.GetCapability(support结帐_指定住院次数, mlng病人ID, mintInsure) Or gclsInsure.GetCapability(support允许一次结多次住院费用, mlng病人ID, mintInsure)
            dtpBegin.Enabled = False
            dtpEnd.Enabled = gclsInsure.GetCapability(support结帐_指定日期范围, mlng病人ID, mintInsure)
        End If
        mblnDBegin = dtpBegin.Enabled
        mblnDEnd = dtpEnd.Enabled
    Else
        mblnDBegin = True
        mblnDEnd = True
    End If
    
    If mbytFuns <> 0 Then
         chkKind(1).Visible = False
    End If
    
    cmdSyn.Enabled = tvwTime.Enabled
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub



