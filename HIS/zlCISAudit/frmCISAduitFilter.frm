VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISAduitFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤条件"
   ClientHeight    =   7695
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7830
   Icon            =   "frmCISAduitFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame frmFind 
      Caption         =   "过滤条件"
      Height          =   6855
      Left            =   210
      TabIndex        =   42
      Top             =   105
      Width           =   7410
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   4
         Left            =   1365
         TabIndex        =   7
         Top             =   870
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   5
         Left            =   5235
         TabIndex        =   8
         Top             =   855
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin VB.CommandButton cmd药品信息 
         Height          =   300
         Left            =   6945
         Picture         =   "frmCISAduitFilter.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   6375
         Width           =   300
      End
      Begin VB.CommandButton cmd疾病名称 
         Height          =   300
         Left            =   6930
         Picture         =   "frmCISAduitFilter.frx":685E
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   5940
         Width           =   300
      End
      Begin VB.CommandButton cmd检查类型 
         Height          =   300
         Left            =   3045
         Picture         =   "frmCISAduitFilter.frx":D0B0
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   6375
         Width           =   300
      End
      Begin VB.CommandButton cmd住院医师 
         Height          =   300
         Left            =   3045
         Picture         =   "frmCISAduitFilter.frx":13902
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   5940
         Width           =   300
      End
      Begin VB.Frame Frame3 
         Caption         =   "病人类型"
         Height          =   765
         Left            =   285
         TabIndex        =   22
         Top             =   3375
         Width           =   6960
         Begin VB.OptionButton opt 
            Caption         =   "医保病人(&F)"
            Height          =   240
            Index           =   2
            Left            =   4665
            TabIndex        =   25
            Top             =   360
            Width           =   1470
         End
         Begin VB.OptionButton opt 
            Caption         =   "非医保病人(&E)"
            Height          =   195
            Index           =   1
            Left            =   2475
            TabIndex        =   24
            Top             =   360
            Width           =   1500
         End
         Begin VB.OptionButton opt 
            Caption         =   "所有病人(&D)"
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   23
            Top             =   345
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "病案状态"
         Height          =   945
         Left            =   285
         TabIndex        =   16
         Top             =   2310
         Width           =   6960
         Begin VB.CheckBox chk 
            Caption         =   "提交待收(&6)"
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   47
            Top             =   255
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chk 
            Caption         =   "接收待审(&7)"
            Height          =   195
            Index           =   0
            Left            =   2490
            TabIndex        =   17
            Top             =   255
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chk 
            Caption         =   "审查归档(&B)"
            Height          =   195
            Index           =   1
            Left            =   4680
            TabIndex        =   19
            Top             =   600
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chk 
            Caption         =   "正在审查(&8)"
            Height          =   195
            Index           =   2
            Left            =   4680
            TabIndex        =   18
            Top             =   270
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chk 
            Caption         =   "审查反馈(&9)"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   20
            Top             =   585
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chk 
            Caption         =   "审查整改(&A)"
            Height          =   195
            Index           =   4
            Left            =   2490
            TabIndex        =   21
            Top             =   600
            Value           =   1  'Checked
            Width           =   1335
         End
      End
      Begin VB.ListBox lst 
         Height          =   1320
         Left            =   1365
         Style           =   1  'Checkbox
         TabIndex        =   27
         Top             =   4290
         Width           =   5880
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   5235
         TabIndex        =   15
         Top             =   1815
         Width           =   2010
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1365
         TabIndex        =   13
         Top             =   1815
         Width           =   2010
      End
      Begin VB.TextBox txt住院医师 
         Height          =   300
         Left            =   1365
         TabIndex        =   29
         Top             =   5940
         Width           =   1650
      End
      Begin VB.TextBox txt疾病名称 
         Height          =   300
         Left            =   5235
         TabIndex        =   32
         Top             =   5940
         Width           =   1650
      End
      Begin VB.TextBox txt检查类型 
         Height          =   300
         Left            =   1365
         TabIndex        =   35
         Top             =   6375
         Width           =   1650
      End
      Begin VB.TextBox txt药品信息 
         Height          =   300
         Left            =   5235
         TabIndex        =   38
         Top             =   6375
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   1365
         TabIndex        =   1
         Top             =   360
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   5235
         TabIndex        =   2
         Top             =   375
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   6
         Left            =   1365
         TabIndex        =   10
         Top             =   1320
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   7
         Left            =   5235
         TabIndex        =   11
         Top             =   1320
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   3
         Left            =   5235
         TabIndex        =   5
         Top             =   780
         Visible         =   0   'False
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   2
         Left            =   1365
         TabIndex        =   4
         Top             =   750
         Visible         =   0   'False
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   1
         Left            =   4110
         TabIndex        =   45
         Top             =   900
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "出院病人(&2)"
         Height          =   180
         Index           =   0
         Left            =   285
         TabIndex        =   6
         Top             =   915
         Width           =   990
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   10995
         Y1              =   5775
         Y2              =   5775
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000005&
         X1              =   15
         X2              =   10995
         Y1              =   5790
         Y2              =   5805
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   11
         Left            =   4110
         TabIndex        =   46
         Top             =   1395
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "医嘱时间(&3)"
         Height          =   180
         Index           =   10
         Left            =   285
         TabIndex        =   9
         Top             =   1365
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   9
         Left            =   4110
         TabIndex        =   44
         Top             =   405
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "提交时间(&1)"
         Height          =   180
         Index           =   8
         Left            =   285
         TabIndex        =   0
         Top             =   420
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保种类(&G)"
         Height          =   180
         Index           =   6
         Left            =   285
         TabIndex        =   26
         Top             =   4305
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前病况(&4)"
         Height          =   180
         Index           =   2
         Left            =   285
         TabIndex        =   12
         Top             =   1860
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院情况(&5)"
         Height          =   180
         Index           =   4
         Left            =   4110
         TabIndex        =   14
         Top             =   1860
         Width           =   990
      End
      Begin VB.Label lab住院医师 
         AutoSize        =   -1  'True
         Caption         =   "住院医师(&H)"
         Height          =   180
         Left            =   285
         TabIndex        =   28
         Top             =   6000
         Width           =   990
      End
      Begin VB.Label lab疾病名称 
         AutoSize        =   -1  'True
         Caption         =   "疾病名称(&I)"
         Height          =   180
         Left            =   4110
         TabIndex        =   31
         Top             =   6000
         Width           =   990
      End
      Begin VB.Label lab检查类型 
         AutoSize        =   -1  'True
         Caption         =   "检查类型(&J)"
         Height          =   180
         Left            =   285
         TabIndex        =   34
         Top             =   6435
         Width           =   990
      End
      Begin VB.Label lab药品信息 
         AutoSize        =   -1  'True
         Caption         =   "药品信息(&K)"
         Height          =   180
         Left            =   4110
         TabIndex        =   37
         Top             =   6435
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "归档病人(&2)"
         Height          =   180
         Index           =   5
         Left            =   285
         TabIndex        =   3
         Top             =   795
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   7
         Left            =   4110
         TabIndex        =   43
         Top             =   825
         Visible         =   0   'False
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5190
      TabIndex        =   40
      Top             =   7140
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6480
      TabIndex        =   41
      Top             =   7140
      Width           =   1100
   End
End
Attribute VB_Name = "frmCISAduitFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################

Private mrsParam As New ADODB.Recordset
Private mblnDataChanged As Boolean
Private mblnOK As Boolean

'######################################################################################################################

Public Function ShowPara(ByVal frmMain As Object, ByRef rsParam As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnOK = False
    
    Set mrsParam = CopyRecordStruct(rsParam)
    Call CopyRecordData(rsParam, mrsParam)
                
    If ExecuteCommand("初始参数") = False Then Exit Function
    If ExecuteCommand("读取参数") = False Then Exit Function
    
    Me.Show 1, frmMain
    
    If mblnOK Then
        Call DeleteRecordData(rsParam)
        Call CopyRecordData(mrsParam, rsParam)
        ShowPara = mblnOK
    End If
    
End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始参数"
        chk(0).Value = 1
        chk(1).Value = 1
        chk(2).Value = 1
        chk(3).Value = 1
        chk(4).Value = 1
        chk(5).Value = 1
        opt(0).Value = True
        
        dtp(0).Value = Format(zlDatabase.Currentdate, dtp(0).CustomFormat)
        dtp(1).Value = Format(zlDatabase.Currentdate, dtp(1).CustomFormat)
        dtp(2).Value = Format(zlDatabase.Currentdate, dtp(2).CustomFormat)
        dtp(3).Value = Format(zlDatabase.Currentdate, dtp(3).CustomFormat)

        dtp(4).Value = Format(zlDatabase.Currentdate, dtp(4).CustomFormat)
        dtp(5).Value = Format(zlDatabase.Currentdate, dtp(5).CustomFormat)
        
        dtp(6).Value = Format(zlDatabase.Currentdate, dtp(6).CustomFormat)
        dtp(7).Value = Format(zlDatabase.Currentdate, dtp(7).CustomFormat)
        
        cbo(0).Clear
        cbo(0).AddItem ""
        Set rs = gclsPackage.GetBaseCode("病情")
        If rs.BOF = False Then
            Call AddComboData(cbo(0), rs, "名称", "名称", , False)
        End If
                
        cbo(1).Clear
        cbo(1).AddItem ""
        Set rs = gclsPackage.GetBaseCode("治疗结果")
        If rs.BOF = False Then
            Call AddComboData(cbo(1), rs, "名称", "名称", , False)
        End If
        
        lst.Clear
        Set rs = gclsPackage.GetInsureKind()
        If rs.BOF = False Then
            Do While Not rs.EOF
                lst.AddItem rs("名称").Value
                lst.ItemData(lst.NewIndex) = rs("序号").Value
                rs.MoveNext
            Loop
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取参数"
        
        chk(0).Value = Val(ParamRead(mrsParam, "接收待审"))
        chk(1).Value = Val(ParamRead(mrsParam, "拒绝接收"))
        chk(2).Value = Val(ParamRead(mrsParam, "正在审查"))
        chk(3).Value = Val(ParamRead(mrsParam, "审查反馈"))
        chk(4).Value = Val(ParamRead(mrsParam, "审查整改"))
        chk(5).Value = Val(ParamRead(mrsParam, "提交待收"))
        cbo(0).Text = Trim(ParamRead(mrsParam, "当前病况"))
        cbo(1).Text = Trim(ParamRead(mrsParam, "出院情况"))
        
        txt住院医师.Text = Trim(ParamRead(mrsParam, "住院医师"))
        txt疾病名称.Tag = Trim(ParamRead(mrsParam, "疾病名称"))
        txt检查类型.Text = Trim(ParamRead(mrsParam, "检查类型"))
        txt药品信息.Tag = Trim(ParamRead(mrsParam, "药品信息"))
        
        '读取疾病名称
        If txt疾病名称.Tag <> "" Then
            gstrSQL = "Select 编码,名称 From 疾病编码目录 where ID = [1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txt疾病名称.Tag)
            If rs Is Nothing Then
                txt疾病名称.Tag = ""
            ElseIf rs.EOF Or rs.BOF Then
                txt疾病名称.Tag = ""
            Else
                txt疾病名称.Text = rs!编码 & "[" & rs.Fields!名称 & "]"
            End If
        End If
        '读取药品名称
        If txt药品信息.Tag <> "" Then
        
            gstrSQL = "select 名称 from 药品目录 where 药品ID = [1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(txt药品信息.Tag))
            If rs Is Nothing Then
                txt药品信息.Tag = ""
            ElseIf rs.EOF Or rs.BOF Then
                txt药品信息.Tag = ""
            Else
                txt药品信息.Text = rs.Fields!名称
            End If
        End If
        
        intLoop = Val(ParamRead(mrsParam, "病人类型"))
        If intLoop >= 0 And intLoop <= 2 Then opt(intLoop).Value = True
        
        strTmp = Trim(ParamRead(mrsParam, "医保种类"))
        If strTmp <> "" Then
            strTmp = "," & strTmp & ","
            For intLoop = 0 To lst.ListCount - 1
                If InStr(strTmp, "," & lst.ItemData(intLoop) & ",") > 0 Then
                    lst.Selected(intLoop) = True
                End If
            Next
        End If
        
        If ParamRead(mrsParam, "审查开始时间") <> "" Then dtp(0).Value = Format(ParamRead(mrsParam, "审查开始时间"), dtp(0).CustomFormat)
        If ParamRead(mrsParam, "审查结束时间") <> "" Then dtp(1).Value = Format(ParamRead(mrsParam, "审查结束时间"), dtp(1).CustomFormat)
        If ParamRead(mrsParam, "归档开始时间") <> "" Then dtp(2).Value = Format(ParamRead(mrsParam, "归档开始时间"), dtp(2).CustomFormat)
        If ParamRead(mrsParam, "归档结束时间") <> "" Then dtp(3).Value = Format(ParamRead(mrsParam, "归档结束时间"), dtp(3).CustomFormat)
        If ParamRead(mrsParam, "出院开始时间") <> "" Then dtp(4).Value = Format(ParamRead(mrsParam, "出院开始时间"), dtp(4).CustomFormat)
        If ParamRead(mrsParam, "出院结束时间") <> "" Then dtp(5).Value = Format(ParamRead(mrsParam, "出院结束时间"), dtp(5).CustomFormat)
        If ParamRead(mrsParam, "出院开始时间") <> "" Then dtp(6).Value = Format(ParamRead(mrsParam, "医嘱开始时间"), dtp(6).CustomFormat)
        If ParamRead(mrsParam, "出院结束时间") <> "" Then dtp(7).Value = Format(ParamRead(mrsParam, "医嘱结束时间"), dtp(7).CustomFormat)
        
        lst.Enabled = opt(2).Value
        DataChanged = False
    '------------------------------------------------------------------------------------------------------------------
    Case "校验数据"
        
        If chk(0).Value = 0 And chk(1).Value = 0 And chk(2).Value = 0 And chk(3).Value = 0 And chk(4).Value = 0 And chk(5).Value = 0 Then
            ShowSimpleMsg "接收待审、拒绝接收、正在审查和审查反馈必须选择一项！"
            chk(0).SetFocus
            Exit Function
        End If
        
        If Abs(DateDiff("m", dtp(1).Value, dtp(0).Value)) > 3 Or Abs(DateDiff("m", dtp(3).Value, dtp(2).Value)) > 3 Or Abs(DateDiff("m", dtp(5).Value, dtp(4).Value)) > 3 Then
            If MsgBox("您设置的时间范围超过了3个月，可能会很慢，是否继续？", vbYesNo + vbDefaultButton2, ParamInfo.产品名称) = vbNo Then
                Exit Function
            End If
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "保存参数"

        
        Call ParamWrite(mrsParam, "接收待审", chk(0).Value)
        Call ParamWrite(mrsParam, "拒绝接收", chk(1).Value)
        Call ParamWrite(mrsParam, "正在审查", chk(2).Value)
        Call ParamWrite(mrsParam, "审查反馈", chk(3).Value)
        Call ParamWrite(mrsParam, "审查整改", chk(4).Value)
        Call ParamWrite(mrsParam, "提交待收", chk(5).Value)
        Call ParamWrite(mrsParam, "当前病况", cbo(0).Text)
        Call ParamWrite(mrsParam, "出院情况", cbo(1).Text)
        
        If opt(0).Value Then
            Call ParamWrite(mrsParam, "病人类型", 0)
        ElseIf opt(1).Value Then
            Call ParamWrite(mrsParam, "病人类型", 1)
        Else
            Call ParamWrite(mrsParam, "病人类型", 2)
        End If
        
        strTmp = ""
        For intLoop = 0 To lst.ListCount - 1
            If lst.Selected(intLoop) Then
                strTmp = strTmp & "," & lst.ItemData(intLoop)
            End If
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        Call ParamWrite(mrsParam, "医保种类", strTmp)
        
        Call ParamWrite(mrsParam, "审查开始时间", Format(dtp(0).Value, dtp(0).CustomFormat))
        Call ParamWrite(mrsParam, "审查结束时间", Format(dtp(1).Value, dtp(1).CustomFormat))
        Call ParamWrite(mrsParam, "归档开始时间", Format(dtp(2).Value, dtp(2).CustomFormat))
        Call ParamWrite(mrsParam, "归档结束时间", Format(dtp(3).Value, dtp(3).CustomFormat))
        Call ParamWrite(mrsParam, "出院开始时间", Format(dtp(4).Value, dtp(4).CustomFormat))
        Call ParamWrite(mrsParam, "出院结束时间", Format(dtp(5).Value, dtp(5).CustomFormat))
        Call ParamWrite(mrsParam, "医嘱开始时间", Format(dtp(6).Value, dtp(6).CustomFormat))
        Call ParamWrite(mrsParam, "医嘱结束时间", Format(dtp(7).Value, dtp(7).CustomFormat))
        
        Call ParamWrite(mrsParam, "住院医师", txt住院医师.Text)
        Call ParamWrite(mrsParam, "疾病名称", IIf(txt疾病名称.Text = "", "", txt疾病名称.Tag))
        Call ParamWrite(mrsParam, "检查类型", txt检查类型.Text)
        Call ParamWrite(mrsParam, "药品信息", IIf(txt药品信息.Text = "", "", txt药品信息.Tag))
        
    End Select
    
    ExecuteCommand = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbo_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

'######################################################################################################################

Private Sub chk_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    If DataChanged Then
        If ExecuteCommand("校验数据") = False Then Exit Sub
        If ExecuteCommand("保存参数") Then
            mrsParam.Filter = ""
            mblnOK = True
            DataChanged = False
        End If
    End If
    Unload Me
End Sub

Private Sub cmd疾病名称_Click()
On Error GoTo ErrH
    SelectSick
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmd检查类型_Click()
On Error GoTo ErrH
    SelectCheck
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmd药品信息_Click()
On Error GoTo ErrH
    SelectDrug
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmd住院医师_Click()
On Error GoTo ErrH
    SelectDoctor
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub dtp_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lst_ItemCheck(Item As Integer)
    DataChanged = True
End Sub

Private Sub lst_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub opt_Click(Index As Integer)
    DataChanged = True
    
    lst.Enabled = (Index = 2)
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt疾病名称_Change()
    DataChanged = True
End Sub

Private Sub txt检查类型_Change()
    DataChanged = True
End Sub

Private Sub txt药品信息_Change()
    DataChanged = True
End Sub

Private Sub txt住院医师_Change()
    DataChanged = True
End Sub

Private Sub txt住院医师_KeyPress(KeyAscii As Integer)
    If Trim(txt住院医师.Text) = "" Then Exit Sub
    If KeyAscii = 13 Then
        SelectDoctor txt住院医师.Text
    End If
End Sub

Private Sub txt疾病名称_KeyPress(KeyAscii As Integer)
    If Trim(txt疾病名称.Text) = "" Then Exit Sub
    If KeyAscii = 13 Then
        SelectSick txt疾病名称.Text
    End If
End Sub

Private Sub txt药品信息_KeyPress(KeyAscii As Integer)
    If Trim(txt药品信息.Text) = "" Then Exit Sub
    If KeyAscii = 13 Then
        SelectDrug txt药品信息.Text
    End If
End Sub

Private Sub txt检查类型_KeyPress(KeyAscii As Integer)

    If Trim(txt检查类型.Text) = "" Then Exit Sub
    If KeyAscii = 13 Then
        SelectCheck txt检查类型.Text
    End If

End Sub

'选择医生
Private Sub SelectDoctor(Optional strShortName As String = "")
    Dim rsTmp           As ADODB.Recordset
    Dim rsResult        As ADODB.Recordset
    Dim bytRet          As Byte
On Error GoTo ErrH
    gstrSQL = ""
    If strShortName <> "" Then
        gstrSQL = gstrSQL & vbCrLf & "Select c.ID,c.编号,c.姓名 As 名称"
        gstrSQL = gstrSQL & vbCrLf & "From 人员表 C, 人员性质说明 D"
        gstrSQL = gstrSQL & vbCrLf & "Where  c.id = d.人员id And D.人员性质 = '医生'"
        gstrSQL = gstrSQL & vbCrLf & "And (c.姓名 like '%'||[1]||'%' or 简码 like '%'||[1]||'%')"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(strShortName))
        bytRet = ShowPubSelect(Me, txt住院医师, 2, "编号,1200,0,;名称,1200,0,", Me.Name & "\医生选择", "请从下表中选择一个医生", rsTmp, rsResult, 8790, 4500, False)
    Else
        gstrSQL = gstrSQL & vbCrLf & "Select Id,上级id,0 As 末级,编码 as 编号,名称 From 部门表"
        gstrSQL = gstrSQL & vbCrLf & "Start With 上级id Is Null"
        gstrSQL = gstrSQL & vbCrLf & "Connect By Prior ID = 上级id"
        gstrSQL = gstrSQL & vbCrLf & "Union All"
        gstrSQL = gstrSQL & vbCrLf & "Select c.id,b.部门id As 上级Id,1 As 末级,c.编号,c.姓名 As 名称"
        gstrSQL = gstrSQL & vbCrLf & "From 部门人员 b,人员表 C, 人员性质说明 D"
        gstrSQL = gstrSQL & vbCrLf & "Where c.Id=b.人员id and c.id = d.人员id And D.人员性质 = '医生' And b.缺省=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        bytRet = ShowPubSelect(Me, txt住院医师, 3, "编号,1200,0,;名称,1200,0,", Me.Name & "\医生选择", "请从下表中选择一个医生", rsTmp, rsResult, 8790, 4500, False)
    End If
    
    If rsResult Is Nothing Then
        txt住院医师.Text = ""
    ElseIf rsResult.EOF Or rsResult.BOF Then
        txt住院医师.Text = ""
    Else
        txt住院医师.Text = rsResult!名称
    End If
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub

'选择疾病
Private Sub SelectSick(Optional strShortName As String = "")
    Dim rsTmp           As ADODB.Recordset
    Dim rsResult        As ADODB.Recordset
    Dim bytRet          As Byte
On Error GoTo ErrH
    gstrSQL = ""
    gstrSQL = gstrSQL & vbCrLf & "select ID,编码,名称 from 疾病编码目录"
    If strShortName <> "" Then
        gstrSQL = gstrSQL & vbCrLf & "Where (编码 like '%'||[1]||'%' or 名称 like '%'||[1]||'%' or 简码 like '%'||[1]||'%')"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(strShortName))
    bytRet = ShowPubSelect(Me, txt住院医师, 2, "编码,1200,0,;名称,1200,0,", Me.Name & "\疾病选择", "请从下表中选择一种疾病", rsTmp, rsResult, 8790, 4500, False)
    
    If rsResult Is Nothing Then
        txt疾病名称.Text = ""
        txt疾病名称.Tag = ""
    ElseIf rsResult.EOF Or rsResult.BOF Then
        txt疾病名称.Text = ""
        txt疾病名称.Tag = ""
    Else
        txt疾病名称.Tag = rsResult!ID
        txt疾病名称.Text = rsResult!编码 & "[" & rsResult!名称 & "]"
    End If
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub

'选择药品
Private Sub SelectDrug(Optional strShortName As String = "")
    Dim rsTmp           As ADODB.Recordset
    Dim rsResult        As ADODB.Recordset
    Dim bytRet          As Byte
On Error GoTo ErrH
    gstrSQL = ""
    If strShortName = "" Then
        gstrSQL = gstrSQL & vbCrLf & "select 5 as ID,null as 上级ID,0 as 末级,'5' as 编码,'西成药' as 名称,'西成药' as 通用名称,'' as 规格 from dual union all"
        gstrSQL = gstrSQL & vbCrLf & "select 6 as ID,null as 上级ID,0 as 末级,'6' as 编码,'中成药' as 名称,'中成药' as 通用名称,'' as 规格 from dual union all"
        gstrSQL = gstrSQL & vbCrLf & "select 7 as ID,null as 上级ID,0 as 末级,'7' as 编码,'中草药' as 名称,'中草药' as 通用名称,'' as 规格 from dual union all"
        gstrSQL = gstrSQL & vbCrLf & "Select * from("
        gstrSQL = gstrSQL & vbCrLf & "select a.药品ID as ID,decode(b.材质分类,'西成药',5,'中成药',6,7) as 上级ID,1 as 末级,a.编码 as 编码,a.名称 as 名称,b.通用名称 as 通用名称 ,a.规格"
        gstrSQL = gstrSQL & vbCrLf & "from 药品目录 a,药品信息 b"
        gstrSQL = gstrSQL & vbCrLf & "Where a.药名ID = b.药名ID"
        gstrSQL = gstrSQL & vbCrLf & "Order by b.材质分类"
        gstrSQL = gstrSQL & vbCrLf & ")"
    Else
        gstrSQL = gstrSQL & vbCrLf & "select 5 as ID,null as 上级ID,0 as 末级,'5' as 编码,'西成药' as 名称,'西成药' as 通用名称,'' as 规格 from dual union all"
        gstrSQL = gstrSQL & vbCrLf & "select 6 as ID,null as 上级ID,0 as 末级,'6' as 编码,'中成药' as 名称,'中成药' as 通用名称,'' as 规格 from dual union all"
        gstrSQL = gstrSQL & vbCrLf & "select 7 as ID,null as 上级ID,0 as 末级,'7' as 编码,'中草药' as 名称,'中草药' as 通用名称,'' as 规格 from dual union all"
        gstrSQL = gstrSQL & vbCrLf & "Select * from("
        gstrSQL = gstrSQL & vbCrLf & "select a.药品ID as ID,decode(b.材质分类,'西成药',5,'中成药',6,7) as 上级ID,1 as 末级,a.编码 as 编码,a.名称 as 名称,b.通用名称 as 通用名称 ,a.规格"
        gstrSQL = gstrSQL & vbCrLf & "from 药品目录 a,药品信息 b"
        gstrSQL = gstrSQL & vbCrLf & "Where a.药名ID = b.药名ID"
        gstrSQL = gstrSQL & vbCrLf & "And (a.编码 like '%' || [1] || '%' or a.名称  like '%' || [1] || '%' or zlSpellCode(a.名称)  like '%' || [1] || '%')"
        gstrSQL = gstrSQL & vbCrLf & "Order by b.材质分类"
        gstrSQL = gstrSQL & vbCrLf & ")"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(strShortName))
    bytRet = ShowPubSelect(Me, txt住院医师, 3, "编码,1200,0,;名称,1200,0,;通用名称,1200,0,;规格,800,0,", Me.Name & "\疾病选择", "请从下表中选择一种疾病", rsTmp, rsResult, 8790, 4500, False)
    
    If rsResult Is Nothing Then
        txt药品信息.Text = ""
        txt药品信息.Tag = ""
    ElseIf rsResult.EOF Or rsResult.BOF Then
        txt药品信息.Text = ""
        txt药品信息.Tag = ""
    Else
        txt药品信息.Tag = rsResult!ID
        txt药品信息.Text = rsResult!名称
    End If
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub

'选择检查类型
Private Sub SelectCheck(Optional strShortName As String = "")
    Dim rsTmp           As ADODB.Recordset
    Dim rsResult        As ADODB.Recordset
    Dim bytRet          As Byte
On Error GoTo ErrH
    gstrSQL = ""
    gstrSQL = gstrSQL & vbCrLf & "select 编码 as ID,编码,名称 from 诊疗检查类型"
    If strShortName <> "" Then
        gstrSQL = gstrSQL & vbCrLf & "Where (编码 like '%'||[1]||'%' or 名称 like '%'||[1]||'%' or 简码  like '%' || [1] || '%')"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(strShortName))
    bytRet = ShowPubSelect(Me, txt住院医师, 2, "编码,1200,0,;名称,1200,0,", Me.Name & "\检查类型选择", "请从下表中选择一种检查类型", rsTmp, rsResult, 8790, 4500, False)
    
    If rsResult Is Nothing Then
        txt检查类型.Text = ""
    ElseIf rsResult.EOF Or rsResult.BOF Then
        txt检查类型.Text = ""
    Else
        txt检查类型.Text = rsResult!名称
    End If
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub

