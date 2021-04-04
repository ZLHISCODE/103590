VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDailyListAsk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "查询条件设置"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboPage 
      Height          =   300
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2745
      Width           =   1410
   End
   Begin VB.Frame fraUnit 
      Caption         =   "病人病区按"
      Height          =   1215
      Left            =   3360
      TabIndex        =   16
      Top             =   1440
      Width           =   1920
      Begin VB.OptionButton optUnit 
         Caption         =   "病人当前病区"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1500
      End
      Begin VB.OptionButton optUnit 
         Caption         =   "有费用的病区"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   1500
      End
   End
   Begin VB.Frame fraTime 
      Caption         =   "查询费用按"
      Height          =   1215
      Left            =   3840
      TabIndex        =   13
      Top             =   105
      Width           =   1440
      Begin VB.OptionButton opttime 
         Caption         =   "登记时间"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   380
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton opttime 
         Caption         =   "发生时间"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   800
         Width           =   1020
      End
   End
   Begin VB.Frame fraState 
      Caption         =   "病人状态"
      Height          =   1215
      Left            =   1680
      TabIndex        =   12
      Top             =   1440
      Width           =   1440
      Begin VB.CheckBox chkInOut 
         Caption         =   "出院病人"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.CheckBox chkInOut 
         Caption         =   "在院病人"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   1080
      End
   End
   Begin VB.Frame fraType 
      Caption         =   "病人类型"
      Height          =   1230
      Left            =   135
      TabIndex        =   11
      Top             =   1425
      Width           =   1440
      Begin VB.CheckBox chkPatiType 
         Caption         =   "非医保病人"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   1200
      End
      Begin VB.CheckBox chkPatiType 
         Caption         =   "医保病人"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Value           =   1  'Checked
         Width           =   1155
      End
   End
   Begin VB.Frame fraRangeSelect 
      Caption         =   "费用时间范围"
      Height          =   1215
      Left            =   135
      TabIndex        =   8
      Top             =   105
      Width           =   3600
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   900
         TabIndex        =   1
         Top             =   750
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   84017155
         CurrentDate     =   36257.9583333333
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   900
         TabIndex        =   0
         Top             =   300
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   84017155
         CurrentDate     =   36257.9597337963
      End
      Begin VB.Label lblEnd 
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblBegin 
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5385
      TabIndex        =   7
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5385
      TabIndex        =   6
      Top             =   330
      Width           =   1100
   End
   Begin VB.Label lblPage 
      AutoSize        =   -1  'True
      Caption         =   "住院次数"
      Height          =   180
      Left            =   210
      TabIndex        =   20
      Top             =   2805
      Width           =   720
   End
End
Attribute VB_Name = "frmDailyListAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mbytInFun As Byte    '0-一日清单中调用,1-病人费用查询中调用
Public mdatBegin As Date
Public mdatEnd As Date
Public mlngPageID As Long
Public mlng病人ID As Long

Public mblnAskOk As Boolean
Public mstrPrivs As String
Public mlngModul As Long
Public mblnDateMoved As Boolean '当前所选条件的数据是否在后备数据表中

Private Sub chkInOut_Click(Index As Integer)
    If chkInOut(Index).Value = 0 Then
        If chkInOut((Index + 1) Mod 2).Value = 0 Then
            chkInOut(Index).Value = 1
        End If
    End If
End Sub

Private Sub chkPatiType_Click(Index As Integer)
    If chkPatiType(Index).Value = 0 Then
        If chkPatiType((Index + 1) Mod 2).Value = 0 Then
            chkPatiType(Index).Value = 1
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnAskOk = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim lngTmp As Long
    Dim blnHavePara As Boolean
    
        
    If dtpBegin.Value >= dtpEnd.Value Then
        MsgBox "开始时间应小于结束时间。", vbInformation, gstrSysName
        Exit Sub
    End If
    blnHavePara = InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "开始时间", Format(Me.dtpBegin.Value, "hh:mm:ss"), glngSys, mlngModul, blnHavePara
    zlDatabase.SetPara "结束时间", Format(Me.dtpEnd.Value, "hh:mm:ss"), glngSys, mlngModul, blnHavePara

    lngTmp = DateDiff("d", Me.dtpEnd.Value, zlDatabase.Currentdate)
    zlDatabase.SetPara "结束间隔", lngTmp, glngSys, mlngModul, blnHavePara
    lngTmp = DateDiff("d", Me.dtpBegin.Value, Me.dtpEnd.Value)
    zlDatabase.SetPara "开始间隔", lngTmp, glngSys, mlngModul, blnHavePara
    
    
    If mbytInFun = 0 Then
        zlDatabase.SetPara "非医保病人", chkPatiType(0).Value, glngSys, mlngModul, blnHavePara
        zlDatabase.SetPara "医保病人", chkPatiType(1).Value, glngSys, mlngModul, blnHavePara
        zlDatabase.SetPara "在院病人", chkInOut(0).Value, glngSys, mlngModul, blnHavePara
        zlDatabase.SetPara "出院病人", chkInOut(1).Value, glngSys, mlngModul, blnHavePara
        zlDatabase.SetPara "病人病区模式", IIf(optUnit(0).Value = True, 0, 1), glngSys, mlngModul, blnHavePara
                
        '费用期间
        zlDatabase.SetPara "费用时间", IIf(opttime(1).Value, 1, 0), glngSys, mlngModul, blnHavePara
        
        mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), "yyyy-MM-dd HH:mm:ss"), , , Me.Caption)
    End If
    
    mdatBegin = dtpBegin.Value
    mdatEnd = dtpEnd.Value
    If cboPage.Visible Then
        mlngPageID = Val(cboPage.ItemData(cboPage.ListIndex))
    End If
    mblnAskOk = True
    Me.Hide
End Sub

Private Sub dtpBegin_Change()
    If Me.dtpBegin.Value > Me.dtpEnd.Value Then
        Me.dtpEnd.Value = Me.dtpBegin.Value
    End If
End Sub

Private Sub dtpEnd_Change()
    If Me.dtpBegin.Value > Me.dtpEnd.Value Then
        Me.dtpBegin.Value = Me.dtpEnd.Value
    End If
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim lngTmp As Long
    Dim strStartTime  As String
    Dim strEndTime As String, blnParSet As Boolean
    
    On Error Resume Next
    blnParSet = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    strEndTime = zlDatabase.GetPara("结束时间", glngSys, mlngModul, "23:59:59", Array(lblEnd, dtpEnd), blnParSet)
    lngTmp = Val(zlDatabase.GetPara("结束间隔", glngSys, mlngModul, 0, Array(lblEnd, dtpEnd), blnParSet))
    If lngTmp > 7 Then lngTmp = 7
    Me.dtpEnd.Value = CDate(Format(zlDatabase.Currentdate() - lngTmp, "yyyy-MM-dd") & " " & strEndTime)
    
    strStartTime = zlDatabase.GetPara("开始时间", glngSys, mlngModul, "00:00:00", Array(lblBegin, dtpBegin), blnParSet)
    lngTmp = Val(zlDatabase.GetPara("开始间隔", glngSys, mlngModul, 0, Array(lblBegin, dtpBegin), blnParSet))
    If lngTmp > 7 Then lngTmp = 7
    Me.dtpBegin.Value = CDate(Format(Me.dtpEnd.Value - lngTmp, "yyyy-MM-dd") & " " & strStartTime)
    
    If mbytInFun = 0 Then
        '费用期间
        lngTmp = IIf(zlDatabase.GetPara("费用时间", glngSys, mlngModul, 0, Array(opttime(0), opttime(1)), blnParSet) = "1", 1, 0)
        opttime(lngTmp).Value = True
        
        chkPatiType(0).Value = IIf(zlDatabase.GetPara("非医保病人", glngSys, mlngModul, "1", Array(chkPatiType(0)), blnParSet) = "1", 1, 0)
        chkPatiType(1).Value = IIf(zlDatabase.GetPara("医保病人", glngSys, mlngModul, "1", Array(chkPatiType(1)), blnParSet) = "1", 1, 0)
        
        lngTmp = IIf(zlDatabase.GetPara("病人病区模式", glngSys, mlngModul, "0", Array(optUnit(0), optUnit(1)), blnParSet) = "1", 1, 0)
        optUnit(lngTmp).Value = True
        
        If InStr(";" & mstrPrivs, ";出院病人查询;") = 0 Then
            chkInOut(0).Enabled = False
            chkInOut(1).Enabled = False
            chkInOut(0).Value = 1
            chkInOut(1).Value = 0
        Else
            chkInOut(0).Enabled = True
            chkInOut(1).Enabled = True
            chkInOut(0).Value = IIf(zlDatabase.GetPara("在院病人", glngSys, mlngModul, "1", Array(chkInOut(0)), blnParSet) = "1", 1, 0)
            chkInOut(1).Value = IIf(zlDatabase.GetPara("出院病人", glngSys, mlngModul, "1", Array(chkInOut(1)), blnParSet) = "1", 1, 0)
        End If
        lblPage.Visible = False
        cboPage.Visible = False
        Me.Height = 3150
    Else
        fraType.Visible = False
        fraState.Visible = False
        fraTime.Visible = False
        fraUnit.Visible = False
        cmdOk.Left = fraTime.Left
        cmdCancel.Left = fraTime.Left
        Me.Width = Me.Width - fraTime.Width
        Me.Height = Me.Height - fraType.Height - 100
        lblPage.Top = fraType.Top
        cboPage.Top = fraType.Top - 30
        Call Load住院次数(mlng病人ID, mlngPageID)
    End If
End Sub

Private Sub Load住院次数(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    Dim strSql As String, rsPage As ADODB.Recordset
    On Error GoTo errH
    strSql = "Select Distinct 主页ID From 病案主页 Where 病人ID = [1] And 病人性质 = 0 Order By 主页ID Desc"
    Set rsPage = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID)
    cboPage.Clear
    cboPage.AddItem "所有住院"
    cboPage.ItemData(cboPage.NewIndex) = 0
    Do While Not rsPage.EOF
        cboPage.AddItem "第" & Val(NVL(rsPage!主页ID)) & "次住院"
        cboPage.ItemData(cboPage.NewIndex) = Val(NVL(rsPage!主页ID))
        If Val(NVL(rsPage!主页ID)) = lng主页ID Then cboPage.ListIndex = cboPage.NewIndex
        rsPage.MoveNext
    Loop
    If cboPage.ListIndex < 0 Then cboPage.ListIndex = 0
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInFun = 0
End Sub

