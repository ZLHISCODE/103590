VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiPressMoneySet 
   Caption         =   "报警方案设置"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11010
   Icon            =   "frmPatiPressMoneySet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   11010
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic站点 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   45
      ScaleHeight     =   390
      ScaleWidth      =   5625
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1260
      Width           =   5625
      Begin VB.ComboBox cbo站点 
         Height          =   300
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   60
         Width           =   3195
      End
      Begin VB.Label lbl站点 
         AutoSize        =   -1  'True
         Caption         =   "站点"
         Height          =   180
         Left            =   90
         TabIndex        =   17
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox picList类别 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3555
      Left            =   4455
      ScaleHeight     =   3525
      ScaleWidth      =   2310
      TabIndex        =   8
      Top             =   2670
      Visible         =   0   'False
      Width           =   2340
      Begin VB.ListBox lst类别 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   3180
         Left            =   -30
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   360
         Width           =   2355
      End
      Begin XtremeSuiteControls.ShortcutCaption shtCaption 
         Height          =   360
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2325
         _Version        =   589884
         _ExtentX        =   4101
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "类别选择"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16744576
         GradientColorDark=   16761024
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPressMoney 
      Height          =   4110
      Left            =   195
      TabIndex        =   2
      Top             =   2145
      Width           =   10515
      _cx             =   18547
      _cy             =   7250
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPatiPressMoneySet.frx":6852
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
      Editable        =   2
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
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   0
      ScaleHeight     =   1245
      ScaleWidth      =   11010
      TabIndex        =   4
      Top             =   0
      Width           =   11010
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmPatiPressMoneySet.frx":698B
         Height          =   555
         Left            =   600
         TabIndex        =   7
         Top             =   615
         Width           =   7740
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -45
         X2              =   11640
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报警方案：每种方案包括各病区报警线及报警方式，需和 zl_PatiWarnScheme 函数配合使用"
         Height          =   180
         Left            =   600
         TabIndex        =   6
         Top             =   390
         Width           =   7290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报警方案设置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   5
         Top             =   135
         Width           =   1170
      End
   End
   Begin MSComctlLib.TabStrip tab报警 
      Height          =   4650
      Left            =   105
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1725
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   8202
      HotTracking     =   -1  'True
      TabMinWidth     =   0
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "普通病人"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   1425
      Left            =   45
      ScaleHeight     =   1425
      ScaleWidth      =   11025
      TabIndex        =   11
      Top             =   6525
      Width           =   11025
      Begin VB.CommandButton cmdWarnNew 
         Caption         =   "增加报警方案(&A)"
         Height          =   350
         Left            =   45
         TabIndex        =   15
         Top             =   60
         Width           =   1710
      End
      Begin VB.CommandButton cmdWarnDel 
         Caption         =   "删除报警方案(&D)"
         Height          =   350
         Left            =   1860
         TabIndex        =   14
         Top             =   60
         Width           =   1710
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "保存方案(&O)"
         Height          =   350
         Left            =   9465
         TabIndex        =   3
         Top             =   60
         Width           =   1395
      End
      Begin VB.Frame fraSplit 
         Height          =   90
         Left            =   -60
         TabIndex        =   13
         Top             =   465
         Width           =   11025
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   9570
         TabIndex        =   12
         Top             =   735
         Width           =   1150
      End
      Begin VB.Label lbl门诊 
         AutoSize        =   -1  'True
         Caption         =   "注意:门诊的不区分站点."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   195
         TabIndex        =   18
         Top             =   735
         Visible         =   0   'False
         Width           =   2820
      End
   End
End
Attribute VB_Name = "frmPatiPressMoneySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mblnEdit As Boolean, mblnSort As Boolean, mblnChange As Boolean
Private mrsWarn As ADODB.Recordset
Private mrs类别 As ADODB.Recordset
Private mstrDel适用病人 As String
Private mblnOK As Boolean
Private mblnNotClick As Boolean
Private mlngPreSelIdx As Long   '上次索引
Private mblnFirst As Boolean

Private Sub LoadClients()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载站点
    '编制:刘兴洪
    '日期:2011-02-11 10:10:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim str站点 As String
    
    On Error GoTo errHandle
    
    str站点 = zlDatabase.GetPara("上次选择站点", glngSys, mlngModule, "", Array(cbo站点, lbl站点), InStr(1, mstrPrivs, ";参数设置;") > 0)
    gstrSQL = "" & _
    "   Select Distinct q.编号, Q.名称 As 站点名称 " & _
     "  From 部门性质说明 B, 部门表 A ,Zlnodelist Q " & _
     "  Where B.服务对象 In (1, 2, 3) And B.工作性质 = '护理' And B.部门id = A.ID And A.站点 = Q.编号 And " & _
     "         (A.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or A.撤档时间 Is Null) " & _
     "    Order By 编号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    mblnNotClick = True
    With rsTemp
        cbo站点.Clear
        Do While Not .EOF
            cbo站点.AddItem NVL(rsTemp!编号) & "-" & NVL(rsTemp!站点名称)
            If cbo站点.ListIndex < 0 And NVL(rsTemp!编号) = gstrNodeNo Then
                cbo站点.ListIndex = cbo站点.NewIndex
            End If
            If str站点 = NVL(rsTemp!编号) Then
                cbo站点.ListIndex = cbo站点.NewIndex
            End If
            .MoveNext
        Loop
        If cbo站点.ListIndex < 0 And cbo站点.ListCount > 0 Then cbo站点.ListIndex = 0
        pic站点.ZOrder
        pic站点.Visible = cbo站点.ListCount > 0
        lbl门诊.Visible = cbo站点.ListCount > 0
    End With
    mlngPreSelIdx = cbo站点.ListIndex
    mblnNotClick = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function zlShowMe(ByVal frmMain As Form, lngModule As Long, strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:frmMain-调用的主程序
    '        lngModule -模块号
    '        strPrivs-权限串
    '出参:
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-20 09:39:27
    '问题:35386
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrPrivs = strPrivs: mlngModule = lngModule: mblnOK = False: mblnChange = False
    Me.Show 1, frmMain
    zlShowMe = mblnOK
End Function

Private Sub InitGridData()
    Dim rsTemp As ADODB.Recordset
'    gstrSQL = "" & _
'    "   Select -1 as ID,'Z' as 编码,'* 门诊 * ' as 名称 From dual Union All " & _
'    "   Select A.ID,A.编码 ,A.编码||'-'||A.名称  as 名称" & _
'    "   From  部门性质说明 b,部门表 a " & _
'    "   Where B.服务对象 in(1,2,3) And B.工作性质='护理'  " & _
'    "           And  b.部门ID=a.ID and " & Where撤档时间("A") & _
'    "   Order by 编码"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With vsPressMoney
        .Clear 1
'        If rsTemp.EOF Then
'            .ColComboList(.ColIndex("病区")) = " "
'        Else
'            .ColComboList(.ColIndex("病区")) = .BuildComboList(rsTemp, "名称", "ID")
'        End If
        .ColComboList(.ColIndex("报警方法")) = "1-累计费用|2-每日费用"
        .ColComboList(.ColIndex("病区")) = "..."
        .ColComboList(.ColIndex("报警方式1")) = "..."
        .ColComboList(.ColIndex("报警方式2")) = "..."
        .ColComboList(.ColIndex("报警方式3")) = "..."
        mblnEdit = InStr(1, mstrPrivs, ";报警方案设置;") > 0
        If mblnEdit Then .Editable = flexEDKbdMouse
        cmdWarnNew.Enabled = mblnEdit
        cmdWarnDel.Enabled = mblnEdit
        cmdOK.Visible = mblnEdit
        zl_vsGrid_Para_Restore mlngModule, vsPressMoney, Me.Caption, "报警列表", False
    End With
 End Sub
Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2011-01-20 09:31:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strCoding As String, i As Long
     
    On Error GoTo errHandle
    
   '记帐报警类别
    gstrSQL = "Select RowID as ID,编码,类别 From 收费类别 Order by 编码"
    Set mrs类别 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    lst类别.Clear
    lst类别.AddItem "所有类别"
    Do While Not mrs类别.EOF
        lst类别.AddItem mrs类别!类别
        lst类别.ItemData(lst类别.NewIndex) = Asc(mrs类别!编码)
        mrs类别.MoveNext
    Loop
    Call LoadScheme
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub LoadScheme()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载指定站点的方案
    '编制:刘兴洪
    '日期:2011-02-12 10:46:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strCoding As String, i As Long
    
    On Error GoTo errHandle
    
    '病区记帐报警线
    Set mrsWarn = New ADODB.Recordset
    mrsWarn.Fields.Append "病区ID", adBigInt, , adFldIsNullable
    mrsWarn.Fields.Append "病区码", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "病区名", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "适用病人", adVarChar, 100
    mrsWarn.Fields.Append "报警方法", adSmallInt
    mrsWarn.Fields.Append "报警值", adCurrency
    mrsWarn.Fields.Append "报警标志1", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "报警标志2", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "报警标志3", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "催款下限", adCurrency
    mrsWarn.Fields.Append "催款标准", adCurrency
    mrsWarn.CursorLocation = adUseClient
    mrsWarn.LockType = adLockOptimistic
    mrsWarn.CursorType = adOpenStatic
    mrsWarn.Open
    gstrSQL = "" & _
    "   Select a.病区ID,B.编码,b.名称 as 病区,a.适用病人,nvl(a.报警方法,1) as 报警方法, " & _
    "               a.报警值,a.报警标志1,a.报警标志2,a.报警标志3,A.催款下限,a.催款标准 " & _
    "   From 记帐报警线 a,部门表 b " & _
    "   Where a.病区ID= b.id(+)  " & IIf(cbo站点.ListCount > 0, " And b.站点=[1]", "") & _
    "   Order by Decode(a.适用病人,'普通病人',1,'医保病人',2,3),a.适用病人,B.编码 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(Split(cbo站点.Text & "-", "-")(0)))
    strCoding = ",普通病人" '至少有一个普通病人
    Do Until rsTemp.EOF
        mrsWarn.AddNew
        mrsWarn!病区ID = rsTemp!病区ID
        mrsWarn!病区码 = rsTemp!编码
        mrsWarn!病区名 = rsTemp!病区
        mrsWarn!适用病人 = rsTemp!适用病人
        mrsWarn!报警方法 = rsTemp!报警方法
        mrsWarn!报警值 = rsTemp!报警值
        mrsWarn!报警标志1 = rsTemp!报警标志1
        mrsWarn!报警标志2 = rsTemp!报警标志2
        mrsWarn!报警标志3 = rsTemp!报警标志3
        mrsWarn!催款下限 = Val(NVL(rsTemp!催款下限))
        mrsWarn!催款标准 = Val(NVL(rsTemp!催款标准))
        mrsWarn.Update
        If InStr(strCoding & ",", "," & rsTemp!适用病人 & ",") = 0 Then
            strCoding = strCoding & "," & rsTemp!适用病人
        End If
        rsTemp.MoveNext
    Loop
    strCoding = Mid(strCoding, 2)
    tab报警.Tabs.Clear
    For i = 0 To UBound(Split(strCoding, ","))
        tab报警.Tabs.Add , , Split(strCoding, ",")(i)
    Next
    tab报警.Tabs(1).Selected = True '之前不会激活Click事件,人为激活
   mblnChange = False
   

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub AfterDeleteRow()
    '删除行后
End Sub
Private Sub AfterAddRow(Row As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:行增加后
    '编制:刘兴洪
    '日期:2011-01-18 18:36:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
     With vsPressMoney
        .Cell(flexcpData, Row, 0, Row, .Cols - 1) = ""
        .Cell(flexcpText, Row, 0, Row, .Cols - 1) = ""
        .TextMatrix(Row, .ColIndex("报警方法")) = "1-累计费用"
    End With
End Sub
Private Sub BeforeDeleteRow(Row As Long, Cancel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除行之前
    '编制:刘兴洪
    '日期:2011-01-18 18:37:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
     With vsPressMoney
        If .Editable = flexEDNone Then Exit Sub
        If Val(.Cell(flexcpData, Row, .ColIndex("病区"))) <> 0 Then
            If MsgBox("你是否真的要删除病区为“" & .TextMatrix(Row, .ColIndex("病区")) & "”的方案记录吗?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
        mblnChange = True
    End With
End Sub

 
Private Sub cbo站点_Click()
    If mblnNotClick = True Then Exit Sub
    If mblnChange Then
        If mlngPreSelIdx <> cbo站点.ListIndex Then
             If MsgBox("注意:" & vbCrLf & "     你已经调整过方案,如果你改变站点,你所修改的方案信息" & vbCrLf & _
                "    将会丢失,你是否真的要改变?", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then
                cbo站点.ListIndex = mlngPreSelIdx: Exit Sub
             End If
        End If
    End If
    mlngPreSelIdx = cbo站点.ListIndex
    Call LoadScheme
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Check记帐报警 = False Then Exit Sub
    If SaveData = False Then Exit Sub
    mblnChange = False: mblnOK = True
    If pic站点.Visible Then
        MsgBox "方案保存成功!", vbInformation + vbOKOnly, gstrSysName
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    If vsPressMoney.Enabled And vsPressMoney.Visible Then vsPressMoney.SetFocus
    Call picDown_Resize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        If picList类别.Visible Then
            picList类别.Visible = False: Exit Sub
        End If
        Call cmdCancel_Click
    Case Else
    End Select
End Sub
 
Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    Call LoadClients
    Call InitGridData
    Call InitData
    mblnFirst = True
    mblnChange = False
     
End Sub
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    picDown.Left = ScaleLeft
    picDown.Top = ScaleHeight - picDown.Height
    picDown.Width = ScaleWidth
    
    With tab报警
        .Top = IIf(pic站点.Visible, pic站点.Top + pic站点.Height + 50, pic站点.Top)
        .Width = ScaleWidth - .Left - 50
        .Height = picDown.Top - .Top
        vsPressMoney.Width = ScaleWidth - vsPressMoney.Left - 120
        vsPressMoney.Height = picDown.Top - vsPressMoney.Top - 100
    End With
    picTop.Width = ScaleWidth - picTop.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = True Then
        If MsgBox("注意:" & vbCrLf & "   你已经更改过方案,是否真的要退出?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    zl_vsGrid_Para_Save mlngModule, vsPressMoney, Me.Caption, "报警列表", False
    Set mrsWarn = Nothing
    Set mrs类别 = Nothing
    Call zlDatabase.SetPara("上次选择站点", CStr(Split(cbo站点.Text & "-", "-")(0)), glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0)
    SaveWinState Me, App.ProductName
End Sub

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - 100
        If cbo站点.ListCount > 0 Then
            cmdOK.Top = cmdWarnDel.Top
            cmdOK.Left = .ScaleWidth - cmdOK.Width - 100
        Else
            cmdOK.Top = cmdCancel.Top
            cmdOK.Left = cmdCancel.Left - cmdOK.Width - 20
        End If
        fraSplit.Width = .ScaleWidth
    End With
End Sub

Private Sub picList类别_Resize()
    Err = 0: On Error Resume Next
    With picList类别
        shtCaption.Left = .ScaleLeft
        shtCaption.Width = .ScaleWidth: shtCaption.Top = .ScaleTop
        lst类别.Left = .ScaleLeft: lst类别.Width = .ScaleWidth
    End With
End Sub
 
Private Sub picTop_Resize()
    Err = 0: On Error Resume Next
    Line4.X2 = picTop.ScaleWidth + 30
End Sub

Private Sub vsPressMoney_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置相关的格式
    '编制:刘兴洪
    '日期:2011-01-18 18:32:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsPressMoney
        Select Case Col
        Case .ColIndex("病区")
            .ColComboList(Col) = "..."
'        Case .ColIndex("报警方式1"), .ColIndex("报警方式2"), .ColIndex("报警方式3")
'            .ColComboList(Col) = "..."
        Case .ColIndex("病区")
        Case .ColIndex("报警方法")
            If InStr(.TextMatrix(Row, .ColIndex("报警方法")), "每日费用") > 0 Then
                .TextMatrix(Row, .ColIndex("报警方式2")) = ""   '每日费用无报警方式2
                '为“每日费用”时判断一下金额不能为负数
                If IsNumeric(.TextMatrix(Row, .ColIndex("报警值"))) Then
                    If Val(.TextMatrix(Row, .ColIndex("报警值"))) < 0 Then
                        .TextMatrix(Row, .ColIndex("报警值")) = "0.00"
                    End If
                Else
                    .TextMatrix(Row, .ColIndex("报警值")) = "0.00"
                End If
            End If
        Case .ColIndex("报警值"), .ColIndex("催款下限"), .ColIndex("催款标准")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), "###0.00;-###0.00;;")
        End Select
    End With
End Sub
Private Sub vsPressMoney_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
        If mblnSort = True Then Exit Sub
        Call zl_VsGridRowChange(vsPressMoney, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsPressMoney_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------
    '功能:按钮选择
    '参数:
    '--------------------------------------------------------------------------
    Dim lngRow As Long
    With vsPressMoney
        Select Case Col
        Case .ColIndex("病区")
             If Select病区("") = False Then Exit Sub
            Call zlVsMoveGridCell(vsPressMoney, .ColIndex("病区"), , mblnEdit, lngRow)
            If lngRow >= 0 Then AfterAddRow lngRow
        Case .ColIndex("报警方式1"), .ColIndex("报警方式2"), .ColIndex("报警方式3")
            If Select的报警方式() = False Then Exit Sub
        End Select
    End With
    
End Sub
Private Sub vsPressMoney_ChangeEdit()
    mblnChange = True
End Sub

Private Sub vsPressMoney_DblClick()
    With vsPressMoney
      If .MouseCol <> .Cols - 1 And .MouseCol <> 1 Then Exit Sub
        If mblnEdit = False Then Exit Sub
        If .Col = 1 Then
            .TextMatrix(.Row, .ColIndex("报警方法")) = IIf(Left(.TextMatrix(.Row, .ColIndex("报警方法")), 1) = "1", "2-每日费用", "1-累计费用")
            If InStr(.TextMatrix(.Row, .ColIndex("报警方法")), "每日费用") > 0 Then
                .TextMatrix(.Row, .ColIndex("报警方式2")) = ""   '每日费用无报警方式2
                '为“每日费用”时判断一下金额不能为负数
                If IsNumeric(.TextMatrix(.Row, .ColIndex("报警值"))) Then
                    If Val(.TextMatrix(.Row, .ColIndex("报警值"))) < 0 Then
                        .TextMatrix(.Row, .ColIndex("报警值")) = "0.00"
                    End If
                Else
                    .TextMatrix(.Row, .ColIndex("报警值")) = "0.00"
                End If
            End If
        End If
        mblnChange = True
    End With
End Sub

Private Sub vsPressMoney_GotFocus()
    Call zl_VsGridGotFocus(vsPressMoney)
End Sub

Private Sub vsPressMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long

    With vsPressMoney
        If KeyCode <> vbKeyReturn And KeyCode <> vbKeyReturn _
            And (KeyCode <> Asc("*")) And KeyCode <> vbKeySpace _
            And KeyCode <> vbKeyShift Then
            If Shift = 1 And (KeyCode = 56 Or KeyCode <> Asc("*")) Then
                vsPressMoney_CellButtonClick .Row, .Col
            Else

            Select Case .Col
            Case .ColIndex("病区")  '.ColIndex("报警方式1"), .ColIndex("报警方式2"), .ColIndex("报警方式3"),
                .ColComboList(.Col) = ""
            Case Else
            End Select
            End If
        End If

        If KeyCode = vbKeyDelete Then
            blnCancel = False
            '删除行前
            Call BeforeDeleteRow(.Row, blnCancel)
            If blnCancel = True Then Exit Sub
            If .Row = .Rows - 1 And .Row = 1 Then
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, lngCol) = ""
                    .Cell(flexcpData, .Row, lngCol) = ""
                Next
            Else
                .RemoveItem .Row
            End If
            '删除行后
            Call AfterDeleteRow
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPressMoney
        If Trim(.TextMatrix(.Row, .ColIndex("病区"))) = "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        Call zlVsMoveGridCell(vsPressMoney, .ColIndex("病区"), , mblnEdit, lngRow)
        If lngRow >= 0 Then
            Call AfterAddRow(lngRow)
        End If
    End With
End Sub

Private Sub vsPressMoney_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '编辑处理
    Dim intCol As Integer, strKey As String, lngRow As Long

    If KeyCode <> vbKeyReturn Then Exit Sub

    With vsPressMoney
        Select Case Col
        Case .ColIndex("病区")
            strKey = Trim(.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
             If Select病区(strKey) = False Then
                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
                 Exit Sub
            End If
            .EditText = .TextMatrix(Row, Col)
'        Case .ColIndex("报警方式1"), .ColIndex("报警方式2"), .ColIndex("报警方式3")
'            strKey = Trim(.EditText)
'            strKey = Replace(strKey, Chr(vbKeyReturn), "")
'            strKey = Replace(strKey, Chr(10), "")
'            If strKey = "" Then Exit Sub
''            If Select报警方法(strKey) = False Then
''                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
''                Exit Sub
''            End If
'            .EditText = .TextMatrix(Row, Col)
        Case Else
        End Select
        Call zlVsMoveGridCell(vsPressMoney, .ColIndex("病区"), -1, mblnEdit, lngRow)
        If lngRow >= 0 Then AfterAddRow lngRow
    End With
End Sub

Private Sub vsPressMoney_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
    With vsPressMoney
        '切换报警方法
        If .Col = .ColIndex("报警方法") Then
            Select Case KeyAscii
                Case Asc(" ")
                    '切换计算标志
                    Select Case Left(.TextMatrix(.Row, .Col), 1)
                        Case "1"
                            .TextMatrix(.Row, .Col) = "2-每日费用"
                        Case Else
                            .TextMatrix(.Row, .Col) = "1-累计费用"
                    End Select
                    mblnChange = True
                Case vbKey1
                    .TextMatrix(.Row, .Col) = "1-累计费用"
                    mblnChange = True
                Case vbKey2
                    .TextMatrix(.Row, .Col) = "2-每日费用"
                    mblnChange = True
            End Select
            If InStr(.TextMatrix(.Row, .Col), "每日费用") > 0 Then
                .TextMatrix(.Row, .ColIndex("报警方式2")) = ""   '每日费用无报警方式2
            End If
        End If
    End With
End Sub

Private Sub vsPressMoney_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsPressMoney
        Select Case .Col
        Case .ColIndex("病区")  '.ColIndex("报警方式1"), .ColIndex("报警方式2"), .ColIndex("报警方式3"),
            VsFlxGridCheckKeyPress vsPressMoney, Row, Col, KeyAscii, m文本式
        Case .ColIndex("报警值"), .ColIndex("催款下限"), .ColIndex("催款标准")
            VsFlxGridCheckKeyPress vsPressMoney, Row, Col, KeyAscii, m金额式
        End Select
    End With
End Sub
Private Sub vsPressMoney_LeaveCell()
    If mblnSort Then Exit Sub
    zlCommFun.OpenIme False
End Sub
Private Sub vsPressMoney_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        '设置单元格的编辑长度
        With vsPressMoney
           Select Case .Col
               Case .ColIndex("病区") ' .ColIndex("报警方式1"), .ColIndex("报警方式2"), .ColIndex("报警方式3")
                   .EditMaxLength = 100
               Case .ColIndex("报警值"), .ColIndex("催款下限"), .ColIndex("催款标准")
                   .EditMaxLength = 16
           End Select
    End With
End Sub

Private Sub vsPressMoney_EnterCell()
    If mblnSort = True Then Exit Sub
    '新增或修改才存在设置
    If mblnEdit Then Exit Sub
    With vsPressMoney
        zlCommFun.OpenIme (False)
        Select Case .Col
        Case .ColIndex("病区"), .ColIndex("报警方式1"), .ColIndex("报警方式2"), .ColIndex("报警方式3")
             .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

 Private Sub vsPressMoney_LostFocus()
    zlCommFun.OpenIme False
     Call zl_VsGridLOSTFOCUS(vsPressMoney)
End Sub
Private Sub vsPressMoney_Validate(Cancel As Boolean)
        Dim lngRow As Long
        If Not mblnChange Then Exit Sub
        If zlControl.MouseInRect(cmdCancel.hWnd) Then Exit Sub
        '检查记帐报警设置
        If Not Check记帐报警 Then Cancel = True: Exit Sub
        '收集记帐报警数据
        With mrsWarn
            .Filter = "适用病人='" & tab报警.SelectedItem.Caption & "'"
            Do While Not .EOF
                .Delete
                .Update
                .MoveNext
            Loop
            .Filter = 0
        End With
        With vsPressMoney
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, .ColIndex("病区")) <> "" And .TextMatrix(lngRow, .ColIndex("报警值")) <> "" Then
                    mrsWarn.AddNew
                    mrsWarn!适用病人 = tab报警.SelectedItem.Caption
                    If Val(.Cell(flexcpData, lngRow, .ColIndex("病区"))) <> 0 Then
                        mrsWarn!病区ID = Val(.Cell(flexcpData, lngRow, .ColIndex("病区")))
                        If mrsWarn!病区ID <= 0 Then
                            mrsWarn!病区ID = Null
                            mrsWarn!病区码 = Null
                            mrsWarn!病区名 = Trim(.TextMatrix(lngRow, .ColIndex("病区")))
                        Else
                            mrsWarn!病区码 = Split(.TextMatrix(lngRow, .ColIndex("病区")), "-")(0)
                            mrsWarn!病区名 = Split(.TextMatrix(lngRow, .ColIndex("病区")), "-")(1)
                        End If
                    End If
                    mrsWarn!报警方法 = CInt(Left(.TextMatrix(lngRow, .ColIndex("报警方法")), 1))
                    mrsWarn!报警值 = CCur(.TextMatrix(lngRow, .ColIndex("报警值")))

                    mrsWarn!报警标志1 = Get类别编码串(.TextMatrix(lngRow, .ColIndex("报警方式1")))
                    mrsWarn!报警标志2 = Get类别编码串(.TextMatrix(lngRow, .ColIndex("报警方式2")))
                    mrsWarn!报警标志3 = Get类别编码串(.TextMatrix(lngRow, .ColIndex("报警方式3")))
                    mrsWarn!催款下限 = FormatEx(Val(.TextMatrix(lngRow, .ColIndex("催款下限"))), 2)
                    mrsWarn!催款标准 = FormatEx(Val(.TextMatrix(lngRow, .ColIndex("催款标准"))), 2)
                    mrsWarn.Update
                End If
            Next
        End With
End Sub

Private Sub vsPressMoney_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer
    '数据验证
    With vsPressMoney
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
            Case .ColIndex("报警值"), .ColIndex("催款下限"), .ColIndex("催款标准")
                If zlDblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    .EditText = Format(Val(strKey), "###0.00;-###0.00;;")
                End If
        End Select
    End With
End Sub
Private Sub vsPressMoney_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long, arrSplit As Variant
    With vsPressMoney
        If mblnEdit = False Then Cancel = True: Exit Sub
        Select Case Col
        Case .ColIndex("病区")
        Case .ColIndex("报警方式1"), .ColIndex("报警方式3")
        Case .ColIndex("报警方式2")
            '每日费用不能编辑报警方式2
            If InStr(Trim(.TextMatrix(Row, .ColIndex("报警方法"))), "每日费用") > 0 Then Cancel = True: Exit Sub
        Case .ColIndex("报警值"), .ColIndex("催款下限"), .ColIndex("催款标准")
        Case Else: Cancel = True
        End Select
    End With
End Sub
Private Sub tab报警_Click()
    Dim lngRow As Long
    mrsWarn.Filter = "适用病人='" & tab报警.SelectedItem.Caption & "'"
    With vsPressMoney
        If mrsWarn.RecordCount = 0 Then
            .Clear 1
            .Rows = 2: .Row = 1: .Col = 1
            .RowData(1) = ""
            .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
            .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
        Else
            .Clear 1
            .Rows = mrsWarn.RecordCount + 1: .Row = 1: .Col = 1
            lngRow = 1
            Do Until mrsWarn.EOF
                .RowData(lngRow) = NVL(mrsWarn!病区ID, 0)
                .TextMatrix(lngRow, .ColIndex("病区")) = IIf(IsNull(mrsWarn!病区ID), "*门诊*", mrsWarn!病区码 & "-" & mrsWarn!病区名)
                .Cell(flexcpData, lngRow, .ColIndex("病区")) = IIf(IsNull(mrsWarn!病区ID), -1, NVL(mrsWarn!病区ID, 0))
                .TextMatrix(lngRow, .ColIndex("报警方法")) = IIf(mrsWarn!报警方法 = 1, "1-累计费用", "2-每日费用")
                .TextMatrix(lngRow, .ColIndex("报警值")) = Format(mrsWarn!报警值, "###0.00;-###0.00;;")
                .TextMatrix(lngRow, .ColIndex("报警方式1")) = Get类别名称串(NVL(mrsWarn!报警标志1), mrs类别)
                .TextMatrix(lngRow, .ColIndex("报警方式2")) = Get类别名称串(NVL(mrsWarn!报警标志2), mrs类别)
                .TextMatrix(lngRow, .ColIndex("报警方式3")) = Get类别名称串(NVL(mrsWarn!报警标志3), mrs类别)
                .TextMatrix(lngRow, .ColIndex("催款下限")) = Format(mrsWarn!催款下限, "###0.00;-###0.00;0.00;0.00")
                .TextMatrix(lngRow, .ColIndex("催款标准")) = Format(mrsWarn!催款标准, "###0.00;-###0.00;0.00;0.00")
                lngRow = lngRow + 1
                mrsWarn.MoveNext
            Loop
          If .Enabled And .Visible Then .SetFocus
        End If
    End With
End Sub
Private Function Check记帐报警() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查记帐报警是否正确
    '返回:正确,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-18 18:58:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngTemp As Long
    Dim lngCol1 As Long, lngCol2 As Long
    Dim arr类别() As String

    With vsPressMoney
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, .ColIndex("病区"))) <> "" Then
                For lngTemp = lngRow + 1 To .Rows - 1
                    If .TextMatrix(lngRow, .ColIndex("病区")) = .TextMatrix(lngTemp, .ColIndex("病区")) And .TextMatrix(lngTemp, 2) <> "" Then
                        MsgBox "病区“" & .TextMatrix(lngTemp, .ColIndex("病区")) & "”出现多次。", vbExclamation, gstrSysName
                        .Row = lngTemp: .Col = .ColIndex("病区"): .SetFocus: Exit Function
                    End If
                Next
                If Val(.TextMatrix(lngRow, .ColIndex("催款下限"))) > 999999999 Or Val(.TextMatrix(lngRow, .ColIndex("催款下限"))) < 0 Then
                    MsgBox "病区“" & .TextMatrix(lngRow, .ColIndex("病区")) & "”中的催款下限设置有误(应该在0~999999999)!", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = .ColIndex("催款下限"): .SetFocus: Exit Function
                End If
                If Val(.TextMatrix(lngRow, .ColIndex("催款标准"))) > 999999999 Or Val(.TextMatrix(lngRow, .ColIndex("催款标准"))) < 0 Then
                    MsgBox "病区“" & .TextMatrix(lngRow, .ColIndex("病区")) & "”中的催款标准有误(应该在0~999999999)!", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = .ColIndex("催款标准"): .SetFocus: Exit Function
                End If
            End If
        Next

        '检查同一病区不同报警方式的类别是否一个都没有设置或重复
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("病区")) <> "" And .TextMatrix(lngRow, .ColIndex("报警值")) <> "" Then
                If Trim(.TextMatrix(lngRow, .ColIndex("报警方式1"))) = "" And Trim(.TextMatrix(lngRow, .ColIndex("报警方式2"))) = "" And Trim(.TextMatrix(lngRow, .ColIndex("报警方式3"))) = "" Then
                    MsgBox "病区“" & .TextMatrix(lngRow, .ColIndex("病区")) & "”未设置要报警的收费类别。", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = .ColIndex("报警方式1"): .SetFocus: Exit Function
                End If

                If (.TextMatrix(lngRow, .ColIndex("报警方式1")) = "所有类别" And (Trim(.TextMatrix(lngRow, .ColIndex("报警方式2"))) <> "" Or Trim(.TextMatrix(lngRow, .ColIndex("报警方式3"))) <> "")) _
                    Or (.TextMatrix(lngRow, .ColIndex("报警方式2")) = "所有类别" And (Trim(.TextMatrix(lngRow, .ColIndex("报警方式1"))) <> "" Or Trim(.TextMatrix(lngRow, .ColIndex("报警方式3"))) <> "")) _
                    Or (.TextMatrix(lngRow, .ColIndex("报警方式3")) = "所有类别" And (Trim(.TextMatrix(lngRow, .ColIndex("报警方式2"))) <> "" Or Trim(.TextMatrix(lngRow, .ColIndex("报警方式1"))) <> "")) Then

                    MsgBox "病区“" & .TextMatrix(lngRow, .ColIndex("病区")) & "”不同的报警方式包含相同的收费类别。", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = .ColIndex("报警方式1"): .SetFocus: Exit Function
                End If
                If .TextMatrix(lngRow, .ColIndex("报警方式1")) <> "所有类别" And Trim(.TextMatrix(lngRow, .ColIndex("报警方式2"))) <> "所有类别" And Trim(.TextMatrix(lngRow, .ColIndex("报警方式3"))) <> "所有类别" Then
                    For lngCol1 = .ColIndex("报警方式1") To .ColIndex("报警方式3")
                        If Trim(.TextMatrix(lngRow, lngCol1)) <> "" Then
                            For lngCol2 = .ColIndex("报警方式1") To .ColIndex("报警方式3")
                                If lngCol1 <> lngCol2 Then
                                    arr类别 = Split(.TextMatrix(lngRow, lngCol1), ",")
                                    For lngTemp = 0 To UBound(arr类别)
                                        If InStr("," & .TextMatrix(lngRow, lngCol2) & ",", "," & arr类别(lngTemp) & ",") > 0 Then
                                            MsgBox "病区“" & .TextMatrix(lngRow, .ColIndex("病区")) & "”不同的报警方式包含相同的收费类别。", vbExclamation, gstrSysName
                                            .Row = lngRow: .Col = .ColIndex("报警方式1"): .SetFocus: Exit Function
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        Next
    End With

    Check记帐报警 = True
End Function


Private Function Get类别编码串(str类别 As String) As String
'功能：根据类似"检查,治疗"的串返回类似"CDEFG"的串
    Dim i As Integer, j As Integer
    Dim arr类别() As String, strTmp As String

    If Trim(str类别) = "" Then Exit Function
    If str类别 = "所有类别" Then
        Get类别编码串 = "-"
    Else
        arr类别 = Split(str类别, ",")
        For i = 0 To UBound(arr类别)
            For j = 1 To lst类别.ListCount - 1
                If lst类别.List(j) = arr类别(i) Then
                    strTmp = strTmp & Chr(lst类别.ItemData(j))
                    Exit For
                End If
            Next
        Next
        Get类别编码串 = strTmp
    End If
End Function

Private Sub cmdWarnNew_Click()
    Dim strName As String, strCopy As String
    Dim strSchemes As String, i As Integer
    Dim rsCopy As ADODB.Recordset
    
    For i = 1 To tab报警.Tabs.Count
        strSchemes = strSchemes & "," & tab报警.Tabs(i).Caption
    Next
    
    strName = frmWarnEdit.ShowMe(Me, Mid(strSchemes, 2), strCopy)
    If strName = "" Then Exit Sub
    
    '复制内容
    Set rsCopy = mrsWarn.Clone
    rsCopy.Filter = "适用病人='" & strCopy & "'"
    Do While Not rsCopy.EOF
        mrsWarn.AddNew
        mrsWarn!适用病人 = strName
        mrsWarn!病区ID = rsCopy!病区ID
        mrsWarn!病区码 = rsCopy!病区码
        mrsWarn!病区名 = rsCopy!病区名
        mrsWarn!报警方法 = rsCopy!报警方法
        mrsWarn!报警值 = rsCopy!报警值
        mrsWarn!报警标志1 = rsCopy!报警标志1
        mrsWarn!报警标志2 = rsCopy!报警标志2
        mrsWarn!报警标志3 = rsCopy!报警标志3
        mrsWarn!催款下限 = rsCopy!催款下限
        mrsWarn!催款标准 = rsCopy!催款标准
        mrsWarn.Update
        rsCopy.MoveNext
    Loop
    
    tab报警.Tabs.Add , , strName
    tab报警.Tabs(tab报警.Tabs.Count).Selected = True
    
    mblnChange = True
End Sub
Private Sub cmdWarnDel_Click()
    If tab报警.SelectedItem.Caption = "普通病人" Then
        MsgBox """" & tab报警.SelectedItem.Caption & """报警方案不允许删除。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("确实要删除""" & tab报警.SelectedItem.Caption & """报警方案吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    With mrsWarn
        .Filter = "适用病人='" & tab报警.SelectedItem.Caption & "'"
        
        '记录删除的适用病人类型
        If InStr(1, mstrDel适用病人, tab报警.SelectedItem.Caption) = 0 Then
            mstrDel适用病人 = IIf(mstrDel适用病人 = "", "", mstrDel适用病人 & ";") & tab报警.SelectedItem.Caption
        End If
        
        Do While Not .EOF
            .Delete
            .Update
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    tab报警.Tabs.Remove tab报警.SelectedItem.Index
    tab报警.Tabs(1).Selected = True
    
    mblnChange = True
End Sub
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据保存
    '编制:刘兴洪
    '日期:2011-01-20 09:35:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str适用病人 As String, strTemp As String, i As Long
    Dim str站点 As String
    
    On Error GoTo errHandle
    If cbo站点.ListCount = 0 Or cbo站点.ListIndex < 0 Then
        str站点 = "NULL"
    Else
        str站点 = "'" & Split(cbo站点.Text & "-", "-")(0) & "'"
    End If
    '按适用病人分批保存
    mrsWarn.Filter = 0
    For i = 1 To tab报警.Tabs.Count
        strTemp = ""
        str适用病人 = tab报警.Tabs.Item(i).Caption
        mrsWarn.Filter = "适用病人='" & str适用病人 & "'"
        Do While Not mrsWarn.EOF
            strTemp = strTemp & NVL(mrsWarn!病区ID) & "," & mrsWarn!报警方法 & "," & _
            mrsWarn!报警值 & "," & NVL(mrsWarn!报警标志1) & "," & NVL(mrsWarn!报警标志2) & "," & NVL(mrsWarn!报警标志3) & "," & NVL(mrsWarn!催款下限) & "," & NVL(mrsWarn!催款标准) & ","
            mrsWarn.MoveNext
        Loop
        strTemp = str适用病人 & "|" & strTemp
        ' Zl_记帐报警线_Modify
        gstrSQL = "zl_记帐报警线_Modify("
        '  报警线_In In Varchar2,
        gstrSQL = gstrSQL & "'" & strTemp & "',"
        '  站点_In Varchar2:=Null
        gstrSQL = gstrSQL & "" & str站点 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function Get类别名称串(str类别 As String, rs类别 As ADODB.Recordset) As String
    '功能：将类似"CDEFG"的类别转换为类似"检查,检验..."串
    Dim i As Integer, strTmp As String
    If str类别 = "" Then
        Get类别名称串 = " " '为了能按回车新增行
        Exit Function
    End If
    If str类别 = "-" Then
        Get类别名称串 = "所有类别"
        Exit Function
    End If
    For i = 1 To Len(str类别)
        rs类别.Filter = "编码='" & Mid(str类别, i, 1) & "'"
        If Not rs类别.EOF Then strTmp = strTmp & "," & rs类别!类别
    Next
    Get类别名称串 = Mid(strTmp, 2)
End Function
Private Function Select病区(ByVal strSearch As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病区选择器
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-20 10:39:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strFind As String
    Dim sngX As Single, sngY As Single, bytStyle As Byte
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    strTittle = "病区选择": bytStyle = 0
    strKey = gstrLike & strSearch & "%"
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.编码 like upper([1]) or a.简码 like upper([1]) or a.名称 like [1] )"
        If IsNumeric(strSearch) Then                         '如果是数字,则只取编码
            strFind = " And (A.编码 Like Upper([1]))"
        ElseIf zlCommFun.IsCharAlpha(strSearch) Then           '01,11.输入全是字母时只匹配简码
            '0-拼音码,1-五笔码,2-两者
            '.int简码方式 = Val(zlDatabase.GetPara("简码方式" ))
            strFind = " And  (a.简码 Like Upper([1]))"
        ElseIf zlCommFun.IsCharChinese(strSearch) Then  '全汉字
            strFind = " And a.名称 Like [1] "
        End If
    End If
    If strSearch = "" Then
        gstrSQL = "" & _
            "Select * " & _
            "  From (With M As (Select Distinct A.ID, -10 * Ascii(A.站点) As 上级id, A.编码, A.名称, A.简码, A.站点, 1 As 末级,Q.名称 as 站点名称" & _
            "                   From 部门性质说明 B, 部门表 A,Zlnodelist Q " & _
            "                   Where B.服务对象 In (1, 2, 3) And B.工作性质 = '护理' " & _
                                    IIf(cbo站点.ListCount > 0, " And A.站点=[2] ", "") & " And B.部门id = A.ID And a.站点=Q.编号(+) And " & _
            "                         (A.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or A.撤档时间 Is Null)) " & _
            "         Select -10 * Ascii(A.编号) As ID, -1 * Null As 上级id, To_Char(编号) As 编码, 名称, '' As 简码, " & _
            "                编号 as 站点 , 0 As 末级,名称 As 站点名称  " & _
            "         From Zlnodelist A " & _
            "         Where Exists (Select 1 From M Where M.站点 = A.编号) " & _
            "         Union All " & _
            "         Select -1 As ID, -1 * Null As 上级id, '-' As 编码, '* 门诊 * ' As 名称, 'MZ' As 简码, '' As 站点, 1 As 末级,'' as 站点名称 " & _
            "         From Dual " & _
            "         Union All " & _
            "         Select * From M) " & _
            "   "
            bytStyle = 2
    Else
        gstrSQL = "" & _
          "   Select * From ( " & _
          "   Select -1 as ID,'Z' as RID,'-' as 编码,'* 门诊 * ' as 名称, 'MZ' as 简码,'' as 站点名称 From dual Union All " & _
          "   Select distinct A.ID,A.编码 as RID,A.编码 ,A.名称,A.简码,M.名称 as 站点名称 " & _
          "   From  部门性质说明 b,部门表 a,Zlnodelist M  " & _
          "   Where B.服务对象 in(1,2,3) And B.工作性质='护理'  " & IIf(cbo站点.ListCount > 0, " And A.站点=[2] ", "") & _
          "           And A.站点=M.编号(+) And  b.部门ID=a.ID and " & Where撤档时间("A") & _
          "     ) A " & IIf(strSearch <> "", " Where 1=1 " & strFind, "") & _
          "   Order by RID"
    End If
    Call CalcPosition(sngX, sngY, vsPressMoney)
    lngH = vsPressMoney.CellHeight
    sngY = sngY - lngH
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, strTittle, IIf(strSearch = "", True, False), "", "", False, IIf(strSearch = "", True, False), True, sngX, sngY, lngH, blnCancel, False, False, strKey, CStr(Split(cbo站点.Text & "-", "-")(0)))
    If blnCancel = True Then
        vsPressMoney.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "没有满足条件的病区,请检查!"
        If vsPressMoney.Enabled Then vsPressMoney.SetFocus
        Exit Function
    End If
    '检查是否有重复的病区
    With vsPressMoney
        For i = 1 To .Rows - 1
            If i <> .Row Then
                If .Cell(flexcpData, i, .Col) = Val(rsTemp!ID) Then
                    MsgBox "在第: " & i & "行已经存在相同的病区,不能再选择该病区!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                    If vsPressMoney.Enabled Then vsPressMoney.SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
    vsPressMoney.SetFocus
    With vsPressMoney
        .TextMatrix(.Row, .Col) = IIf(NVL(rsTemp!编码) = "-", "", NVL(rsTemp!编码) & "-") & NVL(rsTemp!名称)
        .Cell(flexcpData, .Row, .Col) = Val(rsTemp!ID)
    End With
    'zlVsMoveGridCell vsPressMoney, vsPressMoney.ColIndex("病区"), mblnEdit, i
    Select病区 = True
End Function
Private Function Select的报警方式() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:报警方式选择器
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-20 10:39:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsPressMoney
        Call Set类别选择(.TextMatrix(.Row, .Col))
        picList类别.Left = .Left + .CellLeft
        If .Top + .CellTop + .CellHeight + picList类别.Height <= .Container.Height Then
            picList类别.Top = .Top + .CellTop + .CellHeight
        Else
            picList类别.Top = .Top + .CellTop - .Height - 30
        End If
        picList类别.Width = IIf(.CellWidth < 1200, 1200, .CellWidth + 30)
        picList类别.ZOrder
        picList类别.Visible = True
        lst类别.SetFocus
    End With
    Select的报警方式 = True
End Function

Private Sub Set类别选择(str类别 As String)
'功能：根据类似"检查,治疗..."的串设置列表的选择情况
    Dim i As Integer, j As Integer
    Dim arr类别() As String
    
    For i = 0 To lst类别.ListCount - 1
        lst类别.Selected(i) = False
    Next
    
    If Trim(str类别) = "" Then
        Exit Sub
    ElseIf str类别 = "所有类别" Then
        For i = 0 To lst类别.ListCount - 1
            lst类别.Selected(i) = (i = 0)
        Next
    Else
        lst类别.Selected(0) = False
        arr类别 = Split(str类别, ",")
        For i = 0 To UBound(arr类别)
            For j = 1 To lst类别.ListCount - 1
                If lst类别.List(j) = arr类别(i) Then
                    lst类别.Selected(j) = True: Exit For
                End If
            Next
        Next
    End If
    
    For i = 0 To lst类别.ListCount - 1
        If lst类别.Selected(i) Then
            lst类别.TopIndex = i: Exit For
        End If
    Next
End Sub
Private Sub lst类别_ItemCheck(Item As Integer)
    Dim i As Integer
    If Item = 0 And lst类别.Selected(Item) Then
        For i = 1 To lst类别.ListCount - 1
            lst类别.Selected(i) = False
        Next
    ElseIf Item > 0 And lst类别.Selected(Item) Then
        lst类别.Selected(0) = False
    End If
End Sub

Private Sub lst类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call lst类别_Validate(False)
        Call Form_KeyDown(vbKeyEscape, 0)
    End If
End Sub

Private Sub lst类别_LostFocus()
    picList类别.Visible = False
End Sub
Private Sub lst类别_Validate(Cancel As Boolean)
    Dim i As Integer
    With vsPressMoney
        .TextMatrix(.Row, .Col) = Get类别选择
        If .TextMatrix(.Row, .Col) = "所有类别" Then
            For i = .ColIndex("报警方式1") To .ColIndex("报警方式3")
                If i <> .Col Then .TextMatrix(.Row, i) = " "
            Next
        End If
    End With
    mblnChange = True
End Sub
Private Function Get类别选择() As String
'功能：根据类别选择框选择的情况返回类似"检查,治疗..."的串
    Dim i As Integer, strTmp As String
    
    If lst类别.Selected(0) Then
        Get类别选择 = "所有类别"
    Else
        For i = 1 To lst类别.ListCount - 1
            If lst类别.Selected(i) Then
                strTmp = strTmp & "," & lst类别.List(i)
            End If
        Next
        Get类别选择 = Mid(strTmp, 2)
        If Get类别选择 = "" Then Get类别选择 = " " '为了能回车新增行
    End If
End Function
