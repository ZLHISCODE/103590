VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
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
Private mlng病人ID As Long
Private mstrPrivs As String
Private mcllFilter As Collection
Private mrsPati As ADODB.Recordset
Private mlngCurRow As Long, mlngTopRow As Long
Private mlngGo As Long, mblnDown As Boolean, mblnGo As Boolean
Private mfrmFilter As frmPatiFilter
Private mfrmFind As frmPatiFind
Private mstrUnitIDs As String '操作员所在病区或科室所属病区
  
Private mblnShowCard As Boolean '卡号是否加密显示

Private mcnOracle As ADODB.Connection
Private mobjDataBase As clsDataBase
Private mobjOneDataObject As clsOneCardDataObject
Private mblnOk As Boolean

Public Function zlShowCard(ByVal cnOracle As ADODB.Connection, frmMain As Object, ByVal strPrivs As String, Optional lng病人ID_Out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示病人选择器
    '入参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-05 16:56:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    mlng病人ID = 0: mstrPrivs = strPrivs
    If zlGetOneDataBase(cnOracle, mobjDataBase) = False Then Exit Function
    If zlGetOneCardDataObject(cnOracle, mobjOneDataObject) = False Then Exit Function
    
    Set mcllFilter = New Collection
    If mobjOneDataObject.zlGetCardFromCardTypeID("就诊卡", False, objCard) Then
        mblnShowCard = objCard.卡号密文规则 = ""
    End If
    
    If frmMain Is Nothing Then
         Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    lng病人ID_Out = mlng病人ID
    zlShowCard = mblnOk
End Function


Private Sub cmdCancel_Click()
    mlng病人ID = 0
    mblnOk = False
    If gobjComLib Is Nothing Then zlInitCommLib
    If gobjComLib Is Nothing Then Exit Sub
    gobjComLib.saveWinState Me, App.ProductName
    Hide
End Sub

Private Sub cmdFilter_Click()
    If mfrmFilter.zlShowCard(Me, Val(mshPati.Tag), mcllFilter, mcnOracle) = False Then Exit Sub
    Call ShowPatis(mcllFilter)
End Sub

Private Sub cmdFind_Click()
    Dim blnOK As Boolean
    blnOK = gblnOk
    mfrmFind.mbytType = Val(mshPati.Tag)
    mfrmFind.Show 1, Me
    If gblnOk Then Call SeekPati(mfrmFind.optHead)
    gblnOk = blnOK
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
    mblnOk = True
    
    If gobjComLib Is Nothing Then zlInitCommLib
    If Not gobjComLib Is Nothing Then
        gobjComLib.saveWinState Me, App.ProductName
    End If
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
    If gobjComLib Is Nothing Then zlInitCommLib
    If Not gobjComLib Is Nothing Then
        gobjComLib.RestoreWinState Me, App.ProductName
     End If
    
    If glngSys Like "8??" Then
        Caption = "客户选择"
        tvw_s.Visible = False
        pic.Visible = False
    End If
    Set mfrmFilter = New frmPatiFilter
    Set mfrmFind = New frmPatiFind
        
    mstrUnitIDs = mobjOneDataObject.zlGetUserUnits
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
    Err = 0: On Error Resume Next
    If Not mfrmFind Is Nothing Then Unload mfrmFind: Set mfrmFind = Nothing
    If Not mfrmFilter Is Nothing Then Unload mfrmFilter: Set mfrmFilter = Nothing
    If Not mrsPati Is Nothing Then Set mrsPati = Nothing
    If Not mobjDataBase Is Nothing Then Set mobjDataBase = Nothing
    If Not mobjOneDataObject Is Nothing Then Set mobjOneDataObject = Nothing
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
    If gobjComLib Is Nothing Then zlInitCommLib
    If Not gobjComLib Is Nothing Then
        gobjComLib.saveWinState Me, App.ProductName
    End If
    
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
    
    On Error GoTo ErrH
    
    strPreKey = ""
    If Not tvw_s.SelectedItem Is Nothing Then strPreKey = tvw_s.SelectedItem.Key
    
    tvw_s.Nodes.Clear
    Set objNode = tvw_s.Nodes.Add(, , "All", "所有病人", 1)
    objNode.Expanded = True
    Set objNode = tvw_s.Nodes.Add("All", 4, "In", "在院病人", 1)
    Set objNode = tvw_s.Nodes.Add("All", 4, "Out", "出院病人", 1)
    Set objNode = tvw_s.Nodes.Add("All", 4, "Clinic", "门诊病人", 1)
    Set objNode = tvw_s.Nodes.Add("All", 4, "Temp", "留观病人", 1)
    
    objNode.Expanded = True
    If objNode.Key = strPreKey Then objNode.Selected = True
    Set rsTmp = mobjOneDataObject.zlGetUnitRecordFromDepdIDs(InStr(mstrPrivs, "所有病区") = 0, "1,2,3", "护理")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            Set objNode = tvw_s.Nodes.Add("In", 4, "D" & rsTmp!id, "[" & rsTmp!编码 & "]" & rsTmp!名称, 1)
            
            If rsTmp!id = UserInfo.部门ID Then objNode.Selected = True
            If objNode.Key = strPreKey Then objNode.Selected = True
            objNode.Expanded = True
            
            rsTmp.MoveNext
        Next
    End If
    If tvw_s.SelectedItem Is Nothing Then tvw_s.Nodes("In").Selected = True

    
    InitUnits = True
    
    Call tvw_s_NodeClick(tvw_s.SelectedItem)
    Exit Function
ErrH:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
    Call mobjDataBase.SaveErrLog
End Function

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If tvw_s.Tag = Node.Key Then Exit Sub
    tvw_s.Tag = Node.Key
    If gobjComLib Is Nothing Then zlInitCommLib
    If Not gobjComLib Is Nothing Then gobjComLib.SaveFlexState mshPati, App.ProductName & "\" & Me.Name
    Call ShowPatis(Nothing, , True)      '切换病人类型时,条件清空,使用缺省条件
End Sub

  



Private Sub ShowPatis(ByVal cllFilter As Collection, Optional blnSort As Boolean, Optional blnSet As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前菜单浏览要求(自动生成条件),读取病人信息
    '入参:cllfilter-过滤条件
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-30 16:49:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str病区IDs As String, strInfo As String, curDate As Date
    Dim blnLimitUnit As Boolean, blnFirst As Boolean
    Dim blnPatiQuery As Boolean    '是否按病人信息查询
    
    On Error GoTo ErrH
    If Not blnSort Then
        blnLimitUnit = InStr(mstrPrivs, ";所有病区;") = 0
        If cllFilter.count = 0 Then
            blnFirst = True
            If InStr(1, ",All,Clinic,Temp,", "," & tvw_s.SelectedItem.Key & ",") > 0 Then
                 curDate = mobjDataBase.Currentdate
                 cllFilter.Add Array("登记时间", Format(curDate, "yyyy-mm-dd 00:00:00"), Format(curDate, "yyyy-mm-dd 23:59:59"), "登记时间")
            ElseIf tvw_s.SelectedItem.Key = "Out" Then
                 curDate = mobjDataBase.Currentdate
                 cllFilter.Add Array("出院日期", Format(curDate, "yyyy-mm-dd 00:00:00"), Format(curDate, "yyyy-mm-dd 23:59:59"), "出院日期")
            ElseIf tvw_s.SelectedItem.Key = "In" Then
                 curDate = mobjDataBase.Currentdate
                 cllFilter.Add Array("入院日期", Format(curDate, "yyyy-mm-dd 00:00:00"), Format(curDate, "yyyy-mm-dd 23:59:59"), "入院日期")
            End If
        End If
        
         
        If tvw_s.SelectedItem.Key = "All" Then '所有病人
            
            strInfo = "正在读取所有病人清单,请稍候 ..."
            gobjCommFun.ShowFlash strInfo
            If Val(mshPati.Tag) <> 0 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 0
            Call GetHospitalizationPatientData(IIf(blnLimitUnit, mstrUnitIDs, ""), cllFilter, mrsPati)
             
        ElseIf tvw_s.SelectedItem.Key = "In" Or Mid(tvw_s.SelectedItem.Key, 1, 1) = "D" Then  '在院病人
            strInfo = "正在读取在院病人清单,请稍候 ..."
            gobjCommFun.ShowFlash strInfo
            If Val(mshPati.Tag) <> 1 Then
               Unload mfrmFind
               Unload mfrmFilter
            End If
            mshPati.Tag = 1
                    
            str病区IDs = IIf(blnLimitUnit, mstrUnitIDs, "")
            If Mid(tvw_s.SelectedItem.Key, 1, 1) = "D" Then
                str病区IDs = Mid(tvw_s.SelectedItem.Key, 2)
            End If
            Call GetHospitalizationPatientData(str病区IDs, cllFilter, mrsPati)
        ElseIf tvw_s.SelectedItem.Key = "Out" Then '出院病人
            strInfo = "正在读取出院病人清单,请稍候 ..."
            gobjCommFun.ShowFlash strInfo
            If Val(mshPati.Tag) <> 2 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 2
            Call GetLeavePatientData(IIf(blnLimitUnit, mstrUnitIDs, ""), cllFilter, mrsPati)
        ElseIf tvw_s.SelectedItem.Key = "Clinic" Then '门诊病人
            strInfo = "正在读取门诊病人清单,请稍候 ..."
            gobjCommFun.ShowFlash strInfo
            tvw_s.Tag = tvw_s.SelectedItem.Key
            sta.SimpleText = strInfo
            Screen.MousePointer = 11
            DoEvents
            Me.Refresh
            If Val(mshPati.Tag) <> 3 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 3
            Call GetOutPatientData(cllFilter, mrsPati) '获取门诊病人信息'
        ElseIf tvw_s.SelectedItem.Key = "Temp" Then
            '门诊留观和住院留观病人
            strInfo = "正在读取留观病人清单,请稍候 ..."
            gobjCommFun.ShowFlash strInfo
            If Val(mshPati.Tag) <> 4 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 4
            Call GetObservationPatiData(cllFilter, mrsPati) '获取留观病人数据

        End If
        
        tvw_s.Tag = tvw_s.SelectedItem.Key
        sta.SimpleText = strInfo
        Screen.MousePointer = 11
        DoEvents
        Me.Refresh
        gobjCommFun.StopFlash
    End If
    
    mshPati.Clear
    mshPati.Rows = 2
    If mrsPati.EOF Then
        Call SetHeader(blnSet)
        sta.SimpleText = IIf(blnFirst, "当天", "") & "没有找到符合条件的病人,请点击[筛选],选择查询条件."
    Else
        Set mshPati.DataSource = mrsPati
        Call SetHeader(blnSet)
        sta.SimpleText = IIf(blnFirst, "当天", "") & "共找到 " & mrsPati.RecordCount & " 位符合条件的病人."
    End If
    
    Screen.MousePointer = 0
    Me.Refresh
    Exit Sub
ErrH:
    Screen.MousePointer = 0
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
   Call gobjComLib.SaveErrLog
     gobjCommFun.StopFlash
End Sub


Private Function GetAllPatientData(ByVal str病区IDs As String, ByVal cllFilter As Collection, ByRef rsPatiInfo_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取所有病人信息
    '入参:cllFilter-过滤条件
    '     str病区IDs-指定病区
    '出参:rsPatiInfo_Out-返回病人数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 15:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiCons As Collection, cllPageCons As Collection
    Dim cllPatiData As Collection, cllPageData As Collection
    Dim cllTemp As Collection, cllTemp1 As Collection, cllPageTemp As Collection
    Dim blnPatiQuery As Boolean, str病人Ids As String
    Dim i As Long, lng病人ID As Long, J As Long
    
    
    
    On Error GoTo errHandle


'    strSQL = "" & _
'    "   Select A.病人ID,A.门诊号,A.住院号," & strCard & "A.姓名,A.性别,A.年龄,A.费别 as 门诊费别," & _
'    "           B.名称 as 病区,C.名称 as 科室,A.当前床号 as 床号,To_Char(A.入院时间,'YYYY-MM-DD') as 入院时间," & _
'    "           To_Char(A.出院时间,'YYYY-MM-DD') as 出院时间,A.住院次数,To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期," & _
'    "           A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份,A.身份证号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间," & _
'    "           Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型" & _
'    " From 病案主页 P,病人信息 A,部门表 B,部门表 C" & _
'    " Where A.当前病区ID=B.ID(+) And A.当前科室ID=C.ID(+) And A.病人ID=P.病人ID(+) And A.主页ID=P.主页ID(+) " & strIF & _
'    " Order by A.登记时间 Desc"
    
    Set rsPatiInfo_Out = New ADODB.Recordset
    With rsPatiInfo_Out
        If .State = 1 Then .Close
        .fields.Append "病人ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "门诊号", adLongVarChar, 20, adFldIsNullable
        .fields.Append "住院号", adLongVarChar, 20, adFldIsNullable
        .fields.Append "就诊卡号", adLongVarChar, 100, adFldIsNullable
    
        .fields.Append "姓名", adLongVarChar, 100, adFldIsNullable
        .fields.Append "性别", adLongVarChar, 20, adFldIsNullable
        .fields.Append "年龄", adLongVarChar, 20, adFldIsNullable
        .fields.Append "门诊费别", adLongVarChar, 50, adFldIsNullable
        
        .fields.Append "病区", adLongVarChar, 100, adFldIsNullable
        .fields.Append "科室", adLongVarChar, 100, adFldIsNullable
        .fields.Append "床号", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "入院时间", adLongVarChar, 30, adFldIsNullable
        .fields.Append "出院时间", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "住院次数", adVarNumeric, 18, adFldIsNullable
        .fields.Append "出生日期", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "国籍", adLongVarChar, 100, adFldIsNullable
        .fields.Append "民族", adLongVarChar, 100, adFldIsNullable
        .fields.Append "区域", adLongVarChar, 100, adFldIsNullable
        .fields.Append "学历", adLongVarChar, 100, adFldIsNullable
        .fields.Append "职业", adLongVarChar, 100, adFldIsNullable
        .fields.Append "身份", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "身份证号", adLongVarChar, 30, adFldIsNullable
        .fields.Append "家庭地址", adLongVarChar, 200, adFldIsNullable
        .fields.Append "工作单位", adLongVarChar, 200, adFldIsNullable
        .fields.Append "登记时间", adLongVarChar, 30, adFldIsNullable
        .fields.Append "病人类型", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
     '是否按病人信息为主查询
    Set cllPatiCons = New Collection
    Set cllPageCons = New Collection
    
    blnPatiQuery = False
    For i = 1 To cllFilter.count
        If InStr(",登记时间,病人ID,姓名,就诊卡号,门诊号,医保号,身份证号,IC卡号,", "," & cllFilter(i)(0) & ",") > 0 Then
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
            blnPatiQuery = True
        ElseIf InStr(",入院日期,出院日期,", "," & cllFilter(i)(0) & ",") Then
            '住院条件
             cllPageCons.Add cllFilter(i), cllFilter(i)(0)
        Else
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
        End If
    Next
    
    If blnPatiQuery = False Then
    
        '1.先查询病案主页的数据
        '  (0-在院病人;1-出院病人;2-在院或出院 )
        If zl_CisSvr_GetPatPageInfByRange(2, cllPageCons, "", "", cllPageData) = False Then Exit Function
        If cllPageData Is Nothing Then Set cllPageData = New Collection
        
        str病人Ids = ""
        For i = 1 To cllPageData.count
            '病人性质
            Set cllTemp = cllPageData(i)
            If InStr(str病人Ids & ",", "," & cllTemp("_pati_id") & ",") = 0 Then
                str病人Ids = str病人Ids & "," & cllTemp("_pati_id")
            End If
        Next
        If str病人Ids = "" Then Exit Function
        str病人Ids = Mid(str病人Ids, 2)
       '2.再以病人为查询条件
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData, str病人Ids, str病区IDs) = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
    Else
        '1.按病人信息查询
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData, "", "") = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
        
        str病人Ids = ""
        For i = 1 To cllPatiData.count
            Set cllTemp = cllPatiData(i)
            If InStr(str病人Ids & ",", "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid") & ",") = 0 Then
                str病人Ids = str病人Ids & "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid")
            End If
        Next
        If str病人Ids = "" Then Exit Function
        
        '2.查询病案主页信息
        str病人Ids = Mid(str病人Ids, 2)
        If zl_CisSvr_GetPatPageInfByRange(2, cllPageCons, "", "", cllPageData) = False Then Exit Function
        
    End If
    
    '组装数据
    For i = 1 To cllPatiData.count
       Set cllTemp = cllPatiData(i)
       lng病人ID = Val(cllTemp("_pati_id"))
    
       Set cllPageTemp = Nothing
       For J = 1 To cllPageData.count
          Set cllTemp1 = cllPatiData(J)
          If cllTemp1("_pati_id") = lng病人ID And cllTemp1("pati_pageid") = cllTemp("pati_pageid") Then Set cllPageTemp = cllTemp1: Exit For
       Next
       If Not cllPageTemp Is Nothing Then
           With rsPatiInfo_Out
                  !病人ID = lng病人ID
                  !门诊号 = cllTemp("_outpatient_num")
                  !住院号 = cllTemp("_inpatient_num")

                  If mblnShowCard Then
                     !就诊卡号 = LPAD("*", Len(cllTemp("_vcard_no")))
                  Else
                     !就诊卡号 = cllTemp("_vcard_no")
                  End If
                  
                  !姓名 = cllTemp("_pati_name")
                  !性别 = cllTemp("_pati_sex")
                  !年龄 = cllTemp("_pati_age")
                  !门诊费别 = cllTemp("_fee_category")
                  !病区 = cllTemp("_pati_wardarea_name")
                  !科室 = cllTemp("_pati_dept_name")
                  !床号 = cllTemp("_pati_bed")
                  
                  !入院时间 = cllPageTemp("_adta_time")
                  !出院时间 = cllPageTemp("_adtd_time")
                  !住院次数 = cllTemp("_inp_times")
                  
                  !出生日期 = cllTemp("_pati_birthdate")
                  !国籍 = cllTemp("_country_name")
                  !民族 = cllTemp("_pati_nation")
                  !区域 = cllTemp("_pati_area")
                  !学历 = cllTemp("_pati_education")
                  !职业 = cllTemp("_ocpt_name")
                  !身份 = cllTemp("_pati_identity")
                  !身份证号 = cllTemp("_pati_idcard")
                  !家庭地址 = cllTemp("_pat_home_addr")
                  !工作单位 = cllTemp("_emp_name")
                  !登记时间 = cllTemp("_create_time")
                  If cllPageTemp("_pati_type") <> "" Then
                       !病人类型 = cllPageTemp("_pati_type")
                  Else
                       !病人类型 = IIf(Val(cllPageTemp("_insurance_type")) = 0, "普通病人", "医保病人")
                  End If
               rsPatiInfo_Out.Update
           End With
       End If

      Next

    GetAllPatientData = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function




Private Function GetHospitalizationPatientData(ByVal str病区IDs As String, ByVal cllFilter As Collection, ByRef rsPatiInfo_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取在院病人信息
    '入参:cllFilter-过滤条件
    '     str病区IDs-指定病区
    '出参:rsPatiInfo_Out-返回病人数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 15:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiCons As Collection, cllPageCons As Collection
    Dim cllPatiData As Collection, cllPageData As Collection
    Dim cllTemp As Collection, cllTemp1 As Collection, cllPageTemp As Collection
    Dim blnPatiQuery As Boolean, str病人Ids As String, lng病人ID As Long
    Dim i As Long, J As Long
    
    On Error GoTo errHandle

'    If Mid(tvw_s.SelectedItem.Key, 1, 1) = "D" Then
'        lngUnitID = Mid(tvw_s.SelectedItem.Key, 2)
'        strIF = strIF & " And P.当前病区ID+0= [1] "
'    Else
'        If blnLimitUnit Then
'            strIF = strIF & " And Instr(','||[2]||',',','||P.当前病区ID||',')>0"
'        End If
'    End If
'
'    strSQL = "Select A.病人ID,A.住院号," & strCard & "A.姓名,A.性别,A.年龄,P.费别 as 住院费别," & _
'        " B.名称 as 病区,C.名称 as 科室,A.当前床号 as 床号,To_Char(A.入院时间,'YYYY-MM-DD') as 入院时间," & _
'        " A.住院次数,To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
'        " A.身份证号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间,Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型" & _
'        " From 病案主页 P,病人信息 A,部门表 B,部门表 C" & _
'        " Where A.在院=1 And A.当前病区ID=B.ID And A.当前科室ID=C.ID" & strIF & _
'        " And A.病人ID=P.病人ID And A.主页ID=P.主页ID And Nvl(P.主页ID,0)<>0 And P.出院日期 is NULL " & _
'        " Order by A.入院时间 Desc,A.住院号 Desc"
    '
    Set rsPatiInfo_Out = New ADODB.Recordset
    With rsPatiInfo_Out
        If .State = 1 Then .Close
        .fields.Append "病人ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "住院号", adLongVarChar, 18, adFldIsNullable
        .fields.Append "就诊卡号", adLongVarChar, 100, adFldIsNullable
    
        .fields.Append "姓名", adLongVarChar, 100, adFldIsNullable
        .fields.Append "性别", adLongVarChar, 20, adFldIsNullable
        .fields.Append "年龄", adLongVarChar, 20, adFldIsNullable
        .fields.Append "住院费别", adLongVarChar, 50, adFldIsNullable
        
        .fields.Append "病区", adLongVarChar, 100, adFldIsNullable
        .fields.Append "科室", adLongVarChar, 100, adFldIsNullable
        .fields.Append "床号", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "入院时间", adLongVarChar, 30, adFldIsNullable
        .fields.Append "出院时间", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "住院次数", adVarNumeric, 18, adFldIsNullable
        .fields.Append "出生日期", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "国籍", adLongVarChar, 100, adFldIsNullable
        .fields.Append "民族", adLongVarChar, 100, adFldIsNullable
        .fields.Append "区域", adLongVarChar, 100, adFldIsNullable
        .fields.Append "学历", adLongVarChar, 100, adFldIsNullable
        .fields.Append "职业", adLongVarChar, 100, adFldIsNullable
        .fields.Append "身份", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "身份证号", adLongVarChar, 30, adFldIsNullable
        .fields.Append "家庭地址", adLongVarChar, 200, adFldIsNullable
        .fields.Append "工作单位", adLongVarChar, 200, adFldIsNullable
        .fields.Append "登记时间", adLongVarChar, 30, adFldIsNullable
        .fields.Append "病人类型", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
     '是否按病人信息为主查询
    Set cllPatiCons = New Collection
    Set cllPageCons = New Collection
    
    blnPatiQuery = False
    For i = 1 To cllFilter.count
        If InStr(",登记时间,病人ID,姓名,就诊卡号,门诊号,医保号,身份证号,IC卡号,", "," & cllFilter(i)(0) & ",") > 0 Then
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
            blnPatiQuery = True
        ElseIf InStr(",入院日期,出院日期,", "," & cllFilter(i)(0) & ",") Then
            '住院条件
             cllPageCons.Add cllFilter(i), cllFilter(i)(0)
        Else
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
        End If
    Next
    
    If blnPatiQuery = False Then
    
        '1.先查询病案主页的数据
        '  (0-在院病人;1-出院病人;2-在院或出院 )
        If zl_CisSvr_GetPatPageInfByRange(0, cllPageCons, "", str病区IDs, cllPageData) = False Then Exit Function
        If cllPageData Is Nothing Then Set cllPageData = New Collection
        
        str病人Ids = ""
        For i = 1 To cllPageData.count
            '病人性质
            Set cllTemp = cllPageData(i)
            If InStr(str病人Ids & ",", "," & cllTemp("_pati_id") & ",") = 0 Then
                str病人Ids = str病人Ids & "," & cllTemp("_pati_id")
            End If
        Next
        If str病人Ids = "" Then Exit Function
        str病人Ids = Mid(str病人Ids, 2)
       '2.再以病人为查询条件
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData, str病人Ids, "") = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
    Else
        '1.按病人信息查询
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData, "", "") = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
        
        str病人Ids = ""
        For i = 1 To cllPatiData.count
            Set cllTemp = cllPatiData(i)
            If InStr(str病人Ids & ",", "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid") & ",") = 0 Then
                str病人Ids = str病人Ids & "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid")
            End If
        Next
        If str病人Ids = "" Then Exit Function
        
        '2.查询病案主页信息
        str病人Ids = Mid(str病人Ids, 2)
        If zl_CisSvr_GetPatPageInfByRange(0, cllPageCons, "", str病区IDs, cllPageData) = False Then Exit Function
        
    End If
    
    '组装数据
    For i = 1 To cllPatiData.count
       Set cllTemp = cllPatiData(i)
       lng病人ID = Val(cllTemp("_pati_id"))
    
       Set cllPageTemp = Nothing
       For J = 1 To cllPageData.count
          Set cllTemp1 = cllPatiData(J)
          If cllTemp1("_pati_id") = lng病人ID And cllTemp1("pati_pageid") = cllTemp("pati_pageid") Then Set cllPageTemp = cllTemp1: Exit For
       Next
       If Not cllPageTemp Is Nothing Then
           With rsPatiInfo_Out
               rsPatiInfo_Out.AddNew
                  !病人ID = lng病人ID
                    !住院号 = cllTemp("_inpatient_num")

                  If mblnShowCard Then
                     !就诊卡号 = LPAD("*", Len(cllTemp("_vcard_no")))
                  Else
                     !就诊卡号 = cllTemp("_vcard_no")
                  End If
                  
                  !姓名 = cllTemp("_pati_name")
                  !性别 = cllTemp("_pati_sex")
                  !年龄 = cllTemp("_pati_age")
                  !病区 = cllTemp("_pati_wardarea_name")
                  !科室 = cllTemp("_pati_dept_name")
                  !床号 = cllTemp("_pati_bed")
                  
                  !入院时间 = cllPageTemp("_adta_time")
                  !出院时间 = cllPageTemp("_adtd_time")
                  !住院次数 = cllTemp("_inp_times")
                  
                  !出生日期 = cllTemp("_pati_birthdate")
                  !国籍 = cllTemp("_country_name")
                  !民族 = cllTemp("_pati_nation")
                  !区域 = cllTemp("_pati_area")
                  !学历 = cllTemp("_pati_education")
                  !职业 = cllTemp("_ocpt_name")
                  !身份 = cllTemp("_pati_identity")
                  !身份证号 = cllTemp("_pati_idcard")
                  !家庭地址 = cllTemp("_pat_home_addr")
                  !工作单位 = cllTemp("_emp_name")
                  !登记时间 = cllTemp("_create_time")
                  If cllPageTemp("_pati_type") <> "" Then
                       !病人类型 = cllPageTemp("_pati_type")
                  Else
                       !病人类型 = IIf(Val(cllPageTemp("_insurance_type")) = 0, "普通病人", "医保病人")
                  End If
               rsPatiInfo_Out.Update
           End With
        End If
    Next

    
    GetHospitalizationPatientData = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Function GetLeavePatientData(ByVal str病区IDs As String, ByVal cllFilter As Collection, ByRef rsPatiInfo_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取出院病人信息
    '入参:cllFilter-过滤条件
    '     str病区IDs-指定病区
    '出参:rsPatiInfo_Out-返回病人数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 15:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiCons As Collection, cllPageCons As Collection
    Dim cllPatiData As Collection, cllPageData As Collection
    Dim cllTemp As Collection, cllTemp1 As Collection, cllPageTemp As Collection
    Dim blnPatiQuery As Boolean, str病人Ids As String, lng病人ID As Long
    Dim i As Long, J As Long
    On Error GoTo errHandle
    
'
'    strIF = strIF & IIf(blnLimitUnit, " And Instr(','||[2]||',',','||P.当前病区ID||',')>0", "")
'
'    strSQL = "Select A.病人ID,A.住院号," & strCard & "A.姓名,A.性别,A.年龄,P.费别 as 住院费别," & _
'    " To_Char(A.入院时间,'YYYY-MM-DD') as 入院时间,To_Char(A.出院时间,'YYYY-MM-DD') as 出院时间," & _
'    " A.住院次数,To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
'    " A.身份证号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间,Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型" & _
'    " From 病案主页 P,病人信息 A" & _
'    " Where A.病人ID=P.病人ID And A.主页ID=P.主页ID" & _
'    " And Nvl(P.主页ID,0)<>0 And P.出院日期 Is Not NULL " & strIF & _
'    " Order by A.出院时间 Desc,A.住院号"
    '
    Set rsPatiInfo_Out = New ADODB.Recordset
    With rsPatiInfo_Out
        If .State = 1 Then .Close
        .fields.Append "病人ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "住院号", adLongVarChar, 18, adFldIsNullable
        .fields.Append "就诊卡号", adLongVarChar, 100, adFldIsNullable
    
        .fields.Append "姓名", adLongVarChar, 100, adFldIsNullable
        .fields.Append "性别", adLongVarChar, 20, adFldIsNullable
        .fields.Append "年龄", adLongVarChar, 20, adFldIsNullable
        .fields.Append "住院费别", adLongVarChar, 50, adFldIsNullable
        
        .fields.Append "入院时间", adLongVarChar, 30, adFldIsNullable
        .fields.Append "出院时间", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "住院次数", adVarNumeric, 18, adFldIsNullable
        .fields.Append "出生日期", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "国籍", adLongVarChar, 100, adFldIsNullable
        .fields.Append "民族", adLongVarChar, 100, adFldIsNullable
        .fields.Append "区域", adLongVarChar, 100, adFldIsNullable
        .fields.Append "学历", adLongVarChar, 100, adFldIsNullable
        .fields.Append "职业", adLongVarChar, 100, adFldIsNullable
        .fields.Append "身份", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "身份证号", adLongVarChar, 30, adFldIsNullable
        .fields.Append "家庭地址", adLongVarChar, 200, adFldIsNullable
        .fields.Append "工作单位", adLongVarChar, 200, adFldIsNullable
        .fields.Append "登记时间", adLongVarChar, 30, adFldIsNullable
        .fields.Append "病人类型", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
     '是否按病人信息为主查询
    Set cllPatiCons = New Collection
    Set cllPageCons = New Collection
    
    blnPatiQuery = False
    For i = 1 To cllFilter.count
        If InStr(",登记时间,病人ID,姓名,就诊卡号,门诊号,医保号,身份证号,IC卡号,", "," & cllFilter(i)(0) & ",") > 0 Then
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
            blnPatiQuery = True
        ElseIf InStr(",入院日期,出院日期,", "," & cllFilter(i)(0) & ",") Then
            '住院条件
             cllPageCons.Add cllFilter(i), cllFilter(i)(0)
        Else
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
        End If
    Next
    
    If blnPatiQuery = False Then
    
        '1.先查询病案主页的数据
        '  (0-在院病人;1-出院病人;2-在院或出院 )
        If zl_CisSvr_GetPatPageInfByRange(1, cllPageCons, "", str病区IDs, cllPageData) = False Then Exit Function
        If cllPageData Is Nothing Then Set cllPageData = New Collection
        
        str病人Ids = ""
        For i = 1 To cllPageData.count
            '病人性质
            Set cllTemp = cllPageData(i)
            If InStr(str病人Ids & ",", "," & cllTemp("_pati_id") & ",") = 0 Then
                str病人Ids = str病人Ids & "," & cllTemp("_pati_id")
            End If
        Next
        If str病人Ids = "" Then Exit Function
        str病人Ids = Mid(str病人Ids, 2)
       '2.再以病人为查询条件
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData, str病人Ids, "") = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
    Else
        '1.按病人信息查询
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData) = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
        
        str病人Ids = ""
        For i = 1 To cllPatiData.count
            Set cllTemp = cllPatiData(i)
            If InStr(str病人Ids & ",", "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid") & ",") = 0 Then
                str病人Ids = str病人Ids & "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid")
            End If
        Next
        If str病人Ids = "" Then Exit Function
        
        '2.查询病案主页信息
        str病人Ids = Mid(str病人Ids, 2)
        If zl_CisSvr_GetPatPageInfByRange(2, cllPageCons, "", str病区IDs, cllPageData) = False Then Exit Function
        
    End If
    
    '组装数据
    For i = 1 To cllPatiData.count
       Set cllTemp = cllPatiData(i)
       lng病人ID = Val(cllTemp("_pati_id"))
    
       Set cllPageTemp = Nothing
       For J = 1 To cllPageData.count
          Set cllTemp1 = cllPatiData(J)
          If cllTemp1("_pati_id") = lng病人ID And cllTemp1("pati_pageid") = cllTemp("pati_pageid") Then Set cllPageTemp = cllTemp1: Exit For
       Next
       If Not cllPageTemp Is Nothing Then
           With rsPatiInfo_Out
               rsPatiInfo_Out.AddNew
                  !病人ID = lng病人ID
                    !住院号 = cllTemp("_inpatient_num")

                  If mblnShowCard Then
                     !就诊卡号 = LPAD("*", Len(cllTemp("_vcard_no")))
                  Else
                     !就诊卡号 = cllTemp("_vcard_no")
                  End If
                  
                  !姓名 = cllTemp("_pati_name")
                  !性别 = cllTemp("_pati_sex")
                  !年龄 = cllTemp("_pati_age")
                  !住院费别 = cllPageTemp("_fee_category")
                  !入院时间 = cllPageTemp("_adta_time")
                  !出院时间 = cllPageTemp("_adtd_time")
                  !住院次数 = cllTemp("_inp_times")
                  
                  !出生日期 = cllTemp("_pati_birthdate")
                  !国籍 = cllTemp("_country_name")
                  !民族 = cllTemp("_pati_nation")
                  !区域 = cllTemp("_pati_area")
                  !学历 = cllTemp("_pati_education")
                  !职业 = cllTemp("_ocpt_name")
                  !身份 = cllTemp("_pati_identity")
                  !身份证号 = cllTemp("_pati_idcard")
                  !家庭地址 = cllTemp("_pat_home_addr")
                  !工作单位 = cllTemp("_emp_name")
                  !登记时间 = cllTemp("_create_time")
                  If cllPageTemp("_pati_type") <> "" Then
                       !病人类型 = cllPageTemp("_pati_type")
                  Else
                       !病人类型 = IIf(Val(cllPageTemp("_insurance_type")) = 0, "普通病人", "医保病人")
                  End If
               rsPatiInfo_Out.Update
           End With
        End If
    Next
    
    GetLeavePatientData = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Function GetOutPatientData(ByVal cllFilter As Collection, ByRef rsPatiInfo_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取门诊病人信息
    '入参:cllFilter-过滤条件
    '出参:rsPatiInfo_Out-返回病人数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 15:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiData As Collection, cllTemp As Collection
    Dim i As Long, lng病人ID As Long
    
    
     On Error GoTo errHandle
    
       
    'strSQL = "Select A.病人ID,A.门诊号," & strCard & "A.姓名,A.性别,A.年龄," & _
    " A.费别 as " & IIf(glngSys Like "8??", "会员等级", "门诊费别") & "," & _
    " To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
    " A.身份证号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间,Decode(A.险类,Null,'普通病人','医保病人') 病人类型" & _
    " From 病人信息 A " & _
    " Where A.当前病区ID is NULL And A.当前科室ID is NULL And A.主页ID is NULL" & strIF & _
    " Order by A.登记时间 Desc,A.门诊号 Desc"
    '
    Set rsPatiInfo_Out = New ADODB.Recordset
    With rsPatiInfo_Out
        If .State = 1 Then .Close
        .fields.Append "病人ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "门诊号", adVarNumeric, 18, adFldIsNullable
        .fields.Append "就诊卡号", adLongVarChar, 100, adFldIsNullable
    
        .fields.Append "姓名", adLongVarChar, 100, adFldIsNullable
        .fields.Append "性别", adLongVarChar, 20, adFldIsNullable
        .fields.Append "年龄", adLongVarChar, 20, adFldIsNullable
        .fields.Append "门诊费别", adLongVarChar, 50, adFldIsNullable
        .fields.Append "出生日期", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "国籍", adLongVarChar, 100, adFldIsNullable
        .fields.Append "民族", adLongVarChar, 100, adFldIsNullable
        .fields.Append "区域", adLongVarChar, 100, adFldIsNullable
        .fields.Append "学历", adLongVarChar, 100, adFldIsNullable
        .fields.Append "职业", adLongVarChar, 100, adFldIsNullable
        .fields.Append "身份", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "身份证号", adLongVarChar, 30, adFldIsNullable
        .fields.Append "家庭地址", adLongVarChar, 200, adFldIsNullable
        .fields.Append "工作单位", adLongVarChar, 200, adFldIsNullable
        .fields.Append "登记时间", adLongVarChar, 30, adFldIsNullable
        .fields.Append "病人类型", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    If zl_PatiSvr_GetPatiInfsByRange(0, cllFilter, cllPatiData) = False Then Exit Function
    If cllPatiData Is Nothing Then Set cllPatiData = New Collection
    If cllPatiData.count = 0 Then Exit Function
    '组装数据
    For i = 1 To cllPatiData.count
        Set cllTemp = cllPatiData(i)
        lng病人ID = Val(cllTemp("_pati_id"))
        With rsPatiInfo_Out
            rsPatiInfo_Out.AddNew
                !病人ID = lng病人ID
                !门诊号 = cllTemp("_outpatient_num")
                
                If mblnShowCard Then
                    !就诊卡号 = LPAD("*", Len(cllTemp("_vcard_no")))
                Else
                    !就诊卡号 = cllTemp("_vcard_no")
                End If
                
                !姓名 = cllTemp("_pati_name")
                !性别 = cllTemp("_pati_sex")
                !年龄 = cllTemp("_pati_age")
                !门诊费别 = cllTemp("_fee_category")
                !出生日期 = cllTemp("_pati_birthdate")
                !国籍 = cllTemp("_country_name")
                !民族 = cllTemp("_pati_nation")
                !区域 = cllTemp("_pati_area")
                !学历 = cllTemp("_pati_education")
                !职业 = cllTemp("_ocpt_name")
                !身份 = cllTemp("_pati_identity")
                !身份证号 = cllTemp("_pati_idcard")
                !家庭地址 = cllTemp("_pat_home_addr")
                !工作单位 = cllTemp("_emp_name")
                !登记时间 = cllTemp("_create_time")
                If cllTemp("_pati_type") <> "" Then
                    !病人类型 = cllTemp("_pati_type")
                Else
                    !病人类型 = IIf(Val(cllTemp("_insurance_type")) = 0, "普通病人", "医保病人")
                End If
            rsPatiInfo_Out.Update
        End With
    Next
     
    
    GetOutPatientData = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetObservationPatiData(ByVal cllFilter As Collection, ByRef rsPatiInfo_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取留观病人信息
    '入参:cllFilter-过滤条件
    '出参:rsPatiInfo_Out-返回病人数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 15:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiCons As Collection, cllPageCons As Collection
    Dim cllPatiData As Collection, cllPageData As Collection
    Dim cllTemp As Collection, cllTemp1 As Collection, cllPageTemp As Collection
    Dim blnPatiQuery As Boolean, str病人Ids As String
    Dim i As Long, lng病人ID As Long, J As Long
    
    
    On Error GoTo errHandle

    '    '门诊留观和住院留观病人
    '    strSQL = "Select Distinct A.病人ID,Decode(P.病人性质,1,'门诊留观','住院留观') as 性质, A.门诊号," & strCard & "A.姓名,A.性别,A.年龄," & _
    '    " A.费别 as " & IIf(glngSys Like "8??", "会员等级", "门诊费别") & "," & _
    '    " To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
    '    " A.身份证号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间,Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型" & _
    '    " From 病案主页 P,病人信息 A " & _
    '    " Where A.病人ID=P.病人ID And P.病人性质<>0 And A.住院号 is Null " & strIF & _
    '    " Order by 性质,登记时间 Desc"
    '
    Set rsPatiInfo_Out = New ADODB.Recordset
    With rsPatiInfo_Out
        If .State = 1 Then .Close
        .fields.Append "病人ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "性质", adLongVarChar, 10, adFldIsNullable
        .fields.Append "门诊号", adLongVarChar, 18, adFldIsNullable
        .fields.Append "就诊卡号", adLongVarChar, 100, adFldIsNullable
    
        .fields.Append "姓名", adLongVarChar, 100, adFldIsNullable
        .fields.Append "性别", adLongVarChar, 20, adFldIsNullable
        .fields.Append "年龄", adLongVarChar, 20, adFldIsNullable
        .fields.Append "门诊费别", adLongVarChar, 50, adFldIsNullable
        .fields.Append "出生日期", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "国籍", adLongVarChar, 100, adFldIsNullable
        .fields.Append "民族", adLongVarChar, 100, adFldIsNullable
        .fields.Append "区域", adLongVarChar, 100, adFldIsNullable
        .fields.Append "学历", adLongVarChar, 100, adFldIsNullable
        .fields.Append "职业", adLongVarChar, 100, adFldIsNullable
        .fields.Append "身份", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "身份证号", adLongVarChar, 30, adFldIsNullable
        .fields.Append "家庭地址", adLongVarChar, 200, adFldIsNullable
        .fields.Append "工作单位", adLongVarChar, 200, adFldIsNullable
        .fields.Append "登记时间", adLongVarChar, 30, adFldIsNullable
        .fields.Append "病人类型", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
     '是否按病人信息为主查询
    Set cllPatiCons = New Collection
    Set cllPageCons = New Collection
    
    blnPatiQuery = False
    For i = 1 To cllFilter.count
        If InStr(",登记时间,病人ID,姓名,就诊卡号,门诊号,医保号,身份证号,IC卡号,", "," & cllFilter(i)(0) & ",") > 0 Then
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
            blnPatiQuery = True
        ElseIf InStr(",入院日期,出院日期,", "," & cllFilter(i)(0) & ",") Then
            '住院条件
             cllPageCons.Add cllFilter(i), cllFilter(i)(0)
        Else
           cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
        End If
    Next
    
    cllPageCons.Add Array("病人性质", "1,2"), "病人性质"
    
    '只查留观病人
    If blnPatiQuery = False Then
    
        '1.先查询病案主页的数据
        If zl_CisSvr_GetPatPageInfByRange(2, cllPageCons, "", "", cllPageData) = False Then Exit Function
        If cllPageData Is Nothing Then Set cllPageData = New Collection
        str病人Ids = ""
        For i = 1 To cllPageData.count
            '病人性质
            Set cllTemp = cllPageData(i)
            If InStr(str病人Ids & ",", "," & cllTemp("_pati_id") & ",") = 0 Then
                str病人Ids = str病人Ids & "," & cllTemp("_pati_id")
            End If
        Next
        If str病人Ids = "" Then Exit Function
        str病人Ids = Mid(str病人Ids, 2)
       '2.再以病人为查询条件
        
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData, str病人Ids) = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
    Else
        '1.按病人信息查询
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData) = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
        
        str病人Ids = ""
        For i = 1 To cllPatiData.count
            Set cllTemp = cllPatiData(i)
            If InStr(str病人Ids & ",", "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid") & ",") = 0 Then
                str病人Ids = str病人Ids & "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid")
            End If
        Next
        If str病人Ids = "" Then Exit Function
        
        '2.查询病案主页信息
        str病人Ids = Mid(str病人Ids, 2)
        If zl_CisSvr_GetPatPageInfByRange(2, cllPageCons, "", "", cllPageData) = False Then Exit Function
    End If
    
    '组装数据
    For i = 1 To cllPatiData.count
       Set cllTemp = cllPatiData(i)
       lng病人ID = Val(cllTemp("_pati_id"))
    
       Set cllPageTemp = Nothing
       For J = 1 To cllPageData.count
          Set cllTemp1 = cllPatiData(J)
          If cllTemp1("_pati_id") = lng病人ID And cllTemp1("pati_pageid") = cllTemp("pati_pageid") Then Set cllPageTemp = cllTemp1: Exit For
       Next
       If Not cllPageTemp Is Nothing Then
           With rsPatiInfo_Out
               rsPatiInfo_Out.AddNew
                  !病人ID = lng病人ID
                  !性质 = Decode(Val(cllPageTemp("_pati_nature")), 1, "门诊留观", 2, "住院留观", "非留观")
                  !门诊号 = cllTemp("_outpatient_num")

                  If mblnShowCard Then
                     !就诊卡号 = LPAD("*", Len(cllTemp("_vcard_no")))
                  Else
                     !就诊卡号 = cllTemp("_vcard_no")
                  End If
                  
                  !姓名 = cllTemp("_pati_name")
                  !性别 = cllTemp("_pati_sex")
                  !年龄 = cllTemp("_pati_age")
                  !门诊费别 = cllTemp("_fee_category")
                  !出生日期 = cllTemp("_pati_birthdate")
                  !国籍 = cllTemp("_country_name")
                  !民族 = cllTemp("_pati_nation")
                  !区域 = cllTemp("_pati_area")
                  !学历 = cllTemp("_pati_education")
                  !职业 = cllTemp("_ocpt_name")
                  !身份 = cllTemp("_pati_identity")
                  !身份证号 = cllTemp("_pati_idcard")
                  !家庭地址 = cllTemp("_pat_home_addr")
                  !工作单位 = cllTemp("_emp_name")
                  !登记时间 = cllTemp("_create_time")
                  If cllPageTemp("_pati_type") <> "" Then
                       !病人类型 = cllPageTemp("_pati_type")
                  Else
                       !病人类型 = IIf(Val(cllPageTemp("_insurance_type")) = 0, "普通病人", "医保病人")
                  End If
               rsPatiInfo_Out.Update
           End With
        End If
    Next
    GetObservationPatiData = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function


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
        
        If Not Visible Or blnSet Then
            If gobjComLib Is Nothing Then zlInitCommLib
            If Not gobjComLib Is Nothing Then Call gobjComLib.RestoreFlexState(mshPati, App.ProductName & "\" & Me.Name)
        End If
        
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
        
        .col = 0: .ColSel = .Cols - 1
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
        
        Call ShowPatis(Nothing, True)
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
 
    sta.SimpleText = "正在定位满足条件的病人,按ESC终止 ..."
 
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshPati.Rows - 1
        DoEvents
        
        '比较条件
        blnFill = True
        With mfrmFind
            If .txt病人ID.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("病人ID")) = .txt病人ID.Text
            End If
            If .txt就诊卡.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("就诊卡")) = .txt就诊卡.Text
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
            mshPati.col = 0: mshPati.ColSel = mshPati.Cols - 1
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

