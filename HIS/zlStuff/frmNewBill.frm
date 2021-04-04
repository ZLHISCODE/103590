VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewBill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "发料单据异常处理"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7170
   Icon            =   "frmNewBill.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7170
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtShow 
      Height          =   2535
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "frmNewBill.frx":014A
      Top             =   1200
      Width           =   6735
   End
   Begin VB.CommandButton Cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6000
      TabIndex        =   8
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame fraBottom 
      Height          =   75
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   7065
   End
   Begin MSComctlLib.ProgressBar prg进度条 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4290
      Visible         =   0   'False
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4800
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtPatiId 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   795
      Width           =   2415
   End
   Begin VB.ComboBox cboDept 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   240
      Picture         =   "frmNewBill.frx":0155
      Top             =   150
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "录入病人信息或病人所在病区来查询是否存在未产生的发料单据，如果存在就自动重新产生。"
      Height          =   420
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   6060
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "病  区"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label lblPatiId 
      AutoSize        =   -1  'True
      Caption         =   "门诊号↓"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   840
   End
   Begin VB.Menu mnuPati 
      Caption         =   "病人"
      Visible         =   0   'False
      Begin VB.Menu mnuPatiItem 
         Caption         =   "门诊号(&0)"
         Index           =   0
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "住院号(&1)"
         Index           =   1
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "就诊卡号(&2)"
         Index           =   2
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "姓名(&3)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmNewBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjServiceCall As Object           '服务
Private mintType As Integer '1-按病人id查找；2-按病区查找

Private Enum FindType
    门诊号 = 0
    住院号
    就诊卡号
    姓名
End Enum

Private mCol显示信息 As New Collection
Private mstrShow As String
Public Sub ShowForm(frmMain As Form, Optional ByVal intType As Integer = 1)
    '程序入口
    mintType = intType
    
    Me.Show vbModal, frmMain
End Sub

Private Sub cmdOK_Click()
    Dim intResult As Integer
    Dim colInput As New Collection, colPati As New Collection
    Dim i As Integer
    Dim rsPati As adodb.Recordset, rsSelPati As adodb.Recordset
    Dim strErrMsg As String
    Dim strPatiOut As String, strPatiIDs As String
    Dim varList As Variant  '集合元素
    Dim cllErrMsg As Collection '错误信息集，成员(Array(错误类型,病人姓名,错误信息,单据S(N)))
        
    On Error GoTo ErrHandle
    If mintType = 1 Then
        If txtPatiId.Text = "" Then
            MsgBox "请先输入病人" & Replace(lblPatiId.Caption, "↓", "") & "！", vbInformation, gstrSysName
            zlControl.ControlSetFocus txtPatiId
            Exit Sub
        End If
        
        If Val(txtPatiId.Tag) = 0 Then
            colInput.Add Null, "pati_id"
            colInput.Add Null, "outpatient_num"
            colInput.Add Null, "inpatient_num"
            colInput.Add Null, "pati_wardarea_id"
            colInput.Add Null, "pati_bed"
            colInput.Add Null, "pati_deptid"
            colInput.Add Null, "pati_name"
            colInput.Add Null, "pati_vcard_no"
    
            '病人姓名
            Select Case Val(lblPatiId.Tag)
            Case FindType.门诊号
                If Not IsNumeric(txtPatiId.Text) Then
                    MsgBox "门诊号无效，请重新输入！", vbInformation, gstrSysName
                    zlControl.ControlSetFocus txtPatiId: zlControl.TxtSelAll txtPatiId
                    Exit Sub
                End If
                
                colInput.Remove ("outpatient_num")
                colInput.Add Val(txtPatiId.Text), "outpatient_num"
            
            Case FindType.住院号
                If Not IsNumeric(txtPatiId.Text) Then
                    MsgBox "住院号无效，请重新输入！", vbInformation, gstrSysName
                    zlControl.ControlSetFocus txtPatiId: zlControl.TxtSelAll txtPatiId
                    Exit Sub
                End If
                
                '通过住院号找病人ID
                If zlSplitService_GetPatiId(mobjServiceCall, 1342, txtPatiId.Text, strPatiOut) = False Then Exit Sub
                If Val(strPatiOut) = 0 Then Exit Sub
                
                '病人id
                colInput.Remove ("pati_id")
                colInput.Add Val(strPatiOut), "pati_id"
            Case FindType.就诊卡号
                colInput.Remove ("pati_vcard_no")
                colInput.Add txtPatiId.Text, "pati_vcard_no"
            Case FindType.姓名
                colInput.Remove ("pati_name")
                colInput.Add txtPatiId.Text, "pati_name"
            End Select
            
            If zlSplitService_GetPatiName(mobjServiceCall, 1342, colInput, colPati) = False Then Exit Sub
            If colPati.Count = 0 Then
                MsgBox "未找到对应的病人信息！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If colPati.Count = 1 Then
                txtPatiId.Text = IIf(colPati(1)("_pati_dept_name") = "", "", colPati(1)("_pati_dept_name") & "-") & colPati(1)("_pati_name")
                txtPatiId.Tag = Val(colPati(1)("_pati_id"))
            Else
                '返回多条记录时
                Set rsPati = New adodb.Recordset
                With rsPati
                    If .State = 1 Then .Close
                    .Fields.Append "病人id", adDouble, 18, adFldIsNullable
                    .Fields.Append "病人姓名", adLongVarChar, 20, adFldIsNullable
                    .Fields.Append "住院号", adDouble, 18, adFldIsNullable
                    .Fields.Append "病区", adLongVarChar, 30, adFldIsNullable
                    .Fields.Append "床号", adLongVarChar, 20, adFldIsNullable
                    .Fields.Append "科室id", adDouble, 18, adFldIsNullable
                    .Fields.Append "科室", adLongVarChar, 30, adFldIsNullable
                    
                    .CursorLocation = adUseClient
                    .CursorType = adOpenStatic
                    .LockType = adLockOptimistic
                    .Open
                    
                    For i = 1 To colPati.Count
                        .AddNew
                        
                        !病人ID = colPati(i)("_pati_id")
                        !病人姓名 = colPati(i)("_pati_name")
                        !住院号 = colPati(i)("_inpatient_num")
                        !病区 = colPati(i)("_pati_wardarea_name")
                        !床号 = colPati(i)("_pati_bed")
                        !科室id = colPati(i)("_pati_dept_id")
                        !科室 = colPati(i)("_pati_dept_name")
                        
                        .Update
                    Next
                End With
                
                If zlDatabase.zlShowListSelect(Me, 100, 1342, txtPatiId, rsPati, True, "", "病人ID,科室ID", rsSelPati) = False Then Exit Sub
            
                rsSelPati.Filter = ""
                If rsSelPati.RecordCount = 0 Then Exit Sub
                
                txtPatiId.Text = IIf(rsSelPati!科室 = "", "", rsSelPati!科室 & "-") & rsSelPati!病人姓名
                txtPatiId.Tag = rsSelPati!病人ID
            End If
        End If

        '检查并重新产生处方
        If Val(txtPatiId.Tag) = 0 Then Exit Sub
        
        intResult = ExecuteDataSync(Val(txtPatiId.Tag), cllErrMsg)
        strErrMsg = GetErrMsg(cllErrMsg)
        Select Case intResult
        Case 0
            MsgBox "病人【" & Mid(txtPatiId.Text, InStr(txtPatiId.Text, "-") + 1) & "】未产生的发料单据已重新产生完成！", vbInformation, gstrSysName
        Case 1
            MsgBox "病人【" & Mid(txtPatiId.Text, InStr(txtPatiId.Text, "-") + 1) & "】不存在未产生的发料单据！", vbInformation, gstrSysName
        Case 2
            MsgBox "在检查病人【" & Mid(txtPatiId.Text, InStr(txtPatiId.Text, "-") + 1) & "】是否存在未产生的发料单据时出现错误！" & _
                    IIf(strErrMsg = "", "", vbCrLf & vbCrLf & strErrMsg), vbInformation, gstrSysName
        Case 3
            MsgBox "病人【" & Mid(txtPatiId.Text, InStr(txtPatiId.Text, "-") + 1) & "】未产生的部分发料单据重新产生时失败！" & _
                IIf(strErrMsg = "", "", vbCrLf & vbCrLf & strErrMsg), vbInformation, gstrSysName
        End Select
        zlControl.ControlSetFocus txtPatiId: zlControl.TxtSelAll txtPatiId
        
        Call Show产生数据信息
        
        Exit Sub
    End If
    
    '2-按病区查找
    If Val(cboDept.ItemData(cboDept.ListIndex)) <= 0 Then
        MsgBox "请先选择一个病区！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '调服务获取病区病人id
    If zlSplitService_GetPatiByRange(mobjServiceCall, 1342, Val(cboDept.ItemData(cboDept.ListIndex)), colPati) = False Then Exit Sub
    If colPati.Count = 0 Then
        MsgBox "【" & cboDept.Text & "】不存在未产生的发料单据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '根据病人id检查并同步异常费用
    For Each varList In colPati
        strPatiIDs = strPatiIDs & "," & Val(varList("_pati_id"))
    Next
    
    intResult = ExecuteDataSync(Mid(strPatiIDs, 2), cllErrMsg)
    strErrMsg = GetErrMsg(cllErrMsg)
    Select Case intResult
    Case 0
        MsgBox "【" & cboDept.Text & "】未产生的发料单据已重新产生完成！", vbInformation, gstrSysName
    Case 1
        MsgBox "【" & cboDept.Text & "】不存在未产生的发料单据！", vbInformation, gstrSysName
    Case 2
        MsgBox "在检查【" & cboDept.Text & "】是否存在未产生的发料单据时出现错误！" & _
            IIf(strErrMsg = "", "", vbCrLf & vbCrLf & strErrMsg), vbInformation, gstrSysName
    Case 3
        MsgBox "【" & cboDept.Text & "】未产生的部分发料单据重新产生时失败！" & _
            IIf(strErrMsg = "", "", vbCrLf & vbCrLf & strErrMsg), vbInformation, gstrSysName
    End Select
    
    Call Show产生数据信息
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetErrMsg(ByVal cllErrMsg As Collection, Optional ByVal bln显示姓名 As Boolean) As String
    '获取错误信息
    '入参：
    '   cllErrMsg-错误信息集，成员(Array(错误类型,病人姓名,错误信息,单据S(N)))  错误类型：2-临床域同步检查失败；1-费用域同步检查失败；0-其它错误
    Dim i As Long, strMsg As String, strErrInfo As String
    Dim lngCount As Long, bytErrType As Byte, strInfo As String
    
    If cllErrMsg Is Nothing Then Exit Function
    
    strMsg = "": lngCount = 0
    For i = 1 To cllErrMsg.Count
        bytErrType = cllErrMsg(i)(0)
        
        strErrInfo = cllErrMsg(i)(2)
        If InStr(UCase(strErrInfo), "[ZLSOFT]") > 0 Then strErrInfo = Split(strErrInfo, "[ZLSOFT]")(1)
        
        strInfo = ""
        If strErrInfo <> "" Then
            If lngCount > 5 Then '超过5个省略号表示
                strMsg = strMsg & vbCrLf & "……"
                Exit For
            End If
            
            strInfo = (lngCount + 1) & "、"
            If cllErrMsg(i)(1) <> "" And bln显示姓名 Then strInfo = strInfo & cllErrMsg(i)(1) & " "
            If bytErrType = 2 Then
                strInfo = strInfo & "[" & cllErrMsg(i)(3) & "] 无法同步，请回退医嘱重新发送。原因："
            ElseIf bytErrType = 1 Then
                strInfo = strInfo & "[" & cllErrMsg(i)(3) & "] 无法同步，请作废费用重新记费。原因："
            Else
                strInfo = strInfo & "同步失败，请重试。原因："
            End If
            strInfo = strInfo & strErrInfo
            
            If strInfo <> "" And InStr(vbCrLf & strMsg & vbCrLf, vbCrLf & strInfo & vbCrLf) = 0 Then
                strMsg = IIf(strMsg = "", "", strMsg & vbCrLf) & strInfo
                lngCount = lngCount + 1
            End If
        End If
    Next
    If lngCount = 1 Then strMsg = Mid(strMsg, 3)
    
    GetErrMsg = strMsg
End Function

Private Sub Cmd取消_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If mintType = 2 Then
        lblDept.Visible = True
        cboDept.Visible = True
        lblPatiId.Visible = False
        txtPatiId.Visible = False
        Call LoadDept
    End If

    '实例化服务
    Call zlSercieCall_Ini(mobjServiceCall)
    mobjServiceCall.InitService gcnOracle, gstrDBUser, glngSys, glngModul
End Sub

Private Sub LoadDept()
    Dim rsTemp As adodb.Recordset, strSQL As String
    
    On Error GoTo ErrHandle
    cboDept.Clear
    cboDept.Tag = ""
    
    strSQL = _
        " Select b.名称 As 站点名称, b.编号 As 站点,A.编码||'-'||A.名称 科室,A.ID" & _
        " From 部门表 A, Zlnodelist B " & _
        " Where a.站点 = b.编号(+) And A.ID in (Select 部门ID From 部门性质说明 Where 工作性质 in('护理','临床') And 服务对象 IN(2,3))" & _
        "           And (A.撤档时间 Is Null Or A.撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
        " Order By a.站点, a.编码 || '-' || a.名称 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取科室")
    Do While Not rsTemp.EOF
        cboDept.AddItem rsTemp!科室
        cboDept.ItemData(cboDept.NewIndex) = rsTemp!Id
        rsTemp.MoveNext
    Loop
    If cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call zlSercieCall_Unload(mobjServiceCall)
End Sub


Private Sub lblPatiId_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        PopupMenu mnuPati, 2, lblPatiId.Left + lblPatiId.Width - 30, lblPatiId.Top
    End If
End Sub

Private Sub mnuPatiItem_Click(index As Integer)
    Dim i As Integer
    
    lblPatiId.Tag = index
    txtPatiId.Text = ""
    txtPatiId.MaxLength = 0
    
    Select Case index
        Case FindType.门诊号
            lblPatiId.Caption = "门诊号↓"
            lblPatiId.Tag = FindType.门诊号
        Case FindType.住院号
            lblPatiId.Caption = "住院号↓"
            lblPatiId.Tag = FindType.住院号
        Case FindType.就诊卡号
            lblPatiId.Caption = "就诊卡号↓"
            lblPatiId.Tag = FindType.就诊卡号
        Case FindType.姓名
            lblPatiId.Caption = "姓名↓"
            lblPatiId.Tag = FindType.姓名
    End Select
    
    For i = 0 To mnuPatiItem.Count - 1
        mnuPatiItem(i).Checked = (i = index)
    Next
End Sub

Private Sub RefrashProgress(Optional ByVal lngValue As Long, Optional ByVal bytMode As Byte = 1, Optional ByVal lngMaxValue As Long)
    '刷新进度显示
    '入参:
    '   bytMode-类型，0-刷新信息，1-初始化显示，2-终止显示
    On Error GoTo ErrHandler
    Select Case bytMode
    Case 0
        Me.MousePointer = vbHourglass
        prg进度条.Visible = True
        prg进度条.Value = 0
        prg进度条.Max = lngMaxValue
    Case 1
        prg进度条.Value = lngValue
    Case 2
        prg进度条.Visible = False
        Me.MousePointer = vbDefault
    End Select
    Exit Sub
ErrHandler:
    prg进度条.Visible = False
    Me.MousePointer = vbDefault
End Sub

Private Function ExecuteDataSync(ByVal strPatiIDs As String, ByRef cllErrMsg_Out As Collection) As Integer
    '执行异常数据同步
    '入参：
    '   strPatiIDs-病人ID，多个用英文逗号分隔
    '出参：
    '   cllErrMsg_Out-错误信息集，成员(Array(错误类型,病人姓名,错误信息,单据S(N)))
    '返回：0-存在未产生的发料单据，且重新全部产生；1-不存在未产生的发料单据；2-其他错误；3-存在未产生的发料单据，部分重新产生成功
    '说明:
    '   1.临床域同步异常，按“病人+发送"进行同步
    '   2.费用域异常，按“单据”进行同步
    Dim cllCisErrData As Collection, cllExseErrData As Collection, cllPatiData As Collection
    Dim cllOrderSendItem As Collection, cllPatiBillItem As Collection
    Dim i As Long, lngCount As Long, lngSccussCount As Long, strErrMsg As String
    Dim cllPati As Collection, bytErrType As Byte, strNos As String
    
    On Error GoTo ErrHandler
    Set cllErrMsg_Out = New Collection
    
    Me.MousePointer = vbHourglass
    '1.根据病人ID取医嘱数据
    ExecuteDataSync = GetCisSyncErrData(strPatiIDs, cllCisErrData, strErrMsg)
    If ExecuteDataSync = 2 Then
        cllErrMsg_Out.Add Array(bytErrType, "", strErrMsg)
        Me.MousePointer = vbDefault: Exit Function
    End If
    
    Me.MousePointer = vbHourglass
    '2.取费用数据
    ExecuteDataSync = GetExseSyncErrData(strPatiIDs, cllCisErrData, cllExseErrData, strErrMsg)
    If ExecuteDataSync = 2 Or ExecuteDataSync = 1 Then
        cllErrMsg_Out.Add Array(bytErrType, "", strErrMsg)
        Me.MousePointer = vbDefault: Exit Function
    End If

    Me.MousePointer = vbHourglass
    '3.获取病人信息：身份，出生日期，身份证号
    If GetPatiData(cllExseErrData, cllPatiData, strErrMsg) = False Then
        ExecuteDataSync = 2
        cllErrMsg_Out.Add Array(bytErrType, "", strErrMsg)
        Me.MousePointer = vbDefault: Exit Function
    End If
    
    Call RefrashProgress(, 0, cllCisErrData.Count)
    lngCount = 0: lngSccussCount = 0
    
    '4.修正临床域同步异常，同步后从 cllExseErrData 移除
    For Each cllOrderSendItem In cllCisErrData
        If ExecuteCisErrDataSync(cllOrderSendItem, cllExseErrData, cllPatiData, strErrMsg, bytErrType, strNos) = False Then
            If ExistsColObject(cllPatiData, "_" & cllOrderSendItem("病人ID")) Then
                Set cllPati = cllPatiData("_" & cllOrderSendItem("病人ID"))
                cllErrMsg_Out.Add Array(bytErrType, cllPati("姓名"), strErrMsg, strNos)
            Else
                cllErrMsg_Out.Add Array(bytErrType, "", strErrMsg, strNos)
            End If
        Else
            lngSccussCount = lngSccussCount + 1
        End If
        bytErrType = 0
        
        lngCount = lngCount + 1
        Call RefrashProgress(lngCount)
    Next
    If cllCisErrData.Count <> lngSccussCount Then ExecuteDataSync = 3
    
    Call RefrashProgress(, 0, cllExseErrData.Count)
    lngCount = 0: lngSccussCount = 0
    
    '5.修正费用域同步异常
    For Each cllPatiBillItem In cllExseErrData
        If ExecuteExseErrDataSync(cllPatiBillItem, cllPatiData, strErrMsg, bytErrType, strNos) = False Then
            If (cllPatiBillItem("单据类型")) = 3 Then '记帐表
                cllErrMsg_Out.Add Array(bytErrType, "", strErrMsg, strNos)
            Else
                If ExistsColObject(cllPatiData, "_" & cllPatiBillItem("病人ID")) Then
                    Set cllPati = cllPatiData("_" & cllPatiBillItem("病人ID"))
                    cllErrMsg_Out.Add Array(bytErrType, cllPati("姓名"), strErrMsg, strNos)
                Else
                    cllErrMsg_Out.Add Array(bytErrType, "", strErrMsg, strNos)
                End If
            End If
        Else
            lngSccussCount = lngSccussCount + 1
        End If
        bytErrType = 0
        
        lngCount = lngCount + 1
        Call RefrashProgress(lngCount)
    Next
    If cllExseErrData.Count <> lngSccussCount Then ExecuteDataSync = 3
    
    Call RefrashProgress(, 2)
    Me.MousePointer = vbDefault
    Exit Function
ErrHandler:
    cllErrMsg_Out.Add Array(bytErrType, "", err.Description)
    Me.MousePointer = vbDefault
    Call RefrashProgress(, 2)
    ExecuteDataSync = 2
End Function

Private Function ExecuteCisErrDataSync(ByVal cllOrderSendItem As Collection, ByRef cllExseErrData As Collection, _
    ByVal cllPatiData As Collection, ByRef strErrMsg As String, ByRef bytErrType As Byte, ByRef strNos As String) As Boolean
    '执行临床域异常数据同步
    '入参：
    '   cllOrderSendItem-病人医嘱发送记录，成员(病人ID,主页ID,挂号ID,挂号单号,发送号,OrderList)
    '           |-cllOrderList-医嘱信息列表=cllOrderSendItem(OrderList)
    '               |-cllOrderItem-医嘱信息，成员(医嘱ID,医嘱期效,紧急标志,计价特性)=cllOrderList(_医嘱ID)
    '           |-cllExseBillList-费用单据列表=cllOrderSendItem(ExseBillList)
    '               |-cllExseBillItem-费用单据信息，成员(费用来源,单据类型,单据号)=cllExseBillList(_费用来源_单据类型_单据号)
    '       其中，单据类型：1-收费单,2-记帐单,3-记帐表；费用来源：1-门诊,2-住院
    '   cllExseErrData=费用域同步异常数据，说明：括号中的均为集合Key值
    '       |-cllPatiBillItem-病人单据记录，成员(单据类型,病人来源,[病人ID,主页ID,姓名,性别编号,性别,年龄,病人科室ID,病人病区ID],BillLists)；其中，方括号中的元素记帐表时无
    '           |-cllBillLists-单据信息集=cllPatiBillItem(BillLists)
    '               |-cllBillItem-单据信息，成员(费用来源,NO,收费标志,划价人,开单科室ID,开单科室名称,
    '                                                          开单医师ID,开单医师,操作员姓名,操作员编号,登记时间,DetailList)=cllBillLists(_费用来源_单据类型_单据号)
    '                   |-cllDetailList-单据明细集=cllBillItem(DetailList)
    '                       |-cllDetailItem-每行明细数据集，成员([病人ID,主页ID,姓名,性别编号,性别,年龄,病人病区ID,病人科室ID],
    '                                 费用ID,序号,库房ID,是否备货卫材,批次,卫材ID,婴儿序号,医嘱ID,付数,数量,售价,零售金额,摘要)；其中，方括号中的元素记帐表时才有
    '       其中，单据类型：1-收费单,2-记帐单,3-记帐表；病人来源：1-门诊,2-住院,4-体检；费用来源：1-门诊,2-住院；
    '                 处方类型：0和空-普通,1-儿科,2-急诊,3-精二,4-精一,5-麻醉；收费标志：0-未收费或记帐划价,1-已收费或记帐
    '   cllPatiData=病人信息数据，说明：括号中的均为集合Key值
    '       |-cllPatiItem-病人信息，成员(病人ID,出生日期,身份证号,身份")=cllPatiData(_病人ID)
    '出参：
    '   strErrMsg=错误信息
    '   bytErrType=错误类型：2-临床域同步检查失败；0-其它错误
    '   strNos=涉及的单据号，格式：A001,A002,...
    '返回:执行成功返回True，执行失败返回False
    Dim strJson As String, strListJson As String, strOrders As String
    Dim cllOrderList As Collection, cllOrderItem As Collection
    Dim cllExseBillList As Collection, cllExseBillItem As Collection
    Dim strNewBillCheckJson As String, strNewBillJson As String, strSyncJson As String
    Dim blnTrans As Boolean, strKey As String
    Dim cllPatiBillItem As Collection, cllBillLists As Collection
    
    On Error GoTo ErrHandler
    strErrMsg = "": bytErrType = 0: strNos = ""
    If cllOrderSendItem Is Nothing Then ExecuteCisErrDataSync = True: Exit Function
    
    Set cllOrderList = cllOrderSendItem("OrderList")
    Set cllExseBillList = cllOrderSendItem("ExseBillList")
    
    If GetNewBillJson_Cis(cllOrderSendItem, cllExseErrData, cllPatiData, _
        strNewBillCheckJson, strNewBillJson, strErrMsg, strNos) = False Then GoTo MoveExseNOsHandler
    
    If strNewBillJson = "" Then '无费用单据，跳过
        ExecuteCisErrDataSync = True
        GoTo MoveExseNOsHandler
    End If
    
    bytErrType = 2
    If mobjServiceCall.CallService("Zl_卫材销售出库_Check", strNewBillCheckJson, , , , False, , , , True) = False Then
        strErrMsg = "调用产生新的处方检查失败！": GoTo MoveExseNOsHandler
    End If
    bytErrType = 0
    
    '获取临床域同步数据JSON
    'Zl_CisSvr_UpdateSyncState
    '  --功能：同步标记录更新
    '  --入参：Json_In:格式
    '  --  input
    '  --      order_list[]
    '  --          order_id          N 1 医嘱id
    '  --          send_no           N 1 发送号
    '  --          sign_type         N 1 设置标记录的类型，说明：1-清除静配标记录,2-清除 生成药品同步标记,3-清除 生成卫材同步标记
    strListJson = ""
    For Each cllOrderItem In cllOrderList
        If InStr("," & strOrders & ",", ",3:" & cllOrderItem("医嘱ID") & ",") = 0 Then
            strOrders = strOrders & ",3:" & cllOrderItem("医嘱ID")
            
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("order_id", cllOrderItem("医嘱ID"), 1)
            strJson = strJson & "," & GetJsonNodeString("send_no", cllOrderSendItem("发送号"), 1)
            strJson = strJson & "," & GetJsonNodeString("sign_type", 3, 1)
            strListJson = strListJson & ",{" & strJson & "}"
        End If
    Next
    strSyncJson = "{""input"":{""order_list"":[" & Mid(strListJson, 2) & "]}}"
    
    gcnOracle.BeginTrans: blnTrans = True
        If mobjServiceCall.CallService("Zl_药品收发记录_Newstuffbill", strNewBillJson, , , , False, , , , True) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            strErrMsg = "调用过程产生药品单据数据失败！": GoTo MoveExseNOsHandler
        End If
        
        If mobjServiceCall.CallService("Zl_CisSvr_UpdateSyncState", strSyncJson, , , , False, , , , True) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            strErrMsg = "调用服务修改医嘱同步标志失败！": GoTo MoveExseNOsHandler
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    '执行成功的数据加入显示集合
    mCol显示信息.Add mCol显示信息.Count + 1 & " " & mstrShow
    
    ExecuteCisErrDataSync = True
    
MoveExseNOsHandler:
    '移除医嘱涉及的费用单据
    If cllExseBillList Is Nothing Then Exit Function
    For Each cllExseBillItem In cllExseBillList
        strKey = "_" & cllExseBillItem("费用来源") & "_" & cllExseBillItem("单据类型") & "_" & cllExseBillItem("单据号")
        Dim i As Long
        For i = cllExseErrData.Count To 1 Step -1
            Set cllPatiBillItem = cllExseErrData(i)
            Set cllBillLists = cllPatiBillItem("BillLists")
            If ExistsColObject(cllBillLists, strKey) Then
                cllBillLists.Remove strKey
                If cllBillLists.Count = 0 Then cllExseErrData.Remove i
                Exit For
            End If
        Next
    Next
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    strErrMsg = err.Description
    GoTo MoveExseNOsHandler
End Function

Private Function ExecuteExseErrDataSync(ByVal cllPatiBillItem As Collection, ByVal cllPatiData As Collection, _
    ByRef strErrMsg As String, ByRef bytErrType As Byte, ByRef strNos As String) As Boolean
    '执行临床域异常数据同步
    '入参：
    '   cllPatiBillItem-病人单据记录，成员(单据类型,病人来源,[病人ID,主页ID,姓名,性别编号,性别,年龄,病人科室ID,病人病区ID],BillLists)；其中，方括号中的元素记帐表时无
    '           |-cllBillLists-单据信息集=cllPatiBillItem(BillLists)
    '               |-cllBillItem-单据信息，成员(费用来源,NO,收费标志,划价人,开单科室ID,开单科室名称,
    '                                                          开单医师ID,开单医师,操作员姓名,操作员编号,登记时间,DetailList)=cllBillLists(_费用来源_单据类型_单据号)
    '                   |-cllDetailList-单据明细集=cllBillItem(DetailList)
    '                       |-cllDetailItem-每行明细数据集，成员([病人ID,主页ID,姓名,性别编号,性别,年龄,病人病区ID,病人科室ID],
    '                                 费用ID,序号,库房ID,是否备货卫材,批次,卫材ID,婴儿序号,医嘱ID,付数,数量,售价,零售金额,摘要)；其中，方括号中的元素记帐表时才有
    '       其中，单据类型：1-收费单,2-记帐单,3-记帐表；病人来源：1-门诊,2-住院,4-体检；费用来源：1-门诊,2-住院；
    '                 处方类型：0和空-普通,1-儿科,2-急诊,3-精二,4-精一,5-麻醉；收费标志：0-未收费或记帐划价,1-已收费或记帐
    '   cllPatiData=病人信息数据，说明：括号中的均为集合Key值
    '       |-cllPatiItem-病人信息，成员(病人ID,出生日期,身份证号,身份")=cllPatiData(_病人ID)
    '出参：
    '   strErrMsg=错误信息
    '   bytErrType=错误类型：1-费用域同步检查失败；0-其它错误
    '   strNos=涉及的单据号，格式：A001,A002,...
    '返回:执行成功返回True，执行失败返回False
    Dim strSyncJson As String, str费用ids As String
    Dim strNewBillCheckJson As String, strNewBillJson As String
    Dim cllBillLists As Collection, cllBillItem As Collection
    Dim cllDetailItem  As Collection, cllDetailList As Collection
    Dim blnTrans As Boolean
    
    On Error GoTo ErrHandler
    strErrMsg = "": bytErrType = 0: strNos = ""
    If cllPatiBillItem Is Nothing Then ExecuteExseErrDataSync = True: Exit Function
    
    Set cllBillLists = cllPatiBillItem("BillLists")
    
    If GetNewBillJson_Exse(cllPatiBillItem, cllPatiData, strNewBillCheckJson, strNewBillJson, strErrMsg, strNos) = False Then Exit Function
    
    bytErrType = 1
    If mobjServiceCall.CallService("Zl_卫材销售出库_Check", strNewBillCheckJson, , , , False, , , , True) = False Then
        strErrMsg = "调用产生新的处方检查失败！": Exit Function
    End If
    bytErrType = 0
    
    str费用ids = ""
    '获取费用域同步数据JSON
    'Zl_Exsesvr_Sync_Update
    '      ---------------------------------------------------------------------------
    '  --功能：费用同步后清空记费同步标志（按NO或按费用ID）
    '  --入参：Json_In:格式
    '  --  input
    '  --    sign_type           N 1 标志类型：0-记费同步标志,1-转费同步标志
    '  --    detail_ids  C  1  处方明细id串(费用id串),支持多个id，用“,”分隔
    '  --    bill_list[]
    '  --      billtype               N   1 单据类型:1-收费处方;2-记帐处方
    '  --      rcp_no                 C   1 处方No
    For Each cllBillItem In cllBillLists
        Set cllDetailList = cllBillItem("DetailList")
        For Each cllDetailItem In cllDetailList
            str费用ids = str费用ids & "," & cllDetailItem("费用ID")
        Next
    Next
    strSyncJson = "{""input"":{""sign_type"":0,""detail_ids"":""" & Mid(str费用ids, 2) & """}}"
    
    gcnOracle.BeginTrans: blnTrans = True
        If mobjServiceCall.CallService("Zl_药品收发记录_Newstuffbill", strNewBillJson, , , , False, , , , True) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            strErrMsg = "调用过程产生药品单据数据失败！": Exit Function
        End If
        
        If mobjServiceCall.CallService("Zl_Exsesvr_Sync_Update", strSyncJson, , , , False, , , , True) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            strErrMsg = "调用服务修改记费同步标志失败！": Exit Function
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    '执行成功的数据加入显示集合
    mCol显示信息.Add mCol显示信息.Count + 1 & " " & mstrShow
    
    ExecuteExseErrDataSync = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    strErrMsg = err.Description
End Function

Private Sub Show产生数据信息()
    Dim strShow As String
    Dim i As Integer
    
    If Not mCol显示信息 Is Nothing Then
        If mCol显示信息.Count > 0 Then
            For i = 1 To mCol显示信息.Count
                strShow = IIf(strShow = "", "", strShow & vbCrLf) & mCol显示信息(i)
            Next
        End If
    End If
    
    If strShow = "" Then
        txtShow.Text = "无数据重新产生！"
    Else
        txtShow.Text = "以下数据已重新产生：" & vbCrLf & strShow
    End If
End Sub
Private Function GetNewBillJson_Cis(ByVal cllOrderSendItem As Collection, ByVal cllExseErrData As Collection, ByVal cllPatiData As Collection, _
    ByRef strNewBillCheckJson_Out As String, ByRef strNewBillJson_Out As String, ByRef strErrMsg As String, ByRef strNos As String) As Boolean
    '执行临床域异常数据同步
    '入参：
    '   cllOrderSendItem-病人医嘱发送记录，成员(病人ID,主页ID,挂号ID,挂号单号,发送号,OrderList)
    '           |-cllOrderList-医嘱信息列表=cllOrderSendItem(OrderList)
    '               |-cllOrderItem-医嘱信息，成员(医嘱ID,医嘱期效,紧急标志,计价特性)=cllOrderList(_医嘱ID)
    '           |-cllExseBillList-费用单据列表=cllOrderSendItem(ExseBillList)
    '               |-cllExseBillItem-费用单据信息，成员(费用来源,单据类型,单据号)=cllExseBillList(_费用来源_单据类型_单据号)
    '       其中，单据类型：1-收费单,2-记帐单,3-记帐表；费用来源：1-门诊,2-住院
    '   cllExseErrData=费用域同步异常数据，说明：括号中的均为集合Key值
    '       |-cllPatiBillItem-病人单据记录，成员(单据类型,病人来源,[病人ID,主页ID,姓名,性别编号,性别,年龄,病人科室ID,病人病区ID],BillLists)；其中，方括号中的元素记帐表时无
    '           |-cllBillLists-单据信息集=cllPatiBillItem(BillLists)
    '               |-cllBillItem-单据信息，成员(费用来源,NO,收费标志,划价人,开单科室ID,开单科室名称,
    '                                                          开单医师ID,开单医师,操作员姓名,操作员编号,登记时间,DetailList)=cllBillLists(_费用来源_单据类型_单据号)
    '                   |-cllDetailList-单据明细集=cllBillItem(DetailList)
    '                       |-cllDetailItem-每行明细数据集，成员([病人ID,主页ID,姓名,性别编号,性别,年龄,病人病区ID,病人科室ID],
    '                                 费用ID,序号,库房ID,是否备货卫材,批次,卫材ID,婴儿序号,医嘱ID,付数,数量,售价,零售金额,摘要)；其中，方括号中的元素记帐表时才有
    '       其中，单据类型：1-收费单,2-记帐单,3-记帐表；病人来源：1-门诊,2-住院,4-体检；费用来源：1-门诊,2-住院；
    '                 处方类型：0和空-普通,1-儿科,2-急诊,3-精二,4-精一,5-麻醉；收费标志：0-未收费或记帐划价,1-已收费或记帐
    '   cllPatiData=病人信息数据，说明：括号中的均为集合Key值
    '       |-cllPatiItem-病人信息，成员(病人ID,出生日期,身份证号,身份")=cllPatiData(_病人ID)
    '出参：
    '   strErrMsg=返回错误信息
    '   strNos=涉及的单据号，格式：A001,A002,...
    '返回:执行成功返回True，执行失败返回False
    Dim cllOrderList As Collection
    Dim cllExseBillList As Collection, cllExseBillItem As Collection
    Dim cllPatiBillItem_New As Collection, cllBillLists_New As Collection
    Dim strKey As String
    Dim cllPatiBillItem As Collection, cllBillLists As Collection
    Dim bln记帐表 As Boolean
    
    On Error GoTo ErrHandler
    strNewBillCheckJson_Out = "": strNewBillJson_Out = "": strErrMsg = "": strNos = ""
    If cllOrderSendItem Is Nothing Then GetNewBillJson_Cis = True: Exit Function
    
    Set cllOrderList = cllOrderSendItem("OrderList")
    Set cllExseBillList = cllOrderSendItem("ExseBillList")
    
    '查找医嘱涉及的费用单据，重组单据记录集
    Set cllPatiBillItem_New = New Collection
    Set cllBillLists_New = New Collection
    For Each cllExseBillItem In cllExseBillList
        strKey = "_" & cllExseBillItem("费用来源") & "_" & cllExseBillItem("单据类型") & "_" & cllExseBillItem("单据号")
        For Each cllPatiBillItem In cllExseErrData
            Set cllBillLists = cllPatiBillItem("BillLists")
            If ExistsColObject(cllBillLists, strKey) Then
                If cllBillLists_New.Count = 0 Then
                    bln记帐表 = (Val(Nvl(cllPatiBillItem("单据类型"))) = 3)
                    cllPatiBillItem_New.Add cllPatiBillItem("单据类型"), "单据类型"
                    cllPatiBillItem_New.Add cllPatiBillItem("病人来源"), "病人来源"
                    If bln记帐表 = False Then
                        cllPatiBillItem_New.Add cllPatiBillItem("病人ID"), "病人ID"
                        cllPatiBillItem_New.Add cllPatiBillItem("主页ID"), "主页ID"
                        cllPatiBillItem_New.Add cllPatiBillItem("姓名"), "姓名"
                        cllPatiBillItem_New.Add cllPatiBillItem("性别编号"), "性别编号"
                        cllPatiBillItem_New.Add cllPatiBillItem("性别"), "性别"
                        cllPatiBillItem_New.Add cllPatiBillItem("年龄"), "年龄"
                        cllPatiBillItem_New.Add cllPatiBillItem("病人科室ID"), "病人科室ID"
                        cllPatiBillItem_New.Add cllPatiBillItem("病人病区ID"), "病人病区ID"
                    End If
                    cllPatiBillItem_New.Add cllBillLists_New, "BillLists"
                End If
                
                cllBillLists_New.Add cllBillLists(strKey), strKey
                Exit For
            End If
        Next
    Next
    If cllPatiBillItem_New.Count = 0 Then
        '无费用单据
        GetNewBillJson_Cis = True: Exit Function
    End If
    
    If GetNewBillJson_Exse(cllPatiBillItem_New, cllPatiData, strNewBillCheckJson_Out, strNewBillJson_Out, strErrMsg, strNos, cllOrderList) = False Then Exit Function
    
    GetNewBillJson_Cis = True
    Exit Function
ErrHandler:
    strErrMsg = err.Description
End Function

Private Function GetNewBillJson_Exse(ByVal cllPatiBillItem As Collection, ByVal cllPatiData As Collection, _
    ByRef strNewBillCheckJson_Out As String, ByRef strNewBillJson_Out As String, ByRef strErrMsg As String, ByRef strNos As String, _
    Optional ByVal cllOrderList As Collection) As Boolean
    '执行临床域异常数据同步
    '入参：
    '   cllPatiBillItem-病人单据记录，成员(单据类型,病人来源,[病人ID,主页ID,姓名,性别编号,性别,年龄,病人科室ID,病人病区ID],BillLists)；其中，方括号中的元素记帐表时无
    '           |-cllBillLists-单据信息集=cllPatiBillItem(BillLists)
    '               |-cllBillItem-单据信息，成员(费用来源,NO,收费标志,划价人,开单科室ID,开单科室名称,
    '                                                          开单医师ID,开单医师,操作员姓名,操作员编号,登记时间,DetailList)=cllBillLists(_费用来源_单据类型_单据号)
    '                   |-cllDetailList-单据明细集=cllBillItem(DetailList)
    '                       |-cllDetailItem-每行明细数据集，成员([病人ID,主页ID,姓名,性别编号,性别,年龄,病人病区ID,病人科室ID],
    '                                 费用ID,序号,库房ID,是否备货卫材,批次,卫材ID,婴儿序号,医嘱ID,付数,数量,售价,零售金额,摘要)；其中，方括号中的元素记帐表时才有
    '       其中，单据类型：1-收费单,2-记帐单,3-记帐表；病人来源：1-门诊,2-住院,4-体检；费用来源：1-门诊,2-住院；
    '                 处方类型：0和空-普通,1-儿科,2-急诊,3-精二,4-精一,5-麻醉；收费标志：0-未收费或记帐划价,1-已收费或记帐
    '   cllPatiData=病人信息数据，说明：括号中的均为集合Key值
    '       |-cllPatiItem-病人信息，成员(病人ID,出生日期,身份证号,身份")=cllPatiData(_病人ID)
    '   cllOrderList-医嘱信息列表
    '               |-cllOrderItem-医嘱信息，成员(医嘱ID,医嘱期效,紧急标志,计价特性)=cllOrderList(_医嘱ID)
    '出参：
    '   strNewBillCheckJson_Out=新单据检查数据JSON
    '   strNewBillJson_Out=新单据保存数据JSON
    '   strErrMsg=返回错误信息
    '   strNos=涉及的单据号，格式：A001,A002,...
    '返回:执行成功返回True，执行失败返回False
    Dim strJson As String, bln记帐表 As Boolean
    Dim cllBillLists As Collection, cllBillItem As Collection
    Dim cllDetailList As Collection, cllDetailItem As Collection
    Dim strBillListJson As String, strDetailListJson As String
    Dim rsTotal As adodb.Recordset, cllOrderItem As Collection
    Dim cllPati As Collection
    Dim strShowNO As String, strShow病人姓名 As String, strShow科室 As String
        
    On Error GoTo ErrHandler
    strNewBillCheckJson_Out = "": strNewBillJson_Out = "": strErrMsg = "": strNos = ""
    If cllPatiBillItem Is Nothing Then GetNewBillJson_Exse = True: Exit Function
    
    Set rsTotal = New adodb.Recordset
    rsTotal.Fields.Append "库房ID", adBigInt, , adFldIsNullable
    rsTotal.Fields.Append "卫材ID", adBigInt, , adFldIsNullable
    rsTotal.Fields.Append "数量", adDouble, , adFldIsNullable
    rsTotal.Fields.Append "单价", adDouble, , adFldIsNullable
    rsTotal.Fields.Append "是否备货卫材", adInteger, , adFldIsNullable
    rsTotal.Fields.Append "批次", adBigInt, , adFldIsNullable
    rsTotal.Fields.Append "费用id", adBigInt, , adFldIsNullable
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open
    
    'Zl_药品收发记录_Newstuffbill
    '  --功能：主要是在记帐（含划价）， 收费(含划价)后产生新的处方或药嘱记录
    '  --入参：Json_In:格式
    '  --  input
    '  --     billtype             N   1 单据类型: 1 -收费处方  ;2- 记帐单处方;3- 记帐表处方
    '  --     pati_source          N   1 病人来源:1-门诊;2-住院;4-体检
    '  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，无以下节点--------------------------------------
    '  --     pati_id                    N   1 病人ID
    '  --     pati_pageid                N   1 主页ID
    '  --     pati_name                  C   1 病人姓名
    '  --     pati_sex_code              C   1 性别编号（新门诊)
    '  --     pati_sex                   C   1 性别
    '  --     pati_age                   C   1 年龄
    '  --     pati_identity              C     身份
    '  --     pati_birthdate             C     出生日期:yyyy-mm-dd hh:mi:ss
    '  --     pati_idcard                C     身份证号
    '  --     pati_deptid                N   1 病人科室ID
    '  --     pati_wardarea_id           N     病人病区ID
    '  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，无以上节点-----------------------------------------
    '  --     bill_list[]                      更新数据列表[数组]
    '  --        stuff_no                  C  1 NO
    '  --        charge_tag                N  1 收费标志:0-未收费或记帐划价;1-已收费或记帐
    '  --        fee_acnter                C    划价人
    '  --        plcdept_id                C    开单科室id（新门诊)
    '  --        plcdept                   C    开单科室名称（新门诊)
    '  --        placer_id                 C    开单医师id（新门诊)
    '  --        placer                    C    开单医师（新门诊)  增加
    '  --        apply_fee_category_code   C    申请单费别编码(医疗付款方式编码)(新门诊) 增加；
    '  --        apply_fee_category_name   C    申请单费别名称（医疗付款方式名称）(新门诊) 增加；
    '  --        operator_name             C  1 操作员姓名
    '  --        operator_code             C  1 操作员编号
    '  --        create_time               C  1 登记时间:yyyy-mm-dd hh:mi:ss
    '  --        item_list[]                    更新数据列表[数组]
    '  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，有以下节点----------------------------------------
    '  --           pati_id                 N  1 病人ID
    '  --           pati_pageid             N    主页ID
    '  --           pati_name               C  1 病人姓名
    '  --           pati_sex                C  1 性别
    '  --           pati_age                C  1 年龄
    '  --           pati_identity           C    身份
    '  --           pati_birthdate          C    出生日期:yyyy-mm-dd hh:mi:ss
    '  --           pati_idcard             C    身份证号
    '  --           pati_wardarea_id        N    病人病区ID
    '  --           pati_deptid             N  1 病人科室ID
    '  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，有以上节点-----------------------------------------
    '  --           stuffdtl_id             N  1 处方明细ID
    '  --           serial_num              N  1 序号
    '  --           warehouse_id            N  1 库房ID
    '  --           is_bakstuff             N  1 是否备货卫材:有高值卫材才需要传入，非0表示是高值卫材模式(如扫码时使用)
    '  --           bakstuff_batch             1 备货材料批次
    '  --           stuff_id                N  1 材料ID
    '  --           baby_num                N    婴儿序号
    '  ---------------------------以下节点为可选参数，医嘱记录产生-----------------------------------------------
    '  --           advice_id               N  0 医嘱ID
    '  --           emergency_tag           N    医嘱记录中的紧急标志(0-普通;1-紧急;2-补录(对门诊无效))
    '  --           effectivetime           N  0 医嘱期效
    '  --           freq_name               C  0 频次名称
    '  --           single                  N  0 单量
    '  ---------------------------以上节点为可选参数，医嘱记录产生-----------------------------------------------
    '  --           packages_num            N  1 付数
    '  --           outbound_num            N  1 出库数量
    '  --           price                   N    售价
    '  --           warehouse_window        C  0 发料窗口
    '  --           memo                    C  0 摘要
    '  --           fee_source              N  0 费用来源
    '  --           stuff_auto_send         N  0 卫材自动发料;0-不自动发料;1-自动发料
    
    bln记帐表 = (cllPatiBillItem("单据类型") = 3)
    Set cllBillLists = cllPatiBillItem("BillLists")
    
    strBillListJson = ""
    For Each cllBillItem In cllBillLists
        strDetailListJson = ""
        Set cllDetailList = cllBillItem("DetailList")
        
        For Each cllDetailItem In cllDetailList
            
            rsTotal.Filter = "库房ID=" & cllDetailItem("库房ID") & " And 卫材ID=" & cllDetailItem("卫材ID") & " And 批次=" & Val(Nvl(cllDetailItem("批次")))
            If rsTotal.EOF Then rsTotal.AddNew
            rsTotal!库房ID = cllDetailItem("库房ID")
            rsTotal!卫材ID = cllDetailItem("卫材ID")
            rsTotal!数量 = Val(Nvl(rsTotal!数量)) + IIf(cllDetailItem("付数") = 0, 1, cllDetailItem("付数")) * cllDetailItem("数量")
            rsTotal!单价 = Val(cllDetailItem("售价"))
            rsTotal!是否备货卫材 = Val(cllDetailItem("是否备货卫材"))
            rsTotal!批次 = Val(cllDetailItem("批次"))
            rsTotal!费用id = Val(cllDetailItem("费用id"))
            
            Set cllOrderItem = Nothing
            If Not cllOrderList Is Nothing And Val(Nvl(cllDetailItem("医嘱ID"))) <> 0 Then
                If ExistsColObject(cllOrderList, "_" & cllDetailItem("医嘱ID")) Then Set cllOrderItem = cllOrderList("_" & cllDetailItem("医嘱ID"))
            End If
            
            strJson = ""
            If bln记帐表 Then
                strJson = strJson & "," & GetJsonNodeString("pati_id", cllDetailItem("病人ID"), 1)
                strJson = strJson & "," & GetJsonNodeString("pati_pageid", cllDetailItem("主页ID"), 1)
                strJson = strJson & "," & GetJsonNodeString("pati_name", cllDetailItem("姓名"), 0)
                strJson = strJson & "," & GetJsonNodeString("pati_sex_code", cllDetailItem("性别编号"), 0)
                strJson = strJson & "," & GetJsonNodeString("pati_sex", cllDetailItem("性别"), 0)
                strJson = strJson & "," & GetJsonNodeString("pati_age", cllDetailItem("年龄"), 0)
                
                Set cllPati = cllPatiData("_" & cllDetailItem("病人ID")) '病人ID,出生日期,身份证号,身份
                strJson = strJson & "," & GetJsonNodeString("pati_identity", cllPati("身份"), 0)
                strJson = strJson & "," & GetJsonNodeString("pati_birthdate", cllPati("出生日期"), 0)
                strJson = strJson & "," & GetJsonNodeString("pati_idcard", cllPati("身份证号"), 0)
                
                strJson = strJson & "," & GetJsonNodeString("pati_wardarea_id", cllDetailItem("病人病区ID"), 1)
                strJson = strJson & "," & GetJsonNodeString("pati_deptid", cllDetailItem("病人科室ID"), 1)
            End If
            strJson = strJson & "," & GetJsonNodeString("stuffdtl_id", cllDetailItem("费用ID"), 1)
            strJson = strJson & "," & GetJsonNodeString("serial_num", cllDetailItem("序号"), 1)
            strJson = strJson & "," & GetJsonNodeString("warehouse_id", cllDetailItem("库房ID"), 1)
            strJson = strJson & "," & GetJsonNodeString("is_bakstuff", cllDetailItem("是否备货卫材"), 1)
            strJson = strJson & "," & GetJsonNodeString("bakstuff_batch", cllDetailItem("批次"), 1)
            strJson = strJson & "," & GetJsonNodeString("stuff_id", cllDetailItem("卫材ID"), 1)
            strJson = strJson & "," & GetJsonNodeString("baby_num", cllDetailItem("婴儿序号"), 1)
            strJson = strJson & "," & GetJsonNodeString("advice_id", cllDetailItem("医嘱ID"), 1)
            If Not cllOrderItem Is Nothing Then
                strJson = strJson & "," & GetJsonNodeString("emergency_tag", cllOrderItem("紧急标志"), 1)
                strJson = strJson & "," & GetJsonNodeString("effectivetime", cllOrderItem("医嘱期效"), 1)
                strJson = strJson & "," & GetJsonNodeString("freq_name", cllOrderItem("频次名称"), 0)
                strJson = strJson & "," & GetJsonNodeString("single", cllOrderItem("单量"), 1)
            End If
            strJson = strJson & "," & GetJsonNodeString("packages_num", cllDetailItem("付数"), 1)
            strJson = strJson & "," & GetJsonNodeString("outbound_num", cllDetailItem("数量"), 1)
            strJson = strJson & "," & GetJsonNodeString("price", cllDetailItem("售价"), 1)
            strJson = strJson & "," & GetJsonNodeString("memo", cllDetailItem("摘要"), 0)
            strJson = strJson & "," & GetJsonNodeString("fee_source", cllBillItem("费用来源"), 1)
            'strJson = strJson & "," & GetJsonNodeString("stuff_auto_send", cllDetailItem(""), 1) '卫材自动发料;0-不自动发料;1-自动发料
            
            strDetailListJson = strDetailListJson & ",{" & Mid(strJson, 2) & "}"
        Next
        
        strJson = ""
        strJson = strJson & "," & GetJsonNodeString("stuff_no", cllBillItem("NO"), 0)
        strJson = strJson & "," & GetJsonNodeString("charge_tag", cllBillItem("收费标志"), 1)
        strJson = strJson & "," & GetJsonNodeString("fee_acnter", cllBillItem("划价人"), 0)
        strJson = strJson & "," & GetJsonNodeString("plcdept_id", cllBillItem("开单科室ID"), 0)
        strJson = strJson & "," & GetJsonNodeString("plcdept", cllBillItem("开单科室名称"), 0)
        strJson = strJson & "," & GetJsonNodeString("placer_id", cllBillItem("开单医师ID"), 0)
        strJson = strJson & "," & GetJsonNodeString("placer", cllBillItem("开单医师"), 0)
        'strJson = strJson & "," & GetJsonNodeString("apply_fee_category_code", cllBillItem(""), 0)'申请单费别编码(医疗付款方式编码)(新门诊)
        'strJson = strJson & "," & GetJsonNodeString("apply_fee_category_name", cllBillItem(""), 0)'申请单费别名称（医疗付款方式名称）(新门诊)
        strJson = strJson & "," & GetJsonNodeString("operator_name", cllBillItem("操作员姓名"), 0)
        strJson = strJson & "," & GetJsonNodeString("operator_code", cllBillItem("操作员编号"), 0)
        strJson = strJson & "," & GetJsonNodeString("create_time", cllBillItem("登记时间"), 0)
        strJson = strJson & ",""item_list"":[" & Mid(strDetailListJson, 2) & "]"
        strBillListJson = strBillListJson & ",{" & Mid(strJson, 2) & "}"
        
        If InStr("," & strNos & ",", "," & cllBillItem("NO") & ",") = 0 Then
            strNos = strNos & "," & cllBillItem("NO")
            
            strShowNO = cllBillItem("NO") & IIf(cllPatiBillItem("单据类型") = 1, "(收费)", "(记账)")
            strShow科室 = cllBillItem("开单科室名称")
            strShow病人姓名 = cllPatiBillItem("姓名") & "(" & cllPatiBillItem("性别") & "," & cllPatiBillItem("年龄") & ")"
            
            mstrShow = strShowNO & " " & strShow病人姓名 & " " & strShow科室
        End If
    Next
    If strNos <> "" Then strNos = Mid(strNos, 2)
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("billtype", cllPatiBillItem("单据类型"), 1)
    strJson = strJson & "," & GetJsonNodeString("pati_source", cllPatiBillItem("病人来源"), 1)
    If bln记帐表 = False Then
        strJson = strJson & "," & GetJsonNodeString("pati_id", cllPatiBillItem("病人ID"), 1)
        strJson = strJson & "," & GetJsonNodeString("pati_pageid", cllPatiBillItem("主页ID"), 1)
        strJson = strJson & "," & GetJsonNodeString("pati_name", cllPatiBillItem("姓名"), 0)
        strJson = strJson & "," & GetJsonNodeString("pati_sex_code", cllPatiBillItem("性别编号"), 0)
        strJson = strJson & "," & GetJsonNodeString("pati_sex", cllPatiBillItem("性别"), 0)
        strJson = strJson & "," & GetJsonNodeString("pati_age", cllPatiBillItem("年龄"), 0)
        
        Set cllPati = cllPatiData("_" & cllPatiBillItem("病人ID")) '病人ID,出生日期,身份证号,身份
        strJson = strJson & "," & GetJsonNodeString("pati_identity", cllPati("身份"), 0)
        strJson = strJson & "," & GetJsonNodeString("pati_birthdate", cllPati("出生日期"), 0)
        strJson = strJson & "," & GetJsonNodeString("pati_idcard", cllPati("身份证号"), 0)
        
        strJson = strJson & "," & GetJsonNodeString("pati_deptid", cllPatiBillItem("病人科室ID"), 1)
        strJson = strJson & "," & GetJsonNodeString("pati_wardarea_id", cllPatiBillItem("病人病区ID"), 1)
    End If
    strJson = strJson & ",""bill_list"":[" & Mid(strBillListJson, 2) & "]"
    strJson = "{""input"":{" & strJson & "}}"
    
    If GetNewBillCheckJson(rsTotal, strNewBillCheckJson_Out) = False Then Exit Function
    
    strNewBillJson_Out = strJson
    GetNewBillJson_Exse = True
    Exit Function
ErrHandler:
    strErrMsg = err.Description
End Function

Private Function GetNewBillCheckJson(ByVal rsTotal As adodb.Recordset, ByRef strCheckJson_Out As String) As Boolean
    '功能:获取生成药品处方检查条件的Json入参串
    '入参:
    '   rsTotal-当前的汇总记录集(卫材ID,库房ID,数量,单价)
    '出参:
    '返回:返回Json串
    Dim strJson As String, strListJson As String
    
    strCheckJson_Out = ""
    If rsTotal Is Nothing Then GetNewBillCheckJson = True: Exit Function

    'Zl_卫材销售出库_Check
    '  --入参      json
    '  --input     根据条件对要产生的处方进行检查
    '  --  fee_list      收费明细信息，支持多个，[数组]
    '  --    stuff_id  N 1 卫材id
    '  --    send_num  N 1 发料数量
    '  --    warehouse_id  N 1 库房id
    '  --    price           N       1       售价
    '  --    is_bakstuff N   是否备货卫材:有高值卫材才需要传入，非0表示是高值卫材模式(如扫码时使用)
    '  --    bakstuff_batch  N   备货材料批次
    '  --    rcpdtl_id    N  1  费用id：0或空-没有传入时忽略；>0传入时检查是否已存在相同的费用ID收发记录
    With rsTotal
        .Filter = ""
        Do While Not .EOF
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("stuff_id", Val(Nvl(!卫材ID)), 1)
            strJson = strJson & "," & GetJsonNodeString("send_num", Val(Nvl(!数量)), 1)
            strJson = strJson & "," & GetJsonNodeString("warehouse_id", Val(Nvl(!库房ID)), 1)
            strJson = strJson & "," & GetJsonNodeString("price", Val(Nvl(!单价)), 1)
            strJson = strJson & "," & GetJsonNodeString("is_bakstuff", Val(Nvl(!是否备货卫材)), 1)
            strJson = strJson & "," & GetJsonNodeString("bakstuff_batch", Val(Nvl(!批次)), 1)
            strJson = strJson & "," & GetJsonNodeString("rcpdtl_id", Val(Nvl(!费用id)), 1)
            strListJson = strListJson & ",{" & strJson & "}"
            .MoveNext
        Loop
    End With
    If strListJson = "" Then GetNewBillCheckJson = True: Exit Function
    
    strCheckJson_Out = "{""input"":{""fee_list"":[" & Mid(strListJson, 2) & "]}}"
    GetNewBillCheckJson = True
End Function

Private Function GetPatiData(ByVal cllExseErrData As Collection, ByRef cllPatiData As Collection, ByRef strErrMsg As String) As Boolean
    '获取病人数据
    '入参：
    '   cllExseErrData=费用域同步异常数据，说明：括号中的均为集合Key值
    '       |-cllPatiBillItem-病人单据记录，成员(单据类型,病人来源,[病人ID,主页ID,姓名,性别编号,性别,年龄,病人科室ID,病人病区ID],BillLists)；其中，方括号中的元素记帐表时无
    '           |-cllBillLists-单据信息集=cllPatiBillItem(BillLists)
    '               |-cllBillItem-单据信息，成员(费用来源,NO,收费标志,划价人,开单科室ID,开单科室名称,
    '                                                          开单医师ID,开单医师,操作员姓名,操作员编号,登记时间,DetailList)=cllBillLists(_费用来源_单据类型_单据号)
    '                   |-cllDetailList-单据明细集=cllBillItem(DetailList)
    '                       |-cllDetailItem-每行明细数据集，成员([病人ID,主页ID,姓名,性别编号,性别,年龄,病人病区ID,病人科室ID],
    '                                 费用ID,序号,库房ID,是否备货卫材,批次,卫材ID,婴儿序号,医嘱ID,付数,数量,售价,零售金额,摘要)；其中，方括号中的元素记帐表时才有
    '       其中，单据类型：1-收费单,2-记帐单,3-记帐表；病人来源：1-门诊,2-住院,4-体检；费用来源：1-门诊,2-住院；收费标志：0-未收费或记帐划价,1-已收费或记帐
    '出参：
    '   cllPatiData=病人信息数据，说明：括号中的均为集合Key值
    '       |-cllPatiItem-病人信息，成员(病人ID,姓名,门诊号,住院号,出生日期,身份证号,身份")=cllPatiData(_病人ID)
    '   strErrMsg=范围值为2时，返回错误信息
    '返回：成功返回True，失败返回False
    Dim bln记帐表 As Boolean, strPatiIDs As String, cllItem As Collection
    Dim cllPatiBillItem As Collection, cllBillLists As Collection
    Dim cllDetailList As Collection, cllDetailItem As Collection
    Dim cllPatiOut As Collection, cllPati As Collection
    Dim p As Long, i As Long, j As Long
    Dim StrJson_In As String
    
    On Error GoTo ErrHandler
    Set cllPatiData = New Collection
    strErrMsg = ""
    
    If cllExseErrData Is Nothing Then GetPatiData = True: Exit Function
    For p = 1 To cllExseErrData.Count
        Set cllPatiBillItem = cllExseErrData(p)
        bln记帐表 = (Val(cllPatiBillItem("单据类型")) = 3)
        
        If bln记帐表 = False Then
            If InStr("," & strPatiIDs & ",", "," & cllPatiBillItem("病人ID") & ",") = 0 Then
                strPatiIDs = strPatiIDs & "," & cllPatiBillItem("病人ID")
            End If
        Else
            Set cllBillLists = cllPatiBillItem("BillLists")
            For i = 1 To cllBillLists.Count
                Set cllDetailList = cllBillLists(i)("DetailList")
                For j = 1 To cllDetailList.Count
                    Set cllDetailItem = cllDetailList(j)
                    If InStr("," & strPatiIDs & ",", "," & cllDetailItem("病人ID") & ",") = 0 Then
                        strPatiIDs = strPatiIDs & "," & cllDetailItem("病人ID")
                    End If
                Next
            Next
        End If
    Next
    
    'Zl_Patisvr_Getpatiinfo
    '  --功能:获取病人信息
    '  --入参：Json_In:格式
    '  --    input
    '  --      pati_id           N 1 病人id  病人ID<>0时，查询列表中的条件无效
    '  --      query_type        N 1 查询类型:如：0-基本;1-基本+联系人;2-所有
    '  --      query_card        N 1 是否包含卡信息:1-包含医疗卡;0-不包含医疗卡
    '  --      query_family      N 1 是否包含家属:1-包含家属信息，0-不包含家属信息
    '  --      query_drug        N 1 是否包含过敏药物:1-包含，0-不包含
    '  --      query_immune      N 1 是否包含免疫修:1-包含;0-不包含
    '  --      query_insurance_pwd C  是否包含医保密码:1-包含;0-不包含
    '  --      query_cons_list   C 1 查询条件:可以选择一定条件进行查询（是And关系),只有一行
    '  --        pati_ids        C   病人IDs:多个用逗号
    '  --        pati_name       C   姓名:可以代%分号表表按姓名匹配
    '  --        outpatient_num  C   门诊号
    '  --        inpatient_num   C   住院号
    '  --        pati_idcard     C   身份证号
    '  --        contacts_idcard C   联系人身份证号
    '  --        cardtype_id     N   医疗卡类别ID
    '  --        medc_card_name  N   医疗卡名称
    '  --        card_no         C   卡号
    '  --        qrcode          C   二维码
    '  --        iccard_no       C   Ic卡号
    '  --        visit_card      C   就诊卡号
    '  --        insurance_num   C   医保号
    '  --        qrspt_statu     C   查询住院状态:0-仅门诊;1-在院 ;2-门诊及在院
    '  --        phone_number    C   手机号
    '  --        pati_bed        C   当前床号
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", 3, 1)
    StrJson_In = StrJson_In & "," & """query_cons_list"":{""qrspt_statu"":2,""pati_ids"":""" & Mid(strPatiIDs, 2) & """}"
    StrJson_In = "{""input"":{" & StrJson_In & "}}"

    If mobjServiceCall.CallService("Zl_Patisvr_Getpatiinfo", StrJson_In, , , , False, , , , True) = False Then
        strErrMsg = "调用服务查询病人信息失败！"
        Exit Function
    End If

    Set cllPatiOut = mobjServiceCall.GetJsonListValue("output.pati_list", "pati_id")
    If cllPatiOut Is Nothing Then Exit Function
    
    For i = 1 To cllPatiOut.Count
        '--    pati_list[]                 病人信息列表
        '--      pati_id             N   1   病人id
        '--      pati_name           C   1   姓名
        '--      outpatient_num      C   1   门诊号
        '--      inpatient_num       C   1   住院号
        '--      pati_birthdate      C   1   出生日期：yyyy-mm-dd hh24:mi:ss
        '  --    pati_idcard         C   1   身份证号
        '--      pati_identity       C   1   身份
        Set cllItem = cllPatiOut(i)
        Set cllPati = New Collection
        cllPati.Add cllItem("_pati_id"), "病人ID"
        cllPati.Add cllItem("_pati_name"), "姓名"
        cllPati.Add cllItem("_outpatient_num"), "门诊号"
        cllPati.Add cllItem("_inpatient_num"), "住院号"
        cllPati.Add cllItem("_pati_birthdate"), "出生日期"
        cllPati.Add cllItem("_pati_idcard"), "身份证号"
        cllPati.Add cllItem("_pati_identity"), "身份"
        cllPatiData.Add cllPati, "_" & cllItem("_pati_id")
    Next
    GetPatiData = True
    Exit Function
ErrHandler:
    strErrMsg = err.Description
End Function

Private Function GetExseSyncErrData(ByVal strPatiIDs As String, ByVal cllCisErrData As Collection, _
    ByRef cllExseErrData As Collection, ByRef strErrMsg As String) As Integer
    '获取医嘱费用数据及费用同步异常数据
    '入参：
    '   strPatiIDs=病人ID,多个用英文的逗号分隔
    '   cllCisErrData-临床域同步异常数据，说明：括号中的均为集合Key值
    '       |-cllOrderSendItem-病人医嘱发送记录，成员(病人ID,主页ID,挂号ID,挂号单号,发送号,OrderList)
    '           |-cllOrderList-医嘱信息列表=cllOrderSendItem(OrderList)
    '               |-cllOrderItem-医嘱信息，成员(医嘱ID,医嘱期效,紧急标志,计价特性)=cllOrderList(_医嘱ID)
    '           |-cllExseBillList-费用单据列表=cllOrderSendItem(ExseBillList)
    '               |-cllExseBillItem-费用单据信息，成员(费用来源,单据类型,单据号)=cllExseBillList(_费用来源_单据类型_单据号)
    '       其中，单据类型：1-收费单,2-记帐单,3-记帐表；费用来源：1-门诊,2-住院
    '出参：
    '   cllExseErrData=费用域同步异常数据，说明：括号中的均为集合Key值
    '       |-cllPatiBillItem-病人单据记录，成员(单据类型,病人来源,[病人ID,主页ID,姓名,性别编号,性别,年龄,病人科室ID,病人病区ID],BillLists)；其中，方括号中的元素记帐表时无
    '           |-cllBillLists-单据信息集=cllPatiBillItem(BillLists)
    '               |-cllBillItem-单据信息，成员(费用来源,NO,收费标志,划价人,开单科室ID,开单科室名称,
    '                                                          开单医师ID,开单医师,操作员姓名,操作员编号,登记时间,DetailList)=cllBillLists(_费用来源_单据类型_单据号)
    '                   |-cllDetailList-单据明细集=cllBillItem(DetailList)
    '                       |-cllDetailItem-每行明细数据集，成员([病人ID,主页ID,姓名,性别编号,性别,年龄,病人病区ID,病人科室ID],
    '                                 费用ID,序号,库房ID,是否备货卫材,批次,卫材ID,婴儿序号,医嘱ID,付数,数量,售价,零售金额,摘要)；其中，方括号中的元素记帐表时才有
    '       其中，单据类型：1-收费单,2-记帐单,3-记帐表；病人来源：1-门诊,2-住院,4-体检；费用来源：1-门诊,2-住院；
    '                 处方类型：0和空-普通,1-儿科,2-急诊,3-精二,4-精一,5-麻醉；收费标志：0-未收费或记帐划价,1-已收费或记帐
    '   strErrMsg=范围值为2时，返回错误信息
    '返回：0-存在未同步的单据，1-不存在未同步的单据，2-其他错误
    Dim StrJson_In As String, strJson_List As String, strJsonItem As String, strJson_PatiList As String
    Dim cllExseBillList As Collection, cllItem As Collection
    Dim p As Long, i As Long, j As Long
    Dim cllOutList As Collection, cllBill_Out As Collection, cllDetail_Out As Collection
    Dim cllPatiBillItem As Collection, cllBillLists As Collection, cllBillItem As Collection
    Dim cllDetailList As Collection, cllDetailItem As Collection
    Dim bln记帐表 As Boolean, varPatiIDs As Variant
    Dim strKey As String, byt单据类型 As Byte
    
    On Error GoTo ErrHandler
    Set cllExseErrData = New Collection
    strErrMsg = ""
    
    If strPatiIDs = "" Then GetExseSyncErrData = 1: Exit Function
    'Zl_Exsesvr_Getstufferrdata
    '  --功能：根据病人ID和医嘱信息返回病人费用信息
    '  --入参：Json_In:格式
    '  --  input
    '  --    pati_list[]病人列表
    '  --       pati_id                    N 1 病人id
    '  --       bill_list[]                费用单据号列表，可以不传，不传时表示获取费用域同步异常的数据
    '  --         fee_source               N 0 费用来源：1-门诊；2-住院
    '  --         fee_billtype             N 0 费用单据类型：1-收费处方；2-记帐单处方
    '  --         fee_no                   C 0 费用单据号
    strJson_PatiList = ""
    varPatiIDs = Split(strPatiIDs, ",")
    For p = 0 To UBound(varPatiIDs)
        strJson_List = ""
        If Not cllCisErrData Is Nothing Then
            For i = 1 To cllCisErrData.Count
                Set cllItem = cllCisErrData(i)
                
                If Val(Nvl(cllItem("病人ID"))) = varPatiIDs(p) Then
                    Set cllExseBillList = cllItem("ExseBillList")
                    For j = 1 To cllExseBillList.Count
                        Set cllItem = cllExseBillList(j)
                        strJsonItem = ""
                        strJsonItem = strJsonItem & "" & GetJsonNodeString("fee_source", cllItem("费用来源"), 1)
                        strJsonItem = strJsonItem & "," & GetJsonNodeString("fee_billtype", cllItem("单据类型"), 1)
                        strJsonItem = strJsonItem & "," & GetJsonNodeString("fee_no", cllItem("单据号"), 0)
                        strJson_List = strJson_List & ",{" & strJsonItem & "}"
                    Next
                End If
            Next
        End If
        
        strJsonItem = ""
        strJsonItem = strJsonItem & "" & GetJsonNodeString("pati_id", varPatiIDs(p), 1)
        If strJson_List <> "" Then
            strJsonItem = strJsonItem & ",""bill_list"":[" & Mid(strJson_List, 2) & "]"
        End If
        strJson_PatiList = strJson_PatiList & ",{" & strJsonItem & "}"
    Next
    StrJson_In = "{""input"":{""pati_list"":[" & Mid(strJson_PatiList, 2) & "]}}"
    
    If mobjServiceCall.CallService("Zl_Exsesvr_Getstufferrdata", StrJson_In, , , , False, , , , True) = False Then
        strErrMsg = "调用费用服务查询未产生单据失败！"
        GetExseSyncErrData = 2: Exit Function
    End If
    
    Set cllOutList = mobjServiceCall.GetJsonListValue("output.pati_bill_list")
    If cllOutList Is Nothing Then GetExseSyncErrData = 1: Exit Function
    If cllOutList.Count = 0 Then GetExseSyncErrData = 1: Exit Function

    '   cllExseErrData=费用域同步异常数据，说明：括号中的均为集合Key值
    '       |-cllPatiBillItem-病人单据记录，成员(单据类型,病人来源,[病人ID,主页ID,姓名,性别编号,性别,年龄,病人科室ID,病人病区ID],BillLists)；其中，方括号中的元素记帐表时无
    '           |-cllBillLists-单据信息集=cllPatiBillItem(BillLists)
    '               |-cllBillItem-单据信息，成员(费用来源,NO,收费标志,划价人,开单科室ID,开单科室名称,
    '                                                          开单医师ID,开单医师,操作员姓名,操作员编号,登记时间,DetailList)=cllBillLists(_费用来源_单据类型_单据号)
    '                   |-cllDetailList-单据明细集=cllBillItem(DetailList)
    '                       |-cllDetailItem-每行明细数据集，成员([病人ID,主页ID,姓名,性别编号,性别,年龄,病人病区ID,病人科室ID],
    '                                 费用ID,序号,库房ID,是否备货卫材,批次,卫材ID,婴儿序号,医嘱ID,付数,数量,售价,零售金额,摘要)；其中，方括号中的元素记帐表时才有
    Set cllExseErrData = New Collection
    For p = 1 To cllOutList.Count
        '  --    pati_bill_list[]
        '  --       billtype                   N   1 单据类型: 1 -收费处方  ;2- 记帐单处方;3- 记帐表处方
        '  --       pati_source                N   1 病人来源:1-门诊;2-住院;4-体检
        '  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，无以下节点--------------------------------------
        '  --       pati_id                    N   1 病人ID
        '  --       pati_pageid                N   1 主页ID
        '  --       pati_name                  C   1 病人姓名
        '  --       pati_sex_code              C   1 性别编号（新门诊)
        '  --       pati_sex                   C   1 性别
        '  --       pati_age                   C   1 年龄
        '  --       pati_deptid                N   1 病人科室ID
        '  --       pati_wardarea_id           N     病人病区ID
        '  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，无以上节点-----------------------------------------
        '  --       bill_list[]                      更新数据列表[数组]
        Set cllBillLists = New Collection
        
        Set cllItem = cllOutList(p)
        Set cllPatiBillItem = New Collection
        byt单据类型 = Val(Nvl(cllItem("_billtype")))
        bln记帐表 = (byt单据类型 = 3)
        cllPatiBillItem.Add cllItem("_billtype"), "单据类型"
        cllPatiBillItem.Add cllItem("_pati_source"), "病人来源"
        If bln记帐表 = False Then
            cllPatiBillItem.Add cllItem("_pati_id"), "病人ID"
            cllPatiBillItem.Add cllItem("_pati_pageid"), "主页ID"
            cllPatiBillItem.Add cllItem("_pati_name"), "姓名"
            cllPatiBillItem.Add cllItem("_pati_sex_code"), "性别编号"
            cllPatiBillItem.Add cllItem("_pati_sex"), "性别"
            cllPatiBillItem.Add cllItem("_pati_age"), "年龄"
            cllPatiBillItem.Add cllItem("_pati_deptid"), "病人科室ID"
            cllPatiBillItem.Add cllItem("_pati_wardarea_id"), "病人病区ID"
        End If
        cllPatiBillItem.Add cllBillLists, "BillLists"
        cllExseErrData.Add cllPatiBillItem
        
        Set cllBill_Out = mobjServiceCall.GetJsonListValue("output.pati_bill_list[" & p - 1 & "].bill_list")
        For i = 1 To cllBill_Out.Count
            '  --       bill_list[]                      更新数据列表[数组]
            '  --         fee_source                N  0 费用来源
            '  --         stuff_no                  C  1 NO
            '  --         charge_tag                N  1 收费标志:0-未收费或记帐划价;1-已收费或记帐
            '  --         fee_acnter                C  0 划价人
            '  --         plcdept_id                C  0 开单科室id（新门诊)
            '  --         plcdept                   C  0 开单科室名称（新门诊)
            '  --         placer_id                 C  0 开单医师id（新门诊)
            '  --         placer                    C  0 开单医师（新门诊) 增加
            '  --         operator_name             C  1 操作员姓名
            '  --         operator_code             C  1 操作员编号
            '  --         create_time               C  1 登记时间:yyyy-mm-dd hh:mi:ss
            '  --         item_list[]                    更新数据列表[数组]
            Set cllDetailList = New Collection
            
            Set cllItem = cllBill_Out(i)
            strKey = "_" & cllItem("_fee_source") & "_" & byt单据类型 & "_" & cllItem("_stuff_no")
            Set cllBillItem = New Collection
            cllBillItem.Add cllItem("_fee_source"), "费用来源"
            cllBillItem.Add cllItem("_stuff_no"), "NO"
            cllBillItem.Add cllItem("_charge_tag"), "收费标志"
            cllBillItem.Add cllItem("_fee_acnter"), "划价人"
            cllBillItem.Add cllItem("_plcdept_id"), "开单科室ID"
            cllBillItem.Add cllItem("_plcdept"), "开单科室名称"
            cllBillItem.Add cllItem("_placer_id"), "开单医师ID"
            cllBillItem.Add cllItem("_placer"), "开单医师"
            cllBillItem.Add cllItem("_operator_name"), "操作员姓名"
            cllBillItem.Add cllItem("_operator_code"), "操作员编号"
            cllBillItem.Add cllItem("_create_time"), "登记时间"
            cllBillItem.Add cllDetailList, "DetailList"
            cllBillLists.Add cllBillItem, strKey
            
            Set cllDetail_Out = mobjServiceCall.GetJsonListValue("output.pati_bill_list[" & p - 1 & "].bill_list[" & i - 1 & "].item_list")
            For j = 1 To cllDetail_Out.Count
                '  --         item_list[]                    更新数据列表[数组]
                '  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，有以下节点----------------------------------------
                '  --           pati_id                 N  1 病人ID
                '  --           pati_pageid             N  0 主页ID
                '  --           pati_name               C  1 病人姓名
                '  --           pati_sex_code           C  1 性别编号（新门诊)
                '  --           pati_sex                C  1 性别
                '  --           pati_age                C  1 年龄
                '  --           pati_wardarea_id        N    病人病区ID
                '  --           pati_deptid             N  1 病人科室ID
                '  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，有以上节点-----------------------------------------
                '  --           stuffdtl_id             N  1 处方明细ID(目前传入的是费用id)
                '  --           serial_num              N  1 序号:(变更(包括存储)：序号和组号，1、2、3、3、3、4…)
                '  --           warehouse_id            N  1 库房ID
                '  --           is_bakstuff             N  1 是否备货卫材:有高值卫材才需要传入，非0表示是高值卫材模式(如扫码时使用)
                '  --           bakstuff_batch          N  1 备货材料批次
                '  --           stuff_id                N  1 卫材ID
                '  --           baby_num                N  0 婴儿序号
                '  --           advice_id               N  0 医嘱ID
                '  --           packages_num            N  1 付数
                '  --           outbound_num            N  1 出库数量
                '  --           price                   N  0 售价
                '  --           money                   N  0 零售金额(新门诊)
                '  --           memo                    C  0 摘要
                Set cllItem = cllDetail_Out(j)
                Set cllDetailItem = New Collection
                If bln记帐表 Then
                    cllDetailItem.Add cllItem("_pati_id"), "病人ID"
                    cllDetailItem.Add cllItem("_pati_pageid"), "主页ID"
                    cllDetailItem.Add cllItem("_pati_name"), "姓名"
                    cllDetailItem.Add cllItem("_pati_sex_code"), "性别编号"
                    cllDetailItem.Add cllItem("_pati_sex"), "性别"
                    cllDetailItem.Add cllItem("_pati_age"), "年龄"
                    cllDetailItem.Add cllItem("_pati_wardarea_id"), "病人病区ID"
                    cllDetailItem.Add cllItem("_pati_deptid"), "病人科室ID"
                End If
                cllDetailItem.Add cllItem("_stuffdtl_id"), "费用ID"
                cllDetailItem.Add cllItem("_serial_num"), "序号"
                cllDetailItem.Add cllItem("_warehouse_id"), "库房ID"
                cllDetailItem.Add cllItem("_is_bakstuff"), "是否备货卫材"
                cllDetailItem.Add cllItem("_bakstuff_batch"), "批次"
                cllDetailItem.Add cllItem("_stuff_id"), "卫材ID"
                cllDetailItem.Add cllItem("_baby_num"), "婴儿序号"
                cllDetailItem.Add cllItem("_advice_id"), "医嘱ID"
                cllDetailItem.Add cllItem("_packages_num"), "付数"
                cllDetailItem.Add cllItem("_outbound_num"), "数量"
                cllDetailItem.Add cllItem("_price"), "售价"
                cllDetailItem.Add cllItem("_money"), "零售金额"
                cllDetailItem.Add cllItem("_memo"), "摘要"
                cllDetailList.Add cllDetailItem
            Next
        Next
    Next
 
    Exit Function
ErrHandler:
    GetExseSyncErrData = 2
    strErrMsg = err.Description
End Function

Private Function GetCisSyncErrData(ByVal strPatiIDs As String, ByRef cllCisErrData As Collection, ByRef strErrMsg As String) As Integer
    '获取临床域同步异常数据
    '入参：
    '   strPatiIDs=病人ID,多个用英文的逗号分隔
    '出参：
    '   cllCisErrData-临床域同步异常数据，说明：括号中的均为集合Key值
    '       |-cllOrderSendItem-病人医嘱发送记录，成员(病人ID,主页ID,挂号ID,挂号单号,发送号,OrderList)
    '           |-cllOrderList-医嘱信息列表=cllOrderSendItem(OrderList)
    '               |-cllOrderItem-医嘱信息，成员(医嘱ID,医嘱期效,紧急标志,计价特性)=cllOrderList(_医嘱ID)
    '           |-cllExseBillList-费用单据列表=cllOrderSendItem(ExseBillList)
    '               |-cllExseBillItem-费用单据信息，成员(费用来源,单据类型,单据号)=cllExseBillList(_费用来源_单据类型_单据号)
    '       其中，单据类型：1-收费单,2-记帐单,3-记帐表；费用来源：1-门诊,2-住院
    '   strErrMsg=范围值为2时，返回错误信息
    '返回：0-存在未同步的单据，1-不存在未同步的单据，2-其他错误
    Dim StrJson_In As String, strKey As String
    Dim i As Long, j As Long
    Dim cllOutList As Collection, cllOrder_Out As Collection
    Dim cllOrderSendItem As Collection, cllItem As Collection
    Dim cllOrderList As Collection, cllOrderItem As Collection
    Dim cllExseBillList As Collection, cllExseBillItem As Collection
    
    On Error GoTo ErrHandler
    Set cllCisErrData = New Collection
    strErrMsg = ""

    If strPatiIDs = "" Then GetCisSyncErrData = 1: Exit Function
    'Zl_Cissvr_Getstufferrdata
'  --功能：临床医嘱发送生成卫材数据同步
'  --入参：Json_In:格式
'  --  input
'  --      pati_ids                        C 1 病人ids逗号拼串
'  --出参: Json_Out,格式如下
'  --   output:
'  --     code: 1,
'  --     message: 成功,
'  --     data[]
'  --         pati_id                      N 1 病人id
'  --         pati_pageid                  N 0 主页id，住院病人传入，门诊传0
'  --         rgst_id                      N 0 挂号id，门诊病人传入，住院病人传空
'  --         rgst_no                      C 0 挂号单号
'  --         send_no                      N 1 发送号
'  --         order_list[]医嘱信息列表
'  --             advice_id                N 1 医嘱id
'  --             effectivetime            N 1 医嘱期效
'  --             emergency_tag            N 1 紧急标志
'  --             denominated              N 1 计价特性
'  --             fee_source               N 0 费用来源：1-门诊；2-住院
'  --             fee_billtype             N 0 费用单据类型：1-收费处方；2-记帐单处方
'  --             fee_no                   C 0 费用单据号
'  --             freq_name                C 0 频次名称
'  --             single                   N 0 单量
'  -----------------------------------------------------------------------------------
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_ids", strPatiIDs, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
 
    If mobjServiceCall.CallService("Zl_Cissvr_Getstufferrdata", StrJson_In, , , , False, , , , True) = False Then
        strErrMsg = "调用医嘱服务查询未产生单据失败！"
        GetCisSyncErrData = 2: Exit Function
    End If
    
    Set cllOutList = mobjServiceCall.GetJsonListValue("output.pati_bill_list")
    If cllOutList Is Nothing Then GetCisSyncErrData = 1: Exit Function
    If cllOutList.Count = 0 Then GetCisSyncErrData = 1: Exit Function
        
    '   cllCisErrData-临床域同步异常数据，说明：括号中的均为集合Key值
    '       |-cllOrderSendItem-病人医嘱发送记录，成员(病人ID,主页ID,挂号ID,挂号单号,发送号,OrderList)
    '           |-cllOrderList-医嘱信息列表=cllOrderSendItem(OrderList)
    '               |-cllOrderItem-医嘱信息，成员(医嘱ID,医嘱期效,紧急标志,计价特性)=cllOrderList(_医嘱ID)
    '           |-cllExseBillList-费用单据列表=cllOrderSendItem(ExseBillList)
    '               |-cllExseBillItem-费用单据信息，成员(费用来源,单据类型,单据号)=cllExseBillList(_费用来源_单据类型_单据号)
    Set cllCisErrData = New Collection
    For i = 1 To cllOutList.Count
        '  --     pati_bill_list[]
        '  --         pati_id                      N 1 病人id
        '  --         pati_pageid                  N 0 主页id，住院病人传入，门诊传0
        '  --         rgst_id                      N 0 挂号id，门诊病人传入，住院病人传空
        '  --         rgst_no                      C 0 挂号单号
        '  --         send_no                      N 1 发送号
        '  --         order_list[]医嘱信息列表
        Set cllOrderList = New Collection
        Set cllExseBillList = New Collection
        
        Set cllItem = cllOutList(i)
        Set cllOrderSendItem = New Collection
        cllOrderSendItem.Add cllItem("_pati_id"), "病人ID"
        cllOrderSendItem.Add cllItem("_pati_pageid"), "主页ID"
        cllOrderSendItem.Add cllItem("_rgst_id"), "挂号ID"
        cllOrderSendItem.Add cllItem("_rgst_no"), "挂号单号"
        cllOrderSendItem.Add cllItem("_send_no"), "发送号"
        cllOrderSendItem.Add cllOrderList, "OrderList"
        cllOrderSendItem.Add cllExseBillList, "ExseBillList"
        cllCisErrData.Add cllOrderSendItem
        
        Set cllOrder_Out = mobjServiceCall.GetJsonListValue("output.pati_bill_list[" & i - 1 & "].order_list")
        For j = 1 To cllOrder_Out.Count
            '  --         order_list[]医嘱信息列表
            '  --             advice_id                N 1 医嘱id
            '  --             effectivetime            N 1 医嘱期效
            '  --             emergency_tag            N 1 紧急标志
            '  --             denominated              N 1 计价特性
            '  --             fee_source               N 0 费用来源：1-门诊；2-住院
            '  --             fee_billtype             N 0 费用单据类型：1-收费处方；2-记帐单处方
            '  --             fee_no                   C 0 费用单据号
            '  --             freq_name                C 0 频次名称
            '  --             single                   N 0 单量
            Set cllItem = cllOrder_Out(j)
             
            '加入医嘱信息列表，相同的只加一次
            strKey = "_" & cllItem("_advice_id")
            If ExistsColObject(cllOrderList, strKey) = False Then
                Set cllOrderItem = New Collection
                cllOrderItem.Add cllItem("_advice_id"), "医嘱ID"
                cllOrderItem.Add cllItem("_effectivetime"), "医嘱期效"
                cllOrderItem.Add cllItem("_emergency_tag"), "紧急标志"
                cllOrderItem.Add cllItem("_denominated"), "计价特性"
                cllOrderItem.Add cllItem("_freq_name"), "频次名称"
                cllOrderItem.Add cllItem("_single"), "单量"
                cllOrderList.Add cllOrderItem, strKey
            End If
            
            '加入费用单据信息列表，相同的只加一次
            strKey = "_" & cllItem("_fee_source") & "_" & cllItem("_fee_billtype") & "_" & cllItem("_fee_no")
            If ExistsColObject(cllExseBillList, strKey) = False Then
                Set cllExseBillItem = New Collection
                '  --             fee_source               N 0 费用来源：1-门诊；2-住院
                '  --             fee_billtype             N 0 费用单据类型：1-收费处方；2-记帐单处方
                '  --             fee_no                   C 0 费用单据号
                cllExseBillItem.Add cllItem("_fee_source"), "费用来源"
                cllExseBillItem.Add cllItem("_fee_billtype"), "单据类型"
                cllExseBillItem.Add cllItem("_fee_no"), "单据号"
                cllExseBillList.Add cllExseBillItem, strKey
            End If
        Next
    Next
    
    Exit Function
ErrHandler:
    GetCisSyncErrData = 2
    strErrMsg = err.Description
End Function

Private Sub txtPatiId_Change()
    txtPatiId.Tag = ""
End Sub
