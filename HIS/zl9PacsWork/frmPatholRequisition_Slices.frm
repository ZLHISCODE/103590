VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholRequisition_Slices 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "制片申请"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6075
   Icon            =   "frmPatholRequisition_Slices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5915.493
   ScaleMode       =   0  'User
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdSure 
      Cancel          =   -1  'True
      Caption         =   "确定(&S)"
      Height          =   400
      Left            =   3120
      TabIndex        =   9
      Top             =   4748
      Width           =   900
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   400
      Left            =   5040
      TabIndex        =   8
      Top             =   4748
      Width           =   900
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&D)"
      Height          =   400
      Left            =   4080
      TabIndex        =   7
      Top             =   4748
      Width           =   900
   End
   Begin VB.PictureBox picShow 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   3135
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox txtShow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   2655
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.TextBox txtDescription 
      Height          =   975
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3600
      Width           =   4755
   End
   Begin VB.TextBox txtRequestDoctor 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3120
      Width           =   1515
   End
   Begin MSComCtl2.DTPicker dtpRequestTime 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   3120
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67633155
      CurrentDate     =   40646.4399652778
   End
   Begin zl9PACSWork.ucFlexGrid ufgData 
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   5106
      GridRows        =   51
      IsCopyAdoMode   =   0   'False
      IsEjectConfig   =   -1  'True
      HeadFontCharset =   134
      HeadFontWeight  =   400
      DataFontCharset =   134
      DataFontWeight  =   400
      ExtendLastCol   =   -1  'True
   End
   Begin VB.Label labRequestTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "申请时间："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   3180
      Width           =   900
   End
   Begin VB.Label labRequestDoctor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "申请医师："
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3480
      TabIndex        =   4
      Top             =   3180
      Width           =   900
   End
   Begin VB.Label labDescription 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "申请描述："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   900
   End
End
Attribute VB_Name = "frmPatholRequisition_Slices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mufgParentContextGrid As ucFlexGrid
Private mufgParentRequestGrid As ucFlexGrid

Private mlngCurRequestId As Long
Private mlngPatholAdviceId As Long
Private mlngRequestType As Long


Private mfrmOwner As Form

Private Const M_STR_REQUSLICES_COLS = "|材块ID,hide,uncfg|材块号,hide,uncfg|标本名称,cbx<>,w2600,uncfg|制片方式,cbx<1-重切,2-深切,3-连切,4-白片,5-重染,6-薄片>,uncfg,w1300|制片数量,w1200,uncfg|"
Private Const M_STR_REQUSLICES_CONVERTFORMAT = "制片方式:1-重切,2-深切,3-连切,4-白片,5-重染,6-薄片"

Public blnIsOk As Boolean


Public Function ShowSlicesRequestWindow(ufgParentRequestGrid As ucFlexGrid, ufgParentContextGrid As ucFlexGrid, _
    ByVal lngPatholAdviceId As Long, ByVal lngRequestId As Long, ByVal lngRequestType As Long, owner As Form) As Boolean
'显示制片申请窗口
    Set mufgParentRequestGrid = ufgParentRequestGrid
    Set mufgParentContextGrid = ufgParentContextGrid
    
    Set mfrmOwner = owner

    mlngPatholAdviceId = lngPatholAdviceId
    mlngCurRequestId = lngRequestId
    mlngRequestType = lngRequestType
    blnIsOk = False

    Call Me.Show(1, owner)
End Function

Private Sub InitRequisitionSlicesList()
On Error GoTo errHandle
'初始化申请制片列表
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    
    ufgData.IsKeepRows = True
    
    '禁止右键弹出列表配置窗口
    ufgData.IsEjectConfig = False
     ufgData.GridRows = 31
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    '设置列
    ufgData.DefaultColNames = M_STR_REQUSLICES_COLS
    ufgData.ColNames = M_STR_REQUSLICES_COLS
    ufgData.ColConvertFormat = M_STR_REQUSLICES_CONVERTFORMAT
    
   
    
    strSql = "select 材块ID,序号 as 材块号 ,标本名称,'-'||取材位置 取材位置 from 病理取材信息 where 病理医嘱ID=[1] and 确认状态=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "得到标本名称", mlngPatholAdviceId)
    
    If rsTemp.RecordCount < 1 Then Exit Sub
    For i = 1 To rsTemp.RecordCount
    
        strTemp = strTemp & "|" & "#" & Nvl(rsTemp!材块id) & "!" & Nvl(rsTemp!材块号) & ";" & Nvl(rsTemp!材块号) & "-" & Nvl(rsTemp!标本名称) & IIf(Len(Nvl(rsTemp!取材位置)) <> 1, Nvl(rsTemp!取材位置), "")
        
        rsTemp.MoveNext
    Next i
    
    ufgData.ComboxListFormat(ufgData.GetColIndex("标本名称")) = Mid(strTemp, 2, Len(strTemp))
    
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub LoadMaterialInf()
'载入材块信息
    Dim strSql As String

    strSql = "select 材块ID,序号 as 材块号,标本名称,'0' as 制片方式,1 as 制片数量 from 病理取材信息 where 病理医嘱ID=[1] and 确认状态=1"

    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngPatholAdviceId)

    ufgData.RefreshData
End Sub


Private Sub SaveSlicesRequest()
'保存制片申请
On Error GoTo errHandle
    Dim lngNewRow As Long
    Dim i As Integer
    Dim lngSlicesType As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset


    If mlngCurRequestId <= 0 Then
        '添加检查申请信息
        strSql = "select Zl_病理申请_新增([1],[2],[3],[4],[5],[6],[7]) as 返回值 from dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                mlngPatholAdviceId, _
                                                txtRequestDoctor.Text, _
                                                CDate(dtpRequestTime.value), _
                                                3, 0, _
                                                0, _
                                                txtDescription.Text)

        If rsData.RecordCount <= 0 Then
            Call err.Raise(0, "SaveSpeExamRequest", "未成功获取新增后的申请ID,处理失败。")
            Exit Sub
        End If

        '设置界面信息
        lngNewRow = mufgParentRequestGrid.NewRow

        mufgParentRequestGrid.Text(lngNewRow, gstrRequisition_申请ID) = rsData!返回值
        mufgParentRequestGrid.Text(lngNewRow, gstrRequisition_申请人) = txtRequestDoctor.Text
        mufgParentRequestGrid.Text(lngNewRow, gstrRequisition_申请类型) = "再制片"
        mufgParentRequestGrid.Text(lngNewRow, gstrRequisition_申请时间) = dtpRequestTime.value
        mufgParentRequestGrid.Text(lngNewRow, gstrRequisition_申请描述) = txtDescription.Text
        mufgParentRequestGrid.Text(lngNewRow, gstrRequisition_当前状态) = "已申请"

        mlngCurRequestId = Val(Nvl(rsData!返回值))

        '定位到新增行
        Call mufgParentRequestGrid.LocateRow(lngNewRow)

        Call mufgParentContextGrid.ClearListData
    End If


       For i = 1 To ufgData.GridRows - 1

        If Trim(ufgData.Text(i, "制片方式")) <> "" Then
            
            '添加制片申请项目
            lngSlicesType = GetSlicesTypeCode(Nvl(Mid(ufgData.Text(i, "标本名称"), 1, InStr(ufgData.Text(i, "标本名称"), "!") - 1)))
        
            strSql = "select Zl_病理申请_制片项目_新增([1],[2],[3],[4],[5],[6]) as 返回值 from dual"
            Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                    mlngPatholAdviceId, _
                                                    Nvl(Mid(ufgData.Text(i, "标本名称"), 1, InStr(ufgData.Text(i, "标本名称"), "!") - 1)), _
                                                    mlngCurRequestId, _
                                                    lngSlicesType, _
                                                    Val(Nvl(ufgData.Text(i, "制片方式"))), _
                                                    Val(Nvl(ufgData.Text(i, "制片数量"))))
        
            If rsData.RecordCount <= 0 Then
                Call err.Raise(0, "SaveSpeExamRequest", "未成功获取新增后的制片项目ID,处理失败。")
                Exit Sub
            End If
        
            '设置界面信息
            lngNewRow = mufgParentContextGrid.NewRow
        
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_ID) = rsData!返回值
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_标本名称) = Nvl(ufgData.DisplayText(i, "标本名称"))
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_材块号) = Nvl(Mid(ufgData.Text(i, "标本名称"), InStr(ufgData.Text(i, "标本名称"), "!") + 1, Len(ufgData.Text(i, "标本名称"))))
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_制片类型) = GetSlicesTypeValue(lngSlicesType)
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_制片方式) = Mid(Nvl(ufgData.Text(i, "制片方式")), 3, 4)
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_当前状态) = "已申请"
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_制片数量) = Val(Nvl(ufgData.Text(i, "制片数量")))
            
            Call mufgParentContextGrid.LocateRow(lngNewRow)
        End If
     Next i
     
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub cmdDel_Click()
'删除选中数据行
    ufgData.DelCurRow
End Sub

Private Sub cmdExit_Click()
'退出制片申请
    blnIsOk = False
    Call Me.Hide
End Sub


Private Function GetSlicesTypeValue(ByVal lngSlicesType As Long) As String
'获取制片类型取值
    Select Case lngSlicesType
        Case 0
            GetSlicesTypeValue = "石蜡制片"
        Case 1
            GetSlicesTypeValue = "冰冻切片"
        Case 2
            GetSlicesTypeValue = "细胞制片"
    End Select
End Function


Private Function GetSlicesTypeCode(ByVal strMaterialId As String) As Long
'取得制片类型代码
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    '如果是冰冻检查，则需要判断当前材块是否为冰余，如果为冰余，则制片类型为石蜡制片，否则为冰冻制片
    strSql = "select case 检查类型 when 1 then case 是否冰余 when 0 then 1 else 0 end when 2 then 2 else 0 end as 制片类型 " & _
            "from 病理检查信息 a, 病理取材信息 b where a.病理医嘱ID=b.病理医嘱ID and b.材块id=[1]"

    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strMaterialId)
    
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "GetSlicesTypeCode", "不能获取有效的制片类型。")
        Exit Function
    End If
    
    GetSlicesTypeCode = rsData!制片类型
End Function


Private Sub cmdSure_Click()
On Error GoTo errHandle

    Dim blnIsNumber As Boolean
    Dim blnIsRepetition As Boolean
    Dim blnIsContextRepetion As Boolean
    Dim lngSpecimenNameIndex As Long
    Dim lngSlicesTypeIndex As Long
    Dim lngCompletionStatus As Long
    Dim strOldSpecimenName As String
    Dim i As Integer
    

    For i = 1 To ufgData.GridRows - 1
        If Trim(ufgData.Text(i, "制片方式")) <> "" Then
        
            '判断制片数量 是否是数字  是否大于0  是否是小数
             If IsNumeric(ufgData.Text(i, "制片数量")) And Val(ufgData.Text(i, "制片数量")) > 0 And InStr(ufgData.Text(i, "制片数量"), ".") < 1 Then
                 blnIsNumber = True
             Else
                 blnIsNumber = False
                 Exit For
             End If
             
             
            '检查标本名称和制片方式 是否都重复
            If InStr(strOldSpecimenName, ufgData.DisplayText(i, "标本名称") & "," & ufgData.Text(i, "制片方式")) > 0 Then
                blnIsRepetition = True
                Exit For
            Else
                strOldSpecimenName = strOldSpecimenName & ufgData.DisplayText(i, "标本名称") & "," & ufgData.Text(i, "制片方式") & "|"
                blnIsRepetition = False
            End If
            
            If mlngRequestType = 3 Then
                 '检查标本名称和制片方式 在申请项目中是否存在 并判断是否已有 已完成的记录
                 lngSpecimenNameIndex = mufgParentContextGrid.FindRowIndex(ufgData.DisplayText(i, "标本名称"), "标本名称", True)
                 lngSlicesTypeIndex = mufgParentContextGrid.FindRowIndex(Mid(ufgData.Text(i, "制片方式"), InStr(ufgData.Text(i, "制片方式"), "-") + 1, 10), "制片方式", True)
                 lngCompletionStatus = mufgParentContextGrid.FindRowIndex("已完成", "当前状态", True)
                 
                If lngSpecimenNameIndex > 0 And lngSlicesTypeIndex > 0 And lngCompletionStatus < 1 Then
                    blnIsContextRepetion = True
                    Exit For
                Else
                    blnIsContextRepetion = False
                End If
            Else
                blnIsContextRepetion = False
            End If

        End If
    Next i
  
    If Not blnIsNumber Then
        Call ShowProcessHint("请输入有效的制片数量。")
        Exit Sub
    End If
    
    If blnIsRepetition Then
        Call ShowProcessHint("标本名称或制片方式重复。")
        Exit Sub
    End If
    
    If blnIsContextRepetion Then
        Call ShowProcessHint("申请项目中存在重复数据。")
        Exit Sub
    End If
    
    
    
    '判断申请明细列表是否为制片项目明细列表
    If mufgParentContextGrid.GetColIndex(gstrRequest_Slices_制片方式) < 0 Then
        mufgParentContextGrid.ColNames = gstrRequest_Slices_Cols
        mufgParentContextGrid.ColConvertFormat = gstrRequest_SlicesConvertFormat
        
        '切换主界面的控制界面
        Call mfrmOwner.ChangeControlFace(3)
    End If
    
    
    '保存制片申请
    Call SaveSlicesRequest
    
    blnIsOk = True
    
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle

    picShow.Visible = False
    txtRequestDoctor.Text = UserInfo.姓名
    dtpRequestTime.value = zlDatabase.Currentdate
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitRequisitionSlicesList
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ShowProcessHint(ByVal strHint As String)
'显示处理信息
On Error Resume Next

    txtShow.Text = strHint

    picShow.Visible = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If Trim(ufgData.Text(Row, "制片方式")) = "" Then ufgData.Text(Row, "制片方式") = "1-重切"
    If Trim(ufgData.Text(Row, "制片数量")) = "" Then ufgData.Text(Row, "制片数量") = "1"
    
End Sub
