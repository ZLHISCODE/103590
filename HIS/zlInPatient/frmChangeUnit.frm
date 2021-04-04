VERSION 5.00
Begin VB.Form frmChangeUnit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人转病区"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmChangeUnit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   1860
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1395
         Width           =   3810
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   225
         Width           =   1170
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   225
         Width           =   675
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4095
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   225
         Width           =   690
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   615
         Width           =   1170
      End
      Begin VB.TextBox txt床号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   615
         Width           =   1605
      End
      Begin VB.TextBox txt病区 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1005
         Width           =   3795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "住院号"
         Height          =   180
         Left            =   360
         TabIndex        =   17
         Top             =   675
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转入病区"
         Height          =   180
         Left            =   180
         TabIndex        =   16
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   540
         TabIndex        =   15
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   2370
         TabIndex        =   14
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   3690
         TabIndex        =   13
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前床位"
         Height          =   180
         Left            =   2370
         TabIndex        =   12
         Top             =   675
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前病区"
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   1065
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2715
      TabIndex        =   1
      Top             =   2115
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3915
      TabIndex        =   2
      Top             =   2115
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   270
      TabIndex        =   3
      Top             =   2115
      Width           =   1100
   End
End
Attribute VB_Name = "frmChangeUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrPrivs As String
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlngUnit As Long
Private mrsPatiInfo As ADODB.Recordset
Private mfrmParent As Object

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Function InitData() As Boolean
    Dim i As Integer
    Dim rsTmp As ADODB.Recordset
    Dim str床号 As String, str房间号 As String
    
    On Error GoTo errHandle
    
    Set mrsPatiInfo = GetPatiInfo(mlng病人ID, mlng主页ID)
    
    With mrsPatiInfo
        txt姓名.Text = !姓名
        txt性别.Text = "" & !性别
        txt年龄.Text = "" & !年龄
        txt住院号.Text = "" & !住院号
    End With
    
    str房间号 = ""
    Set rsTmp = GetPatiBeds(mlng病人ID)
    If rsTmp.RecordCount = 0 Then
        str床号 = "家庭病床"
    Else
        Do While Not rsTmp.EOF
            str床号 = str床号 & "," & rsTmp!床号
            If Nvl(rsTmp!床号) = Nvl(mrsPatiInfo!主要床号) And Nvl(rsTmp!科室ID) = Nvl(mrsPatiInfo!入住科室id) Then
                str房间号 = Nvl(rsTmp!房间号)
            End If
            rsTmp.MoveNext
        Loop
        str床号 = Mid(str床号, 2)
    End If
    txt床号.Text = str床号
    txt床号.Tag = str房间号
    
    '目前包含门诊观察室
    
    gstrSQL = "Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,病区科室对应 B,部门性质说明 C " & _
            " Where B.病区ID=A.ID And B.科室ID=[1] " & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " And C.部门ID=A.ID And Instr(',' || [2]|| ',',',' || C.服务对象 || ',')>0 " & _
            " And C.工作性质='护理' " & _
            " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlInPatient", Val("" & mrsPatiInfo!出院科室id), "1,2,3")
    'Set rsTmp = GetDeptOrUnit(1, mrsPatiInfo!出院科室ID, "1,2,3")
    If Not rsTmp.EOF Then
        cboUnit.Clear
        For i = 1 To rsTmp.RecordCount
            
            If rsTmp!ID = mlngUnit Then
                txt病区.Text = rsTmp!名称
            Else
                
                cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
                cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            End If
            rsTmp.MoveNext
        Next
        If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0  '不调用InitBed设置床位
    End If
    
    If cboUnit.ListCount = 0 Then
        MsgBox "该病人所在科室没有设置其他对应的病区！", vbInformation, gstrSysName
        Exit Function
    End If
    
    InitData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function isValid() As Boolean
    
    Dim strSql As String
    Dim strInfo As String
    Dim rsTemp As New ADODB.Recordset

    If gbyt转科时检查未执行 <> 0 Then

        strInfo = ExistWaitExe(mlng病人ID, mlng主页ID)
        If strInfo <> "" Then
            If gbyt转科时检查未执行 = 1 Then
                If MsgBox("该病人存在尚未执行完成的内容：" & _
                    vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "确定要转病区吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                MsgBox "该病人存在尚未执行完成的内容：" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "不允许转病区。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If

    If gbyt转科时检查药品未执行 <> 0 Then
        strInfo = ExistWaitDrug(mlng病人ID, mlng主页ID)
        If strInfo <> "" Then
            If gbyt转科时检查药品未执行 = 1 Then
                If MsgBox("该病人" & strInfo & vbCrLf & vbCrLf & "确定要转病区吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                MsgBox "该病人" & strInfo & vbCrLf & vbCrLf & "不允许转病区。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '61429:刘鹏飞,2013-11-11,转科时销帐未审核单据检查
    If gbyt转科时未审核销帐单据检查 <> 0 Then
        strInfo = ""
        strInfo = ExistWaitQuittance(mlng病人ID, mlng主页ID)
        If strInfo <> "" Then
            If gbyt转科时未审核销帐单据检查 = 1 Then
                If MsgBox("该病人" & strInfo & vbCrLf & vbCrLf & "确定要转病区吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                MsgBox "该病人" & strInfo & vbCrLf & vbCrLf & "不允许转病区。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    isValid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    isValid = False
End Function

Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
        ByVal strPrivs As String) As Boolean
'#########################################################################################################
'### 参数：
'### 返回：目标床号
'#########################################################################################################
    On Error Resume Next
    Set mfrmParent = frmParent
    mlngUnit = lngUnit
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstrPrivs = strPrivs
    
    If InitData = False Then Exit Function
    
    Me.Show 1, frmParent
    
    ShowMe = gblnOK
End Function


Private Sub cmdOK_Click()
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQLtmp As String, rsPati As Recordset
    
    On Error GoTo errH
    
    If isValid = False Then Exit Sub
    
    '转病区费用检查
    If CreatePublicExpenseBillOperation() And gbln转病区转费用 Then
        strSQLtmp = "Select ID, 病区id" & vbNewLine & _
                    "From 病人变动记录" & vbNewLine & _
                    "Where 病人id = [1] And 主页id = [2] And 开始时间 Is Not Null And 终止时间 Is Null And NVL(附加床位,0) = 0"
        Set rsPati = zlDatabase.OpenSQLRecord(strSQLtmp, Me.Caption, mlng病人ID, mlng主页ID)
        If rsPati.RecordCount > 0 Then
            If gobjPublicExpenseBillOperation.zlTurnToWard_Fee_Query(Me, 2, mlng病人ID, mlng主页ID, Val(rsPati!ID & ""), Val(rsPati!病区ID & ""), cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
        End If
    End If
    
    strSql = "zl_病人变动记录_ChangeUnit(" & mlng病人ID & "," & mlng主页ID & "," & _
        cboUnit.ItemData(cboUnit.ListIndex) & ",'" & UserInfo.编号 & "'," & "'" & UserInfo.姓名 & "')"
        
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    '新网96847、118004
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(2, mlng病人ID, mlng主页ID) <> 1 Then
            MsgBox "当前启用了影像信息系统接口，但由于影像信息系统接口(HISModPati)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "当前启用了影像信息系统接口，但由于RIS接口创建失败未调用(HISModPati)接口，请与系统管理员联系。", vbInformation, gstrSysName
    End If
    gblnOK = True
    
    On Error Resume Next
    '转病区成功后触发消息
    If mclsMipModule.IsConnect = True Then
         mclsXML.ClearXmlText '清除缓存中的XML
        '--进行消息组装
        '病人信息
        mclsXML.AppendNode "in_patient"
        'patient_id      病人id  1   N
        mclsXML.appendData "patient_id", mlng病人ID, xsNumber  '病人ID
        'page_id     主页id  1   N
        mclsXML.appendData "page_id", mlng主页ID, xsNumber '主页ID
        'patient_name        姓名    1   S
        mclsXML.appendData "patient_name", txt姓名.Text, xsString '姓名
        'patient_sex     性别    0..1    S
        mclsXML.appendData "patient_sex", txt性别.Text, xsString '性别
        'in_number       住院号  1   S
        mclsXML.appendData "in_number", Nvl(mrsPatiInfo!住院号), xsString '住院号
        mclsXML.AppendNode "in_patient", True
        
        '转出信息
        'current_state       转出信息    1
        mclsXML.AppendNode "current_state"
        'current_area_id     转出病区id  0..1    N
        mclsXML.appendData "current_area_id", Val(Nvl(mrsPatiInfo!当前病区ID)), xsNumber
        'current_area_title      转出病区    0..1    S
        mclsXML.appendData "current_area_title", Nvl(mrsPatiInfo!当前病区), xsString
        'current_dept_id     转出科室id  1   N
        mclsXML.appendData "current_dept_id", Val(Nvl(mrsPatiInfo!出院科室id, 0)), xsNumber
        'current_dept_title      转出科室    1   S
        mclsXML.appendData "current_dept_title", Nvl(mrsPatiInfo!当前科室), xsString
        'current_room        转出病房    0..1    S
        mclsXML.appendData "current_room", txt床号.Tag, xsString
        'current_bed     转出病床    1   S
        mclsXML.appendData "current_bed", Nvl(mrsPatiInfo!主要床号), xsString
        mclsXML.AppendNode "current_state", True
        
        strSql = " Select ID 变动ID,sysdate 变动时间 From 病人变动记录  Where 病人ID=[1] And 主页Id=[2] And 开始原因=[3] and 开始时间 IS NULL"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病人变动记录", mlng病人ID, mlng主页ID, 15)
        '转入信息
        'change_state        转入信息    1
        mclsXML.AppendNode "change_state"
        'change_id       转科变更id  1   N
        mclsXML.appendData "change_id", rsTmp!变动ID, xsNumber
        'change_date     变更时间    1   S
        mclsXML.appendData "change_date", Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss"), xsString
        
        'change_area_id      转入病区id  0..1    N
        mclsXML.appendData "change_area_id", Val(cboUnit.ItemData(cboUnit.ListIndex)), xsNumber
        'change_area_title       转入病区    0..1    S
        mclsXML.appendData "change_area_title", zlCommFun.GetNeedName(cboUnit.Text), xsString
        'change_dept_id      转入科室id  0..1    N
        mclsXML.appendData "change_dept_id", Val(Nvl(mrsPatiInfo!出院科室id, 0)), xsNumber
        'change_dept_title       转入科室    0..1    S
        mclsXML.appendData "change_dept_title", Nvl(mrsPatiInfo!当前科室), xsString
        mclsXML.AppendNode "change_state", True

        mclsMipModule.CommitMessage "ZLHIS_PATIENT_003", mclsXML.XmlText
    End If
    If Err <> 0 Then Err.Clear
    
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    '创建消息对象
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '卸载消息对象
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
End Sub
