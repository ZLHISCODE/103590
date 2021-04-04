VERSION 5.00
Begin VB.Form frmChange 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人转科"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "frmChange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   285
      TabIndex        =   9
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3810
      TabIndex        =   8
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2610
      TabIndex        =   7
      Top             =   2040
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1860
      Left            =   135
      TabIndex        =   10
      Top             =   45
      Width           =   5055
      Begin VB.TextBox txt科室 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1005
         Width           =   3795
      End
      Begin VB.TextBox txt床号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   615
         Width           =   1605
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   615
         Width           =   1170
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4095
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   225
         Width           =   690
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   675
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   225
         Width           =   1170
      End
      Begin VB.ComboBox cbo科室 
         Height          =   300
         Left            =   990
         TabIndex        =   6
         Text            =   "cbo科室"
         Top             =   1395
         Width           =   3810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前科室"
         Height          =   180
         Left            =   180
         TabIndex        =   17
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前床位"
         Height          =   180
         Left            =   2370
         TabIndex        =   16
         Top             =   675
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   3690
         TabIndex        =   14
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   2370
         TabIndex        =   13
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   540
         TabIndex        =   12
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转入科室"
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "住院号"
         Height          =   180
         Left            =   360
         TabIndex        =   15
         Top             =   675
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mlng病人ID As Long
Public mlng主页ID As Long
Public mlngUnit As Long
Public mstrPrivs As String
Private mstr服务对象 As String
Private mintFlag As Integer
Private mstrDeptName As String
Private mrsPatiInfo As ADODB.Recordset
Private mfrmParent As Object

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cbo科室_GotFocus()
    '问题27370 by lesfeng 2010-02-03
    With cbo科室
        .SelStart = 0
        .SelLength = Len(.Text)
        mstrDeptName = .Text
    End With
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    '问题27370 by lesfeng 2010-02-03
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String
    Dim strSql As String, intIdx As Long, i As Long
    Dim lngUnit As Long
    
    If KeyAscii = 13 Then
        mintFlag = 0
        strInput = UCase(cbo科室.Text)
  
        Set rsTmp = InputDept(Me, Frame1, cbo科室, "临床", mstr服务对象, strInput, blnCancel, -1, 0)
        If Not rsTmp Is Nothing Then
            intIdx = cbo.FindIndex(cbo科室, rsTmp!ID)
            If intIdx <> -1 Then
                cbo科室.ListIndex = intIdx
            End If
        Else
            If Not blnCancel Then
                MsgBox "未找到对应的科室。", vbInformation, gstrSysName
                cbo科室.Text = mstrDeptName
                mintFlag = 1
            End If
        End If
    Else
        mintFlag = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '问题27370 by lesfeng 2010-02-03
    If KeyCode = 13 And mintFlag = 0 Then cmdOK_Click
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    If isValid = False Then Exit Sub
    
    strSql = "zl_病人变动记录_Change(" & mlng病人ID & "," & mlng主页ID & "," & _
        cbo科室.ItemData(cbo科室.ListIndex) & ",'" & UserInfo.编号 & "'," & "'" & UserInfo.姓名 & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    '新网96847
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(2, mlng病人ID, mlng主页ID) <> 1 Then
            MsgBox "当前启用了影像信息系统接口，但由于影像信息系统接口(HISModPati)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "当前启用了影像信息系统接口，但由于RIS接口创建失败未调用(HISModPati)接口，请与系统管理员联系。", vbInformation, gstrSysName
    End If
    gblnOK = True
    
    On Error Resume Next
    '转科成功后触发消息
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病人变动记录", mlng病人ID, mlng主页ID, 3)
        '转入信息
        'change_state        转入信息    1
        mclsXML.AppendNode "change_state"
        'change_id       转科变更id  1   N
        mclsXML.appendData "change_id", rsTmp!变动ID, xsNumber
        'change_date     变更时间    1   S
        mclsXML.appendData "change_date", Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss"), xsString
        'change_area_id      转入病区id  0..1    N
        'mclsXML.appendData "change_area_id", Val(Nvl(mrsPatiInfo!当前病区id)), xsNumber
        'change_area_title       转入病区    0..1    S
        'mclsXML.appendData "change_area_title", Nvl(mrsPatiInfo!当前病区), xsString
        'change_dept_id      转入科室id  0..1    N
        mclsXML.appendData "change_dept_id", Val(cbo科室.ItemData(cbo科室.ListIndex)), xsNumber
        'change_dept_title       转入科室    0..1    S
        mclsXML.appendData "change_dept_title", zlCommFun.GetNeedName(cbo科室.Text), xsString
        mclsXML.AppendNode "change_state", True

        mclsMipModule.CommitMessage "ZLHIS_PATIENT_003", mclsXML.XmlText
    End If
    If Err <> 0 Then Err.Clear
    
    '调用外挂接口
    If CreatePlugInOK(glngModul) Then
        On Error Resume Next
        Call gobjPlugIn.InPatiCheckInBranchAfter(mlng病人ID, mlng主页ID)
        Call zlPlugInErrH(Err, "InPatiCheckInBranchAfter")
        Err.Clear: On Error GoTo 0
    End If
    
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function isValid() As Boolean
    
    Dim strSql As String
    Dim strInfo As String
    Dim lng科室ID As Long, lng转入科室ID As Long
    Dim blnSameUnit As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    lng科室ID = mrsPatiInfo!入住科室id
    lng转入科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    
    If gbyt转科时检查未执行 <> 0 Then
        '同一病区之间转科,提示但不禁止
        strSql = "Select Distinct (A.病区id) 病区id " & _
                 "From 病区科室对应 A, 病区科室对应 B " & _
                 "Where A.病区id = B.病区id And A.科室id = [1] And B.科室id = [2]"
        
        On Error GoTo errHandle
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng科室ID, lng转入科室ID)
        If rsTemp.RecordCount > 0 Then blnSameUnit = True
            
        strInfo = ExistWaitExe(mlng病人ID, mlng主页ID)
        If strInfo <> "" Then
            If gbyt转科时检查未执行 = 1 Or blnSameUnit = True Then
                If MsgBox("该病人存在尚未执行完成的内容：" & _
                    vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "确定要转科吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                MsgBox "该病人存在尚未执行完成的内容：" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "不允许转科.", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '问题30208 by lesfeng 2010-08-02 撤分参数22及32 新增154、155
    If gbyt转科时检查药品未执行 <> 0 Then
        '同一病区之间转科,提示但不禁止
        strSql = "Select Distinct (A.病区id) 病区id " & _
                 "From 病区科室对应 A, 病区科室对应 B " & _
                 "Where A.病区id = B.病区id And A.科室id = [1] And B.科室id = [2]"
        
        On Error GoTo errHandle
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng科室ID, lng转入科室ID)
        If rsTemp.RecordCount > 0 Then blnSameUnit = True
        
        strInfo = ExistWaitDrug(mlng病人ID, mlng主页ID)
        If strInfo <> "" Then
            If gbyt转科时检查药品未执行 = 1 Or blnSameUnit = True Then
                If MsgBox("该病人" & strInfo & vbCrLf & vbCrLf & "确定要转科吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                MsgBox "该病人" & strInfo & vbCrLf & vbCrLf & "不允许转科。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    
    '61429:刘鹏飞,2013-11-11,转科时销帐未审核单据检查
    If gbyt转科时未审核销帐单据检查 <> 0 Then
        '同一病区之间转科,提示但不禁止
        blnSameUnit = False
        strSql = "Select Distinct (A.病区id) 病区id " & _
                 "From 病区科室对应 A, 病区科室对应 B " & _
                 "Where A.病区id = B.病区id And A.科室id = [1] And B.科室id = [2]"
        
        On Error GoTo errHandle
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng科室ID, lng转入科室ID)
        If rsTemp.RecordCount > 0 Then blnSameUnit = True
        
        strInfo = ExistWaitQuittance(mlng病人ID, mlng主页ID)
        If strInfo <> "" Then
            If gbyt转科时未审核销帐单据检查 = 1 Or blnSameUnit = True Then
                If MsgBox("该病人" & strInfo & vbCrLf & vbCrLf & "确定要转科吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                MsgBox "该病人" & strInfo & vbCrLf & vbCrLf & "不允许转科。", vbInformation, gstrSysName
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

Private Function LoadBed() As Boolean
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Integer, lng科室ID As Long
    Dim byt病人性质 As Byte
    Dim strTmp As String, str床号 As String, str房间号 As String

    '问题27370 by lesfeng 2010-02-03
    mintFlag = 0
    
    On Error GoTo errH
    
    gblnOK = False
    Set mrsPatiInfo = GetPatiInfo(mlng病人ID, mlng主页ID)
    '住院科室
    With mrsPatiInfo
        txt姓名.Text = !姓名
        txt性别.Text = "" & !性别
        txt年龄.Text = "" & !年龄
        txt住院号.Text = "" & !住院号
        txt科室.Text = !当前科室
        
        lng科室ID = !入住科室id
        byt病人性质 = Val("" & !病人性质)
    End With
    
    Set rsTmp = GetPatiBeds(mlng病人ID)
    If rsTmp.RecordCount = 0 Then
        str床号 = "家庭病床"
        str房间号 = ""
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
        
    '确定病区的服务对象
    strSql = "Select 服务对象 From 部门性质说明 Where 工作性质='护理' And 部门ID=[1]"
     Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnit)
    
    If rsTmp!服务对象 = 1 Then
        strTmp = "1,3"
    ElseIf rsTmp!服务对象 = 2 Then
        strTmp = "2,3"
    ElseIf rsTmp!服务对象 = 3 Then
        If byt病人性质 = 1 Then
            strTmp = "1,3"
        Else
            strTmp = "2,3"
        End If
    End If
    '问题27370 by lesfeng 2010-02-03
    mstr服务对象 = strTmp
    
    '可选科室为临床科室,没有床位的也列出,因为可能使用病区的共用床
    Set rsTmp = GetDepts("临床", strTmp)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp!ID <> lng科室ID Then
                cbo科室.AddItem rsTmp!编码 & "-" & rsTmp!名称
                cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!ID
            End If
            rsTmp.MoveNext
        Next
        cbo科室.ListIndex = 0
    Else
        MsgBox "没有找到与当前病区服务对象相同的临床科室,请到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    LoadBed = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowMe(frmMain As Object, ByVal lngUnit As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String) As Boolean
    
    
    Set mfrmParent = frmMain
    mlngUnit = lngUnit
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstrPrivs = strPrivs
    If LoadBed = False Then Exit Function
    Me.Show 1, frmMain
    
    ShowMe = gblnOK
End Function

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
