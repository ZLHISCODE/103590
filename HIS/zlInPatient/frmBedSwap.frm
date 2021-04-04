VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmBedSwap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人床位对换"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBedSwap.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboNew 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   1845
   End
   Begin VB.Frame fraBedSwap 
      Height          =   1150
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   5565
      Begin VB.TextBox txtSwap科室 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   660
         Width           =   1800
      End
      Begin VB.TextBox txtSwap住院号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   270
         Width           =   1845
      End
      Begin VB.TextBox txtSwap姓名 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   270
         Width           =   1800
      End
      Begin VB.TextBox txtSwapPre 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label lblSwapDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   345
         TabIndex        =   24
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblSwapPName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         Height          =   180
         Left            =   345
         TabIndex        =   23
         Top             =   330
         Width           =   360
      End
      Begin VB.Label lblSwapInHosNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   2910
         TabIndex        =   22
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl床号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原病床"
         Height          =   180
         Left            =   2910
         TabIndex        =   21
         Top             =   720
         Width           =   540
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   5805
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3165
      Width           =   5805
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   345
         Left            =   3240
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   345
         Left            =   4440
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   345
         Left            =   240
         TabIndex        =   5
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.Frame fraBed 
      Height          =   1150
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   5565
      Begin VB.TextBox txtPre 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   660
         Width           =   1845
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         Width           =   1800
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         Width           =   1845
      End
      Begin VB.TextBox txt科室 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label lblPre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原病床"
         Height          =   180
         Left            =   2910
         TabIndex        =   13
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lblInHosNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   2910
         TabIndex        =   12
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lblPName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         Height          =   180
         Left            =   345
         TabIndex        =   11
         Top             =   330
         Width           =   360
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   345
         TabIndex        =   10
         Top             =   720
         Width           =   360
      End
   End
   Begin MSMask.MaskEdBox txtDate 
      Height          =   300
      Left            =   3630
      TabIndex        =   2
      Top             =   1440
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   19
      Format          =   "yyyy-MM-dd hh:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblNew 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "目标病床"
      Height          =   195
      Left            =   105
      TabIndex        =   25
      Top             =   1500
      Width           =   720
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "换床时间"
      Height          =   180
      Left            =   2850
      TabIndex        =   15
      Top             =   1500
      Width           =   720
   End
End
Attribute VB_Name = "frmBedSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng病人ID As Long              '当前病人ID
Private mlng主页ID As Long              '当前病人主页ID
Private mlngBeSwap病人ID As Long        '被换病人ID
Private mlngBeSwap主页ID As Long        '被换病人主页ID
Private mstr床号 As String              '当前病人床号
Private mstr目标床号 As String          '床位对换的目标床号（即被换病人床号）
Private mfrmParent As Object

Public mstrPrivs As String              '权限
Public mlngUnit As Long                 '病人病区ID

Private mrsPatiInfo As ADODB.Recordset
Private mrsSwapPatiInfo As ADODB.Recordset
Private mrsBeds As ADODB.Recordset '可选床位集

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cboNew_Click()
    Dim strBed As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle:
    
    If cboNew.ListIndex <> -1 And cboNew.ListCount > 0 Then
        '去床位
        If InStr(Trim(cboNew.Text), " 房间") > 0 Then
            strBed = Mid(Trim(cboNew.Text), 1, InStr(Trim(cboNew.Text), " 房间") - 1)
        Else
            strBed = Trim(cboNew.Text)
        End If
        '根据床号及 病区科室查找病人ID
        gstrSQL = "Select 病人ID From 床位状况记录 Where (科室ID is Null Or 科室ID=[1] Or 共用=1) And 病区ID=[2] And 床号=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsPatiInfo!出院科室id), mlngUnit, strBed)
        
        mlngBeSwap病人ID = rsTmp!病人ID
        mlngBeSwap主页ID = GetMax主页ID(rsTmp!病人ID) - 1
        
        '根据病人ID 和 主页ID 获取病人信息
        Set mrsSwapPatiInfo = GetPatiInfo(mlngBeSwap病人ID, mlngBeSwap主页ID)
        txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
        With mrsSwapPatiInfo
            txtSwap姓名.Text = !姓名
            txtSwap住院号.Text = "" & !住院号
            txtSwap科室.Text = !当前科室
        End With
        
        txtSwapPre.Text = cboNew.Text
        '目标床号修改为所选床号
        mstr目标床号 = Trim(cboNew.Text)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboNew_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cboNew.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cboNew.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cboNew.ListIndex = lngIdx
    ElseIf cboNew.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim strBeds As String, strSql As String, strMainBed As String
    Dim dMax As Date, dMaxSwap As Date, i As Integer, j As Integer, blnTrans As Boolean
    Dim strRoom As String, strOldRoom As String, Curdate As Date, strBedGrids As String, strBedGridsNew As String
    Dim rsTmp As ADODB.Recordset, intMainBed As Integer
    Dim arrSQL() As String, intLoop As Integer
    Dim strErr As String
    
    '时间不能超过当前时间太长(一个月)
    Curdate = zlDatabase.Currentdate
    If CDate(txtDate.Text) > Curdate Then
        If CDate(txtDate.Text) - Curdate > 30 Then
            MsgBox "换床时间比当前时间大得过多,请检查！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If MsgBox("换床时间大于了当前系统时间,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
        
    dMax = GetMaxDate(mlng病人ID, mlng主页ID)
    dMaxSwap = GetMaxDate(mlngBeSwap病人ID, mlngBeSwap主页ID)
    If CDate(txtDate.Text) <= dMax Then
        MsgBox "病人换床时间必须大于上次变动的时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    If CDate(txtDate.Text) <= dMaxSwap Then
        MsgBox "病人换床时间必须大于上次变动的时间 " & Format(dMaxSwap, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    
    If cboNew.ListIndex = -1 Then
        MsgBox "请选择要换入的床位！", vbInformation, gstrSysName
        cboNew.SetFocus: Exit Sub
    End If
        
    '将选中病人床位换到指定床位
    strMainBed = Trim(Split(cboNew.Text, "房间:")(0))

    '取床位
    '判断目标床位所在房间是否存在男女混住情况
    If InStr(Trim(cboNew.Text), " 房间") > 0 Then
        strBeds = Mid(Trim(cboNew.Text), 1, InStr(Trim(cboNew.Text), " 房间") - 1)
        
        strRoom = Mid(Trim(cboNew.Text), InStr(Trim(cboNew.Text), "房间:") + 3)
        
        strSql = "Select 性别 From 病人信息 A,床位状况记录 B  Where A.病人ID = b.病人id And b.病人ID Is Not Null And 病区ID = [1] And 房间号 =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnit, strRoom)
        
        Do While Not rsTmp.EOF
         
            If Trim(mrsPatiInfo!性别) <> rsTmp!性别 Then
                If (MsgBox("目标床位所在房间存在男女混住情况，是否继续入住？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)) = vbYes Then
                    Exit Do
                Else
                    Exit Sub
                    cboNew.SetFocus
                End If
            End If
            rsTmp.MoveNext
        Loop
    Else
        strBeds = Trim(cboNew.Text)
    End If
    '判断当前床位所在房间是否存在男女混住情况
    strSql = "select 房间号 from 床位状况记录 where 病区id=[1] and 床号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnit, mstr床号)
    If Nvl(rsTmp!房间号) <> "" Then
        strOldRoom = Nvl(rsTmp!房间号)
        strSql = "Select 性别 From 病人信息 A,床位状况记录 B  Where A.病人ID = b.病人id And b.病人ID Is Not Null And 病区ID = [1] And 房间号 =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnit, Val("" + rsTmp!房间号))
        
        Do While Not rsTmp.EOF
            If Trim(mrsSwapPatiInfo!性别) <> rsTmp!性别 Then
                If (MsgBox("目标床位所在房间存在男女混住情况，是否继续入住？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)) = vbYes Then
                    Exit Do
                Else
                    Exit Sub
                    cboNew.SetFocus
                End If
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    '包床病人不允许进行床位对换
    Set rsTmp = GetPatiBeds(mlng病人ID)
    If rsTmp.RecordCount > 1 Then
        MsgBox mstr床号 & "床病人为包床病人，不允许进行床位对换！", vbInformation, gstrSysName
        Exit Sub
    Else
        Set rsTmp = GetPatiBeds(mlngBeSwap病人ID)
        If rsTmp.RecordCount > 1 Then
            MsgBox strBeds & "床病人为包床病人，不允许进行床位对换！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '82383:LPF,换床检查床位是否为空,不为空不允许换床：床位对换如果目标床位不为空，则检查是否是床位对换的病人，病人不同不允许进行换床
    ReDim Preserve arrSQL(0)
    arrSQL(UBound(arrSQL)) = "zl_病人变动记录_Move(" & mlng病人ID & "," & mlng主页ID & "," & _
        "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'),'" & strBeds & "'," & _
        "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngUnit & ",'" & strBeds & "'," & mlngBeSwap病人ID & ")"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_病人变动记录_Move(" & mlngBeSwap病人ID & "," & mlngBeSwap主页ID & "," & _
        "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'),'" & mstr床号 & "'," & _
        "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngUnit & ",'" & mstr床号 & "'," & mlng病人ID & ")"
            
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For intLoop = LBound(arrSQL) To UBound(arrSQL)
        zlDatabase.ExecuteProcedure arrSQL(intLoop), Me.Caption
    
        If Val("" & mrsPatiInfo!险类) <> 0 Then
            If Not gclsInsure.ModiPatiSwap(mlng病人ID, mlng主页ID, Val("" & mrsPatiInfo!险类), "1") Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        ElseIf Val("" & mrsSwapPatiInfo!险类) <> 0 Then
            If Not gclsInsure.ModiPatiSwap(mlngBeSwap病人ID, mlngBeSwap主页ID, Val("" & mrsSwapPatiInfo!险类), "1") Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
    Next
    gcnOracle.CommitTrans: blnTrans = False
    '新网96847、118004
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(2, mlng病人ID, mlng主页ID) <> 1 Or gobjXWHIS.HISModPati(2, mlngBeSwap病人ID, mlngBeSwap主页ID) <> 1 Then
            MsgBox "当前启用了影像信息系统接口，但由于影像信息系统接口(HISModPati)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "当前启用了影像信息系统接口，但由于RIS接口创建失败未调用(HISModPati)接口，请与系统管理员联系。", vbInformation, gstrSysName
    End If
    mstr目标床号 = strBeds
    gblnOK = True
           
    On Error Resume Next
    '换床成功后触发消息
    '--病人一
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
        mclsXML.appendData "patient_sex", Nvl(mrsPatiInfo!性别), xsString '性别
        'in_number       住院号  1   S
        mclsXML.appendData "in_number", txt住院号.Text, xsString  '住院号
        mclsXML.AppendNode "in_patient", True
        
        '当前情况
        'current_state       当前情况    1
        mclsXML.AppendNode "current_state"
        'current_area_id     当前病区id  0..1    N
        mclsXML.appendData "current_area_id", Val(Nvl(mrsPatiInfo!当前病区ID)), xsNumber
        'current_area_title      当前病区    0..1    S
        mclsXML.appendData "current_area_title", Nvl(mrsPatiInfo!当前病区), xsString
        'current_dept_id     当前科室id  1   N
        mclsXML.appendData "current_dept_id", Val(Nvl(mrsPatiInfo!出院科室id, 0)), xsNumber
        'current_dept_title      当前科室    1   S
        mclsXML.appendData "current_dept_title", Nvl(mrsPatiInfo!当前科室), xsString
        'current_room        当前病房    0..1    S
        mclsXML.appendData "current_room", strOldRoom, xsString
        'current_bed     当前病床    1   S
        mclsXML.appendData "current_bed", mstr床号, xsString
        mclsXML.AppendNode "current_state", True
        
        strSql = " Select ID 变动id,开始时间 变动时间 From 病人变动记录 Where 病人ID=[1] And 主页Id=[2] And 开始原因=[3] And 开始时间+0=[4] And NVL(附加床位,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病人变动记录", mlng病人ID, mlng主页ID, 4, CDate(Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss")))
        '转入信息
        'change_state        转入信息    1
        mclsXML.AppendNode "change_state"
        'change_id       转科变更id  1   N
        mclsXML.appendData "change_id", rsTmp!变动ID, xsNumber
        'change_date     变更时间    1   S
        mclsXML.appendData "change_date", Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss"), xsString
        'change_room     入住病房    0..1    S
        mclsXML.appendData "change_room", strRoom, xsString
        'change_bed      入住病床    1   S
        mclsXML.appendData "change_bed", mstr目标床号, xsString
        mclsXML.AppendNode "change_state", True

        mclsMipModule.CommitMessage "ZLHIS_PATIENT_004", mclsXML.XmlText
    End If
    '--病人二
    If mclsMipModule.IsConnect = True Then
         mclsXML.ClearXmlText '清除缓存中的XML
        '--进行消息组装
        '病人信息
        mclsXML.AppendNode "in_patient"
        'patient_id      病人id  1   N
        mclsXML.appendData "patient_id", mlngBeSwap病人ID, xsNumber  '病人ID
        'page_id     主页id  1   N
        mclsXML.appendData "page_id", mlngBeSwap主页ID, xsNumber  '主页ID
        'patient_name        姓名    1   S
        mclsXML.appendData "patient_name", txtSwap姓名.Text, xsString '姓名
        'patient_sex     性别    0..1    S
        mclsXML.appendData "patient_sex", Nvl(mrsSwapPatiInfo!性别), xsString '性别
        'in_number       住院号  1   S
        mclsXML.appendData "in_number", txtSwap住院号.Text, xsString  '住院号
        mclsXML.AppendNode "in_patient", True
        
        '当前情况
        'current_state       当前情况    1
        mclsXML.AppendNode "current_state"
        'current_area_id     当前病区id  0..1    N
        mclsXML.appendData "current_area_id", Val(Nvl(mrsSwapPatiInfo!当前病区ID)), xsNumber
        'current_area_title      当前病区    0..1    S
        mclsXML.appendData "current_area_title", Nvl(mrsSwapPatiInfo!当前病区), xsString
        'current_dept_id     当前科室id  1   N
        mclsXML.appendData "current_dept_id", Val(Nvl(mrsSwapPatiInfo!出院科室id, 0)), xsNumber
        'current_dept_title      当前科室    1   S
        mclsXML.appendData "current_dept_title", Nvl(mrsSwapPatiInfo!当前科室), xsString
        'current_room        当前病房    0..1    S
        mclsXML.appendData "current_room", strRoom, xsString
        'current_bed     当前病床    1   S
        mclsXML.appendData "current_bed", txtSwapPre.Text, xsString
        mclsXML.AppendNode "current_state", True
        
        strSql = " Select ID 变动id,开始时间 变动时间 From 病人变动记录 Where 病人ID=[1] And 主页Id=[2] And 开始原因=[3] And 开始时间+0=[4] And NVL(附加床位,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病人变动记录", mlngBeSwap病人ID, mlngBeSwap主页ID, 4, CDate(Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss")))
        '转入信息
        'change_state        转入信息    1
        mclsXML.AppendNode "change_state"
        'change_id       转科变更id  1   N
        mclsXML.appendData "change_id", rsTmp!变动ID, xsNumber
        'change_date     变更时间    1   S
        mclsXML.appendData "change_date", Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss"), xsString
        'change_room     入住病房    0..1    S
        mclsXML.appendData "change_room", strOldRoom, xsString
        'change_bed      入住病床    1   S
        mclsXML.appendData "change_bed", mstr床号, xsString
        mclsXML.AppendNode "change_state", True

        mclsMipModule.CommitMessage "ZLHIS_PATIENT_004", mclsXML.XmlText
    End If
    If Err <> 0 Then Err.Clear
    
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    gblnOK = False
    
    If Not InitData Then Unload Me: Exit Sub
    
    fraBed.Caption = mstr床号 & "床病人"
    fraBedSwap.Caption = mstr目标床号 & "床病人"
    If cboNew.ListCount = 0 Then
        MsgBox "病人所在科室的病区已没有合适的床位可供对换！", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    '创建消息对象
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
End Sub

Private Function InitData() As Boolean
    Dim i As Integer, rsTmp As ADODB.Recordset, str床号 As String
    
    Set mrsPatiInfo = GetPatiInfo(mlng病人ID, mlng主页ID)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    
    With mrsPatiInfo
        txt姓名.Text = !姓名
        txt住院号.Text = "" & !住院号
        txt科室.Text = !当前科室
        If Trim(mstr床号) = "" Then mstr床号 = !当前床号
    End With
    
    txtPre.Text = mstr床号
    '初始化床位
    If InitBed(mlngUnit) = False Then Exit Function
    
    InitData = True
End Function

Private Function InitBed(ByVal lng病区ID As Long) As Boolean
'功能：初始化床位,此时取该病区及科室对应的所有空床
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strSQLtmp As String, i As Integer
    Dim bytLen As Byte
    Dim strTmp As String
    
    On Error GoTo errH
        
    If InStr(mrsPatiInfo!性别, "男") > 0 Then
        strTmp = "男床,不限床"
    ElseIf InStr(mrsPatiInfo!性别, "女") > 0 Then
        strTmp = "女床,不限床"
    Else
        strTmp = "不限床"
    End If
        
    cboNew.Clear

    bytLen = GetMaxBedLen(lng病区ID)
    
    Set rsTmp = GetPatiBeds(mlng病人ID)
    
    If rsTmp.RecordCount > 1 Then
        MsgBox "该病人为包床病人，不允许进行床位对换！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '性别限制
    If rsTmp!性别分类 = "不限床" Then
        strSQLtmp = " And (A.性别分类 ='不限床' Or (A.性别分类 = '男床' And '" & rsTmp!性别 & "'='男') Or (A.性别分类 = '女床' And '" & rsTmp!性别 & "'='女')) "
    ElseIf rsTmp!性别分类 = "男床" Then
        strSQLtmp = " And ((A.性别分类 = '不限床' And B.性别 = '男') Or (A.性别分类 = '男床' And '" & rsTmp!性别 & "'='男'))"
    ElseIf rsTmp!性别分类 = "女床" Then
        strSQLtmp = " And ((A.性别分类 = '不限床' And B.性别 = '女') Or (A.性别分类 = '女床' And '" & rsTmp!性别 & "'='女'))"
    End If
    
    strSql = "Select Distinct A.床号,A.性别分类,A.房间号,A.等级ID,B.性别,C.状态 From 床位状况记录 A, 病人信息 B, 病案主页 C, 病人变动记录 D " & vbNewLine & _
                " Where A.病人ID=c.病人ID And B.病人ID=C.病人ID And C.病人ID=D.病人ID And C.主页ID=D.主页ID And (" & _
                IIf(rsTmp!共用 = 1, " A.科室ID is Null Or A.科室ID=[1] Or A.共用=1 ", "A.科室ID is Null Or A.科室ID=[1] Or (A.共用=1 And A.科室id=[1])") & _
                ") And A.病区ID=[2] And A.状态='占用' And C.状态 Not In(2,3)" & vbNewLine & _
                strSQLtmp & " Order by  LPad(NVL(A.房间号,0), 10, ' '),LPad(A.床号, 10, ' ')"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsPatiInfo!出院科室id), lng病区ID)
    Set mrsBeds = rsTmp.Clone
    
    For i = 1 To rsTmp.RecordCount
        If Not rsTmp!床号 = mstr床号 Then cboNew.AddItem Space(bytLen - Len(rsTmp!床号)) & rsTmp!床号 & IIf(IsNull(rsTmp!房间号), "", " 房间:" & rsTmp!房间号)
        If rsTmp!床号 = mstr目标床号 Then cboNew.ListIndex = cboNew.NewIndex
        rsTmp.MoveNext
    Next
    InitBed = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str床号 As String, _
            ByRef str目标床号 As String, ByVal strPrivs As String) As Boolean
    On Error Resume Next
    Set mfrmParent = frmParent
    mlngUnit = lngUnit
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstr床号 = str床号
    mstr目标床号 = str目标床号
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    str目标床号 = mstr目标床号
    ShowMe = gblnOK
End Function

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

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub
