VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPreOut 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人预出院"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   Icon            =   "frmPreOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3570
      TabIndex        =   2
      Top             =   1875
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2475
      TabIndex        =   1
      Top             =   1875
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   1710
      Left            =   120
      TabIndex        =   4
      Top             =   15
      Width           =   4710
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   1845
         TabIndex        =   0
         Top             =   1125
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   195
         Picture         =   "frmPreOut.frx":058A
         Top             =   315
         Width           =   480
      End
      Begin VB.Label lblInfo 
         Caption         =   "请输入 ""XXX"" 的预出院时间，预出院之后，不具有相关权限的人员不能再对病人计费，指定时间之后的自动费用也不再发生。"
         ForeColor       =   &H00C00000&
         Height          =   525
         Left            =   855
         TabIndex        =   6
         Top             =   330
         Width           =   3600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预出院时间"
         Height          =   180
         Left            =   840
         TabIndex        =   5
         Top             =   1185
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   285
      TabIndex        =   3
      Top             =   1875
      Width           =   1100
   End
End
Attribute VB_Name = "frmPreOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOk As Boolean
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstr姓名 As String
Private mstrPrivs As String
Private mfrmParent As Object

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Public Function ShowMe(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str姓名 As String, ByVal strPrivs As String) As Boolean
    Set mfrmParent = frmParent
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstr姓名 = str姓名
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    
    ShowMe = mblnOk
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim curDate As Date, dMax As Date
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    If Not IsDate(txtDate.Text) Then
        MsgBox "请输入正确的时间值。", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    '时间不能超过当前时间太长(一周)
    curDate = zlDatabase.Currentdate
    If CDate(txtDate.Text) > curDate Then
        If CDate(txtDate.Text) - curDate > 7 Then
            MsgBox "预出院时间比当前时间大得过多,请检查！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If MsgBox("预出院时间大于了当前系统时间,确实要预出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
    
    dMax = GetMaxDate(mlng病人ID, mlng主页ID)
    If Format(txtDate.Text, "yyyyMMddHHmmss") <= Format(dMax, "yyyyMMddHHmmss") Then
        MsgBox "病人预出院时间必须大于该病人上次变动时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    strSQL = "zl_病人变动记录_PreOut(" & mlng病人ID & "," & mlng主页ID & ",To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'))"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    On Error GoTo 0
    mblnOk = True
    
    If mclsMipModule.IsConnect = True Then
        strSQL = _
                " Select a.姓名, a.性别, a.住院号, a.当前病区id, b.名称　当前病区, a.出院科室id 当前科室id," & _
                        " c.名称 当前科室, d.房间号 当前病房, a.出院病床 当前床号, e.Id  变动id" & _
                " From 病案主页 a,病人变动记录 e, 床位状况记录 d, 部门表 b, 部门表 c" & _
                " Where a.病人id = e.病人id And a.主页id = e.主页id And a.病人id = d.病人id(+)  And a.当前病区id = d.病区id(+) And a.出院病床 = d.床号(+) " & _
                    " And a.当前病区id = b.Id(+) And a.出院科室id = c.Id(+) And a.病人id = [1] And  a.主页id = [2] And e.开始原因 = [3] And Nvl(e.附加床位, 0) = 0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人变动记录", mlng病人ID, mlng主页ID, 10)
        
        mclsXML.ClearXmlText '清除缓存中的XML
        '--进行消息组装
        '病人信息
        mclsXML.AppendNode "in_patient"
        'patient_id      病人id  1   N
        mclsXML.appendData "patient_id", mlng病人ID, xsNumber  '病人ID
        'page_id     主页id  1   N
        mclsXML.appendData "page_id", mlng主页ID, xsNumber '主页ID
        'patient_name        姓名    1   S
        mclsXML.appendData "patient_name", Nvl(rsTmp!姓名), xsString '姓名
        'patient_sex     性别    0..1    S
        mclsXML.appendData "patient_sex", Nvl(rsTmp!性别), xsString '性别
        'in_number       住院号  1   S
        mclsXML.appendData "in_number", Nvl(rsTmp!住院号), xsString '住院号
        mclsXML.AppendNode "in_patient", True
        
        'out_prehospital     病人预出院  1
        mclsXML.AppendNode "out_prehospital"
        'change_id       变更id  1   N
        mclsXML.appendData "change_id", Nvl(rsTmp!变动id), xsNumber
        'out_date        预出院时间  1   s
        mclsXML.appendData "out_date", Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss"), xsString
        'out_area_id     当前病区id  0..1    N
        mclsXML.appendData "out_area_id", Nvl(rsTmp!当前病区ID, 0), xsNumber
        'out_area_title      当前病区    0..1    S
        mclsXML.appendData "out_area_title", Nvl(rsTmp!当前病区), xsString
        'out_dept_id     当前科室id    1   N
        mclsXML.appendData "out_dept_id", Nvl(rsTmp!当前科室id, 0), xsNumber
        'out_dept_title      当前科室  1   S
        mclsXML.appendData "out_dept_title", Nvl(rsTmp!当前科室id), xsString
        'out_room        当前病房    0..1    S
        mclsXML.appendData "out_room", Nvl(rsTmp!当前病房), xsString
        'out_bed     当前病床    1   S
        mclsXML.appendData "out_bed", Nvl(rsTmp!当前床号), xsString
        'order_id        医嘱id  0..1    N
        mclsXML.AppendNode "out_prehospital", True
        mclsMipModule.CommitMessage "ZLHIS_PATIENT_009", mclsXML.XmlText
    End If
    
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    mblnOk = False
    '--55791:刘鹏飞,2012-11-13,作废出院医嘱才能撤销出院
    If gbln医生允许才能出院 Then
        If Not Check医生下达出院医嘱(mlng病人ID, mlng主页ID) Then
            MsgBox "医生尚未下达出院(或转院、死亡)医嘱，不能直接进行预出院操作！", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
    
    lblInfo.Caption = Replace(lblInfo.Caption, "XXX", mstr姓名)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
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

Private Sub txtDate_GotFocus()
    Call zlControl.TxtSelAll(txtDate)
End Sub
