VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQuestion 
   Caption         =   "电子病案审查评分"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6555
   Icon            =   "frmQuestion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   6555
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7335
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQuestion.frx":000C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8652
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsCondition As ADODB.Recordset
Private mfrmParent As Object
Private mblnOK As Boolean
Private mlngModul As Long
Private mstrPrivs As String
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mblnAuditEnter As Boolean '允许自由录入审查意见

Private WithEvents mfrmChildQuestion As frmChildQuestion
Attribute mfrmChildQuestion.VB_VarHelpID = -1
Public Event ShowInfo(ByVal strShowInfo As Long)


Private Property Let DataChanged(ByVal blnData As Boolean)
    mfrmChildQuestion.DataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    If Not (mfrmChildQuestion Is Nothing) Then
        DataChanged = mfrmChildQuestion.DataChanged
    End If
End Property

'################################################################################################################
'   用途：  系统入口。
'################################################################################################################
Public Sub ShowMe(ByVal frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long)

    On Error GoTo errHand
    Dim lng提交Id As Long
    Dim lng出院科室ID As Long
    '初始传入数据
    Set mfrmParent = frmParent
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    
    '初始系统数据
    mlngModul = 1560
    mblnAuditEnter = True '允许自由录入审查意见
    mstrPrivs = GetPrivFunc(glngSys, mlngModul) '获取权限
    
    '初始化
    Call ExecuteCommand("初始控件")
    Call ExecuteCommand("初始数据")
    Call ExecuteCommand("病案审查")
    
    
    '显示窗体
    
    'Me.Show vbModal, mfrmParent
    mfrmChildQuestion.Show vbModal, mfrmParent
    
    If mblnOK Then
        
    Else
 
    End If
    
    Set mrsCondition = Nothing
    If Not (mfrmChildQuestion Is Nothing) Then Unload mfrmChildQuestion
    
    Unload Me
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop         As Integer
    Dim strTmp As String
    
    On Error GoTo errHand

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        Set mfrmChildQuestion = New frmChildQuestion
        Call mfrmChildQuestion.InitData(Me, mlngModul, IsPrivs(mstrPrivs, "审查病案"), mblnAuditEnter, mstrPrivs)
   
        
     Case "初始数据"
                                
        '创建过滤条件项目，并进行初始化
        Call ParamCreate(mrsCondition)
        
        Call ParamAdd(mrsCondition, "等待接收", 1)
        Call ParamAdd(mrsCondition, "拒绝接收", 1)
        Call ParamAdd(mrsCondition, "正在审查", 1)
        Call ParamAdd(mrsCondition, "审查反馈", 1)
        Call ParamAdd(mrsCondition, "审查整改", 1)
        
        Call ParamAdd(mrsCondition, "当前病况", "")
        Call ParamAdd(mrsCondition, "出院情况", "")
        
        Call ParamAdd(mrsCondition, "病人类型", 0)
        Call ParamAdd(mrsCondition, "医保种类", "")
        
        Call ParamAdd(mrsCondition, "审查开始时间", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "审查结束时间", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "归档开始时间", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "归档结束时间", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    
        Call ParamAdd(mrsCondition, "出院病人", 0)
        Call ParamAdd(mrsCondition, "出院开始时间", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "出院结束时间", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        
        Call ParamAdd(mrsCondition, "医嘱开始时间", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "医嘱结束时间", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "住院医师", "")
        Call ParamAdd(mrsCondition, "疾病名称", "")
        Call ParamAdd(mrsCondition, "检查类型", "")
        Call ParamAdd(mrsCondition, "药品信息", "")
                
        '读取缺省时间范围
        strTmp = GetPara("审查缺省范围", mlngModul, "今  天")
        If strTmp = "" Then strTmp = "今  天"
        Call ParamWrite(mrsCondition, "审查开始时间", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "审查结束时间", GetDateTime(strTmp, 2))
        
        strTmp = GetPara("归档缺省范围", mlngModul, "今  天")
        If strTmp = "" Then strTmp = "今  天"
        Call ParamWrite(mrsCondition, "归档开始时间", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "归档结束时间", GetDateTime(strTmp, 2))
        
        strTmp = GetPara("出院缺省范围", mlngModul, "今  天")
        If strTmp = "" Then strTmp = "今  天"
        Call ParamWrite(mrsCondition, "出院开始时间", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "出院结束时间", GetDateTime(strTmp, 2))
        
        '新加条件
        strTmp = GetPara("医嘱缺省范围", mlngModul, "今  天")
        If strTmp = "" Then strTmp = "今  天"
        Call ParamWrite(mrsCondition, "医嘱开始时间", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "医嘱结束时间", GetDateTime(strTmp, 2))
    Case "病案审查"
        Dim strObject As String
        Dim strParam As String
        Dim lng提交Id As Long
        
        If Not (mfrmChildQuestion Is Nothing) Then
            strObject = "首页记录"
            lng提交Id = GetSubmitID(mlng病人ID, mlng主页ID)
            Call mfrmChildQuestion.SetParamter(mlng病人ID, mlng主页ID, strObject, strParam, lng提交Id)
            mfrmChildQuestion.AllowModify = True
            Call mfrmChildQuestion.RefreshData("", mrsCondition, mblnAuditEnter) 'GetChildPatient(mintIndex).Depts
            
            
        End If
    
    End Select
    
    ExecuteCommand = True

    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
EndHand:
End Function

Public Property Get 模块号() As Long
    模块号 = mlngModul
End Property

Private Sub mfrmChildQuestion_AfterDataChanged()
    ' Call ExecuteCommand("控件状态")
End Sub

Private Sub mfrmChildQuestion_AfterDeleteQuestion(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    ' Call ExecuteCommand("刷新指定病人", lng病人ID, lng主页ID)
End Sub

Private Sub mfrmChildQuestion_AfterQuestionType(ByVal blnQuestionType As Boolean)
    'blnQuestionType=True 院级反馈 =Flase 科级反馈
'    If blnQuestionType Then
'        If ObjPtr(dkpMain.Panes(1)) > 0 Then
'            dkpMain.Panes(1).Title = "院级问题反馈"
'        End If
'    Else
'        If ObjPtr(dkpMain.Panes(1)) > 0 Then
'            dkpMain.Panes(1).Title = "科级问题反馈"
'        End If
'    End If
End Sub

Private Sub mfrmChildQuestion_AfterSaveQuestion(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    ' Call ExecuteCommand("刷新指定病人", lng病人ID, lng主页ID)
End Sub

Private Sub mfrmChildQuestion_LocationDocument(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal byt反馈对象 As Byte, ByVal lng文件ID As Long, ByVal lng医嘱id As Long, ByVal lng科室ID As Long)
    '根据信息定位到指定病人的指定病案资料上去
    On Error GoTo errHand
    
    
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)

End Sub
