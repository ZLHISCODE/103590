VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholSpecimen_AcceptOrReject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "核收标本"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6390
   Icon            =   "frmPatholSpecimen_AcceptOrReject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame framStudyInf 
      Height          =   3480
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   6135
      Begin VB.ComboBox cbxReject 
         Height          =   300
         ItemData        =   "frmPatholSpecimen_AcceptOrReject.frx":179A
         Left            =   1080
         List            =   "frmPatholSpecimen_AcceptOrReject.frx":17B6
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2160
         Width           =   4815
      End
      Begin VB.ComboBox cbxSubmitDoctor 
         Height          =   300
         Left            =   4200
         TabIndex        =   16
         Top             =   240
         Width           =   1665
      End
      Begin VB.TextBox txtRejectNotify 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   15
         Top             =   1680
         Width           =   1785
      End
      Begin VB.TextBox txtRegisterDoctor 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   4200
         TabIndex        =   14
         Top             =   1680
         Width           =   1785
      End
      Begin VB.TextBox txtContactWay 
         Height          =   300
         Left            =   1080
         TabIndex        =   13
         Top             =   1200
         Width           =   1785
      End
      Begin VB.TextBox txtFormDepart 
         Height          =   300
         Left            =   4200
         TabIndex        =   12
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtUnitName 
         Height          =   300
         Left            =   1080
         TabIndex        =   11
         Text            =   "本院"
         Top             =   720
         Width           =   1785
      End
      Begin VB.TextBox txtRejectReason 
         Height          =   780
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2520
         Width           =   4770
      End
      Begin VB.ComboBox cbxStudyType 
         ForeColor       =   &H00FF0000&
         Height          =   300
         ItemData        =   "frmPatholSpecimen_AcceptOrReject.frx":18E0
         Left            =   1080
         List            =   "frmPatholSpecimen_AcceptOrReject.frx":18E2
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker dtpSubmitTime 
         Height          =   300
         Left            =   4200
         TabIndex        =   17
         Top             =   1200
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   71696387
         CurrentDate     =   40646.4399652778
      End
      Begin VB.Label labRejectNotify 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "通 知 人："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   1740
         Width           =   900
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "登 记 人："
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3240
         TabIndex        =   28
         Top             =   1740
         Width           =   900
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "送检日期："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3240
         TabIndex        =   27
         Top             =   1260
         Width           =   900
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "联系方式："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   1260
         Width           =   900
      End
      Begin VB.Label labSubmitDoctor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "送 检 人："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3240
         TabIndex        =   25
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "送检科室："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3240
         TabIndex        =   24
         Top             =   780
         Width           =   900
      End
      Begin VB.Label labUnitName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "送检单位："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   780
         Width           =   900
      End
      Begin VB.Label labRejectReason 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "拒收理由："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label labStudyType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "检查类型："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5880
         TabIndex        =   20
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5880
         TabIndex        =   19
         Top             =   765
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5880
         TabIndex        =   18
         Top             =   2520
         Width           =   255
      End
   End
   Begin VB.TextBox txtPatholNum 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   4890
   End
   Begin VB.PictureBox picShow 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   3495
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   3495
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
         TabIndex        =   5
         Top             =   120
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdReject_Cancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   5040
      TabIndex        =   2
      Top             =   5115
      Width           =   1215
   End
   Begin VB.CommandButton cmdReject_Sure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   3720
      TabIndex        =   1
      Top             =   5115
      Width           =   1215
   End
   Begin VB.Label labPatholNumNeed 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label labPatholNum 
      Caption         =   "病理号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmPatholSpecimen_AcceptOrReject.frx":18E4
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6360
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "    请仔细核对送检标本，并正确录入标本核/拒收的详细信息，当标本被核收后，将不能对其修改或删除。"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   195
      Width           =   5175
   End
End
Attribute VB_Name = "frmPatholSpecimen_AcceptOrReject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnIsRejectSpecimen As Boolean
Private mblnIsSucceed As Boolean
Private mlngCurAdviceId As Long

'Public mlngCurStudyProcedure As Long
Public mstrCurDepartmentId As String


Public mtxtAcceptHistory As RichTextBox

Public mstrPatholNum As String
Private mlngStudyType As Long

Public mlngPatholSerialNum As Long
Public mstrPatholInitNum As String

Public mobjSquareCard As Object    '一卡通，卡结算部件

Public mfrmParent As Form

Public mstrPrivs As String          '调用者的权限

Property Get IsRejectSpecimen() As Boolean
    IsRejectSpecimen = mblnIsRejectSpecimen
End Property

Property Let IsRejectSpecimen(value As Boolean)
    mblnIsRejectSpecimen = value
End Property


Property Get AdviceId() As Long
    AdviceId = mlngCurAdviceId
End Property

Property Let AdviceId(value As Long)
    mlngCurAdviceId = value
End Property




Property Get IsSucceed() As Boolean
    IsSucceed = mblnIsSucceed
End Property


Public Function ShowAcceptOrRejectSpecimenWindow(lngAdviceID As Long, _
    ByVal lngCurDepartmentId As String, txtAcceptHis As RichTextBox, blnIsReject As Boolean, owner As Form, _
    strPrivs As String) As Boolean
    
    Dim frmAOR As New frmPatholSpecimen_AcceptOrReject
    
    On Error GoTo errFree
    
    With frmAOR
        .AdviceId = lngAdviceID
'        .mlngCurStudyProcedure = lngCurStudyProcedure
        .mstrCurDepartmentId = lngCurDepartmentId
        .mlngPatholSerialNum = 0
        .mstrPatholInitNum = ""
        .mstrPrivs = strPrivs
        
        Set .mtxtAcceptHistory = txtAcceptHis
        Set .mfrmParent = owner
        
        .IsRejectSpecimen = blnIsReject
        
        .txtRejectReason.Text = ""
        .dtpSubmitTime.value = zlDatabase.Currentdate
        .txtRegisterDoctor.Text = UserInfo.姓名
        
        If blnIsReject Then
            frmAOR.Caption = "拒收标本"
            
            .labSubmitDoctor.Left = .labStudyType.Left
            .cbxSubmitDoctor.Left = .cbxStudyType.Left
            .cbxSubmitDoctor.Width = .txtRejectReason.Width
            
            .labStudyType.Visible = False
            .cbxStudyType.Visible = False
            
            .framStudyInf.Top = .txtPatholNum.Top
            .picShow.Top = .framStudyInf.Top + .framStudyInf.Height + 2400
            
            .cmdReject_Sure.Top = .framStudyInf.Top + .framStudyInf.Height + 120
            .cmdReject_Cancel.Top = .cmdReject_Sure.Top
            
            .Height = .cmdReject_Sure.Top + .cmdReject_Sure.Height + 120 + 430
        Else
            frmAOR.Caption = "核收标本"
            
            .txtRejectReason.Visible = False
            .cbxReject.Visible = False
            
            .labRejectNotify.Visible = False
            .txtRejectNotify.Visible = False
            
            .Label24.Left = .labRejectNotify.Left
            .txtRegisterDoctor.Left = .txtRejectNotify.Left
            
            .framStudyInf.Height = 2160
            .Height = 4655
            
            .cmdReject_Sure.Top = .ScaleHeight - .cmdReject_Sure.Height - 120
            .cmdReject_Cancel.Top = .cmdReject_Sure.Top
            .picShow.Top = .cmdReject_Sure.Top - 120
        End If
        
        '读取检查类型
        Call .GetStudyAcceptInf(lngAdviceID)
        Call .ConfigStudyType
        Call .ConfigSubmitInf(lngAdviceID)
        
        .txtPatholNum.Visible = Not blnIsReject
        .labPatholNum.Visible = Not blnIsReject
        .labPatholNumNeed.Visible = Not blnIsReject
        
        .cbxReject.Enabled = blnIsReject
        .cbxReject.BackColor = IIf(blnIsReject, &H80000005, &H8000000F)
        
        .txtRejectReason.Enabled = blnIsReject
        .txtRejectReason.BackColor = IIf(blnIsReject, &H80000005, &H8000000F)
        
        .txtRejectNotify.Enabled = blnIsReject
        .txtRejectNotify.BackColor = IIf(blnIsReject, &H80000005, &H8000000F)
        
        .labRejectReason.Enabled = blnIsReject
        .labRejectNotify.Enabled = blnIsReject
        
        Call .CloseProcessHint
        
        If Trim(.mstrPatholNum) = "" Then
            '根据默认加载的检查类型得到相关的病理号
            .txtPatholNum.Text = GetPatholNum(Val(.cbxStudyType.Text))
        End If
    End With
    

    Call frmAOR.Show(1, owner)
    
    ShowAcceptOrRejectSpecimenWindow = frmAOR.IsSucceed
        
errFree:
    Unload frmAOR
    Set frmAOR = Nothing
End Function




Public Sub ConfigSubmitInf(lngAdviceID As Long)
'配置送检信息
    If lngAdviceID <= 0 Then Exit Sub
    
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strSubmitDoctor As String
    
    strSQL = "select 名称 from 部门表 a, 病人医嘱记录 b where a.id =b.开嘱科室id and b.id=[1]"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    txtFormDepart.Text = rsData!名称
    
    '读取送检人员信息
    strSQL = "select case when c.开嘱医生=a.姓名 then 1 else 0 end as 是否送检, a.姓名 from 人员表 a, 部门人员 b, 病人医嘱记录 c where a.id=b.人员id and b.部门Id=c.开嘱科室Id and c.id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    strSubmitDoctor = ""
    Call cbxSubmitDoctor.Clear
    While Not rsData.EOF
        Call cbxSubmitDoctor.AddItem(Nvl(rsData!姓名))
        strSubmitDoctor = IIf(Val(Nvl(rsData!是否送检)) = 1, Nvl(rsData!姓名), strSubmitDoctor)
        
        rsData.MoveNext
    Wend
    
    If strSubmitDoctor <> "" Then
        cbxSubmitDoctor.Text = strSubmitDoctor
    End If
End Sub


Public Sub GetStudyAcceptInf(ByVal lngAdviceID As Long)
'获取检查核收信息
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select 病理号,检查类型 from 病理检查信息 where 医嘱ID=[1]"
    'If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    
    mstrPatholNum = ""
    mlngStudyType = -1
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    mstrPatholNum = Nvl(rsData!病理号)
    mlngStudyType = Val(Nvl(rsData!检查类型))
End Sub


Public Sub ConfigStudyType()
'配置检查的录入类型
    If mlngStudyType < 0 Then Exit Sub
    
    cbxStudyType.ListIndex = mlngStudyType
    txtPatholNum.Text = mstrPatholNum
    
    cbxStudyType.BackColor = &H8000000F
    cbxStudyType.Enabled = False
    
    txtPatholNum.BackColor = &H8000000F
    txtPatholNum.Enabled = False
    
    labPatholNum.Enabled = False
End Sub


Private Sub LoadStudyType()
    '载入标本类型
    Dim strSQL As String
    Dim rsStudyType As ADODB.Recordset
    
    Call cbxStudyType.Clear
    
    Call cbxStudyType.AddItem("0-常规")
    Call cbxStudyType.AddItem("1-冰冻")
    Call cbxStudyType.AddItem("2-细胞")
    Call cbxStudyType.AddItem("3-会诊")
    Call cbxStudyType.AddItem("4-尸检")
    Call cbxStudyType.AddItem("5-快速石蜡")
    
    strSQL = "select 执行分类 from 诊疗项目目录 where ID= (select 诊疗项目ID from 病人医嘱记录 where id=[1])"
    Set rsStudyType = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCurAdviceId)
    
    If rsStudyType.RecordCount > 0 Then
        '根据医嘱ID获得该医嘱的 执行分类 并 自动填入标本核收窗口的检查类型
        cbxStudyType.ListIndex = Val(Nvl(rsStudyType!执行分类))
    Else
        cbxStudyType.ListIndex = 0
    End If
    
End Sub


'Private Sub UpdateAcceptOrRejectHistory(blnReject As Boolean)
'    '更新历史核收记录
'    Dim curRecord As String
'    Dim lngStart As Long
'    Dim strFormats As String
'
'    If mtxtAcceptHistory.Text = "" Then
'        strFormats = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
'                    "{\colortbl ;\red255\green104\blue104;\red19\green164\blue251;}" & _
'                    "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\b\f0\fs20 "
'    Else
'        strFormats = mtxtAcceptHistory.TextRTF
'        strFormats = Mid(strFormats, 1, Len(strFormats) - 17)
''        strFormats = Replace(strFormats, "\par }", "")
'    End If
'
'    mtxtAcceptHistory.Text = ""
'    If blnReject Then
'        curRecord = dtpSubmitTime.value & "：由[ " & cbxSubmitDoctor.Text & " ]从[ " & txtUnitName.Text & txtFormDepart.Text & " ]送检的标本已被[ " & txtRegisterDoctor.Text & " ]拒收。已通知[ " & txtRejectNotify.Text & " ] 联系方式[ " & txtContactWay.Text & " ]"
'
'        strFormats = strFormats & "\cf1 " & curRecord & "\par"
'    Else
'        curRecord = dtpSubmitTime.value & "：由[ " & cbxSubmitDoctor.Text & " ]从[ " & txtUnitName.Text & txtFormDepart.Text & " ]送检的标本已被[ " & txtRegisterDoctor.Text & " ]核收。"
'
'        strFormats = strFormats & "\cf2 " & curRecord & "\par"
'    End If
'
'    If mtxtAcceptHistory.Text = "" Then
'        mtxtAcceptHistory.SelRTF = strFormats & "}"
'    Else
'        mtxtAcceptHistory.TextRTF = strFormats & "}"
'    End If
'    mtxtAcceptHistory.Refresh
'End Sub


Private Function CheckSubmitInfoIsValid() As String
    '检查送检信息是否有效
    CheckSubmitInfoIsValid = ""
    
    If Trim(txtFormDepart.Text) = "" Then
        CheckSubmitInfoIsValid = "送检科室不能为空。"
        
        Call txtFormDepart.SetFocus
        Exit Function
    End If
    
    If Trim(cbxSubmitDoctor.Text) = "" Then
        CheckSubmitInfoIsValid = "送检人不能为空。"
        
        Call cbxSubmitDoctor.SetFocus
        Exit Function
    End If
    
    If txtRejectReason.Enabled Then
        If Trim(txtRejectReason.Text) = "" Then
            CheckSubmitInfoIsValid = "拒收原因不能为空。"
            
            Call txtRejectReason.SetFocus
            Exit Function
        End If
    Else
        If Trim(txtPatholNum.Text) = "" Then
            CheckSubmitInfoIsValid = "病理号不能为空。"
            
            Call txtPatholNum.SetFocus
            Exit Function
        End If
    End If
End Function



Private Sub ShowProcessHint(ByVal strHint As String)
'显示处理信息
On Error Resume Next

    txtShow.Text = strHint

    picShow.Visible = True
End Sub


Public Sub CloseProcessHint()
'关闭处理提示
    picShow.Visible = False
End Sub

Private Sub cbxReject_Click()
On Error GoTo ErrHandle
    If Trim(txtRejectReason.Text) <> "" Then txtRejectReason.Text = txtRejectReason.Text & vbCrLf
    txtRejectReason.Text = txtRejectReason.Text & Mid(cbxReject.Text, InStr(cbxReject.Text, "-") + 1, 100)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Public Function GetPatholNum(ByVal lngStudyType As Long) As String
'根据检查类型获取病理号
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetPatholNum = ""
    
    strSQL = "select Zl_病理号码_序号获取([1]) as 病理序号 from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngStudyType)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    mlngPatholSerialNum = Val(Nvl(rsData!病理序号))
    
    strSQL = "select Zl_病理号码_生成([1],[2]) as 病理号 from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngStudyType, mlngPatholSerialNum)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    mstrPatholInitNum = Nvl(rsData!病理号)
    
    GetPatholNum = mstrPatholInitNum
End Function


Private Sub cbxStudyType_Click()
On Error GoTo ErrHandle
    If Trim(mstrPatholNum) = "" Then
        txtPatholNum.Text = GetPatholNum(Val(cbxStudyType.Text))
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdReject_Cancel_Click()
    mblnIsSucceed = False
    
    Call Me.Hide
End Sub


Private Function AutoRegister() As Boolean
'自动报到注册
'取出病人当前划价费用（当执行后自动审核划价单据有效时）
    Dim curMoney As Currency
    Dim str类别 As String
    Dim str类别名 As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngSourceType As Long
    Dim lngPatientID As Long
    Dim arrSQL() As Variant
    Dim i As Integer
    Dim blnTran As Boolean
    Dim rsOneCard As ADODB.Recordset
    Dim int记录性质 As Integer     '病人医嘱发送.记录性质，本次医嘱的记录性质，1-收费记录；2-记帐记录
    Dim int门诊记帐 As Integer     '病人医嘱发送.门诊记帐，门诊和住院医生站发送为门诊记帐时填为1,用于区分门诊记帐和住院记帐，其他的都填为空
    Dim str诊疗类别 As String
    Dim lng发送号 As Long
    Dim str单据号 As String
    Dim str医嘱IDs As String

On Error GoTo ErrHandle

    AutoRegister = True


    strSQL = "select A.病人来源,A.ID,A.姓名,A.性别,A.年龄,A.病人ID,A.主页ID,B.出生日期,B.当前病区ID, decode(c.医嘱id, null, '0',1) as 报到状态, D.发送号 " & _
            " from 病人医嘱记录 A, 病人信息 B, 影像检查记录 C, 病人医嘱发送 D " & _
            " where a.病人id = b.病人id and a.id = c.医嘱id(+) and a.ID =D.医嘱ID and a.id=[1]"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCurAdviceId)
    
    '如果已经写入检查信息，则不进行报到操作
    If Nvl(rsTemp!报到状态, 0) = 1 Then Exit Function
    
    
    lngSourceType = Nvl(rsTemp!病人来源, 3)
    lngPatientID = rsTemp!病人ID
    
    
    '检查费用以及一卡通的处理
        '业务逻辑是：
        '1、总体逻辑没有收费的不能报到，但是如果有“未缴费报到”权限的，可以在没有收费的情况下报到。
        '   在刷新信息的时候已经控制报到的确定按钮。
        '2、对公共基础参数的支持：
        '       参数号28--门诊一卡通，消费减少剩余款额时是否需要验证
        '       参数号81--执行后自动审核
        '       参数号163--门诊一卡通，项目执行前必须先收费或先记帐审核
        '3、先处理需要一卡通消费确认的，条件是以下之一
        '       （1）记录性质=1
        '       （2）执行后自动审核=False，记录性质=2，且 “来源<>住院”  或者 “来源=住院，门诊记帐”。
        '   如果一卡通消费确认成功，则可以报到。如果一卡通消费确认不成功，则根据权限“未缴费报到”提示是否继续报到。
        '4、再处理一卡通费用减少验证的，只处理记账的，条件是：
        '       （1）记录性质=2，执行后自动审核=True
        '       （2）有未审核费用
        '
        '
        '
        gstrSQL = "Select A.记录性质,A.门诊记帐,A.发送号,A.NO,B.诊疗类别 from 病人医嘱发送 A,病人医嘱记录 B  where A.医嘱ID=B.ID and  B.ID =[1]"
        Set rsOneCard = zlDatabase.OpenSQLRecord(gstrSQL, "PACS报到查找记录性质", mlngCurAdviceId)
        If rsOneCard.EOF = False Then
            int记录性质 = Nvl(rsOneCard!记录性质, 0)
            int门诊记帐 = Nvl(rsOneCard!门诊记帐, 0)
            str诊疗类别 = Nvl(rsOneCard!诊疗类别)
            lng发送号 = rsOneCard!发送号
            str单据号 = Nvl(rsOneCard!NO)
        End If
        
        If int记录性质 = 1 Or _
            (gbln执行后审核 = False And int记录性质 = 2 And (lngSourceType <> 2 Or (lngSourceType = 2 And int门诊记帐 = 1))) Then
            
            If Not ItemHaveCash(lngSourceType, False, mlngCurAdviceId, 0, lng发送号, str诊疗类别, str单据号, int记录性质, _
                int门诊记帐, 0) Then
                If gbln执行前先结算 Then
                    '门诊一卡通,项目执行前必须先收费或先记帐审核,不传单据号，根据医嘱ID读取所有未收费单据或未审核的记帐单
                    '读取医嘱ID串
                    str医嘱IDs = mlngCurAdviceId
                    gstrSQL = "Select Id  from 病人医嘱记录 where 相关ID = [1]"
                    Set rsOneCard = zlDatabase.OpenSQLRecord(gstrSQL, "提取医嘱ID串", mlngCurAdviceId)
                    While rsOneCard.EOF = False
                        str医嘱IDs = str医嘱IDs & "," & rsOneCard!ID
                        rsOneCard.MoveNext
                    Wend
                    
                    If mobjSquareCard.zlSquareAffirm(Me, 1294, mstrPrivs, lngPatientID, 0, False, , , str医嘱IDs) = False Then
                        '如果有“未缴费报到”权限，则提示是否确认未收费可以报到？
                        If InStr(mstrPrivs, "未缴费报到") = 1 Then
                            If MsgBoxD(Me, "缴费不成功，该病人还存在未收费的费用，是否继续报到？", vbYesNo, "缴费失败") = vbNo Then
                                AutoRegister = False
                                Exit Function
                            End If
                        Else
                            MsgBoxD Me, "缴费不成功，该病人还存在未收费的费用，无法报到，请检查。", vbOKOnly, "缴费失败"
                            AutoRegister = False
                            Exit Function
                        End If
                    End If
                Else
                    '如果有“未缴费报到”权限，则提示是否确认未收费可以报到？
                    If InStr(mstrPrivs, "未缴费报到") > 0 Then
                        If MsgBoxD(Me, "该病人还存在未收费的费用，是否继续报到？", vbYesNo, "提示信息") = vbNo Then
                            AutoRegister = False
                            Exit Function
                        End If
                    Else
                        MsgBoxD Me, "该病人还存在未收费的费用，请检查。", vbOKOnly, "提示信息"
                        AutoRegister = False
                        Exit Function
                    End If
                End If
            End If
        End If
        
    
    If gbln执行后审核 And int记录性质 = 2 Then
        curMoney = GetAdviceMoney(mlngCurAdviceId, lngSourceType, str类别, str类别名)
        
        '当费用不为0时，检查是否一卡通刷卡，是否需要记账报警
        If curMoney <> 0 Then
            '记账报警
            If Not FinishBillingWarn(Me, "", lngPatientID, rsTemp!主页ID, Val(Nvl(rsTemp!当前病区ID)), curMoney, str类别, str类别名) Then
                AutoRegister = False
                Exit Function
            End If
    
            '问题：34856
            '门诊一卡通消费身份验证
            '参数28--门诊一卡通消费减少剩余款额时是否需要验证
            '参数81--执行后自动审核
            If Val(zlDatabase.GetPara(28, glngSys)) <> 0 And gbln执行后审核 _
                And curMoney > 0 And lngSourceType = 1 Then
                If Not zlDatabase.PatiIdentify(Me, glngSys, lngPatientID, curMoney) Then
                    AutoRegister = False
                    Exit Function
                End If
            End If
        End If
    End If
    
    arrSQL = Array()

    '开始检查
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    
    '影像类别"DG"表示病理
    arrSQL(UBound(arrSQL)) = "ZL_影像检查_BEGIN(Null,Null," & mlngCurAdviceId & "," & rsTemp!发送号 & ",'DG','" & _
        Nvl(rsTemp!姓名, "") & "','','" & Nvl(rsTemp!性别, "") & "','" & _
        Nvl(rsTemp!年龄, "") & "'," & To_Date(Nvl(rsTemp!出生日期, "")) & ",Null,Null,Null,Null,Null,Null,Null,Null,Null," & _
        mstrCurDepartmentId & ")"

    '设置影像检查记录--执行过程为-已报到，报到时处理记账的费用
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_影像检查_State(" & mlngCurAdviceId & "," & rsTemp!发送号 & ",2,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mstrCurDepartmentId & ")"
    
    
    gcnOracle.BeginTrans
    
    blnTran = True
    
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "写入数据")
    Next
    gcnOracle.CommitTrans
    
    Exit Function
ErrHandle:
    AutoRegister = False
    
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function


Private Function IsNewPatholStudy() As Boolean
'返回是否新的病理检查
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    IsNewPatholStudy = True
    
    strSQL = "select 病理号 from 病理检查信息 where 医嘱ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCurAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If Nvl(rsData!病理号) <> "" Then IsNewPatholStudy = False
End Function


Private Function IsHasPatholNum(ByVal strCurPatholNum As String) As Boolean
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select 病理医嘱ID from 病理检查信息 where upper(病理号)=upper([1])"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCurPatholNum)
    
    IsHasPatholNum = IIf(rsData.RecordCount > 0, True, False)
End Function


Private Sub cmdReject_Sure_Click()
On Error GoTo ErrHandle
    '核/拒收标本
    Dim rsPathol As ADODB.Recordset
    Dim i As Integer
    Dim strSQL As String
    Dim strErr As String
    Dim strPatholNum As String

    strErr = CheckSubmitInfoIsValid
    If Trim(strErr) <> "" Then
        Call ShowProcessHint(strErr)
        Exit Sub
    End If

    If mblnIsRejectSpecimen Then
        '拒收标本
        strSQL = "Zl_病理标本_拒收(" & mlngCurAdviceId & ",'" & _
                                    txtUnitName.Text & "','" & _
                                    txtFormDepart.Text & "','" & _
                                    cbxSubmitDoctor.Text & "'," & _
                                    To_Date(dtpSubmitTime.value) & ",'" & _
                                    txtContactWay.Text & "','" & _
                                    txtRegisterDoctor.Text & "','" & _
                                    txtRejectReason.Text & "','" & _
                                    txtRejectNotify.Text & "')"

        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Else
        '如果标本从为核收，则执行自动报到
        If IsNewPatholStudy Then
        
            '判断病理号是否重复
            If IsHasPatholNum(txtPatholNum.Text) Then
                Call MsgBoxD(Me, "病理号重复请修改。", vbInformation, Me.Caption)
                txtPatholNum.SetFocus
                
                Exit Sub
            End If
        
            If Not AutoRegister Then Exit Sub
        End If
        
        '先保存标本信息
        Call mfrmParent.SaveSpecimenData
    
        '核收标本
        strSQL = "Zl_病理标本_核收(" & mlngCurAdviceId & ",'" & _
                                    txtPatholNum.Text & "'," & _
                                    Val(cbxStudyType.Text) & ",'" & _
                                    txtUnitName.Text & "','" & _
                                    txtFormDepart.Text & "','" & _
                                    cbxSubmitDoctor.Text & "'," & _
                                    To_Date(dtpSubmitTime.value) & ",'" & _
                                    txtContactWay.Text & "','" & _
                                    txtRegisterDoctor.Text & "')"

        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        
        If Trim(mstrPatholNum) = "" And mstrPatholInitNum = Trim(txtPatholNum.Text) Then
            '更新病理序号
            strSQL = "ZL_病理号码_序号更新(" & Val(cbxStudyType.Text) & "," & mlngPatholSerialNum & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    End If

    Call mfrmParent.LoadSpecimenAcceptOrRejectHistoryData

    mblnIsSucceed = True
    
    Call Me.Hide

    Exit Sub
ErrHandle:
    Call ShowProcessHint(err.Description)
End Sub



Private Sub Form_Initialize()
    mblnIsSucceed = False
    mblnIsRejectSpecimen = False
    mlngCurAdviceId = 0
    
    Set mtxtAcceptHistory = Nothing
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    
    
    '创建卡结算部件
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    '初始化卡结算部件
    mobjSquareCard.zlInitComponents Me, 1294, glngSys, gstrDBUser, gcnOracle
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call LoadStudyType
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    Set mobjSquareCard = Nothing
End Sub
