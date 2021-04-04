VERSION 5.00
Begin VB.Form frmRegistPlan 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病人转诊"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "frmRegistPlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5385
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   5385
      TabIndex        =   9
      Top             =   0
      Width           =   5385
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -45
         X2              =   5500
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请指定病人所要转诊到的目标科室等信息。"
         Height          =   180
         Left            =   600
         TabIndex        =   11
         Top             =   390
         Width           =   3420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转诊信息"
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
         TabIndex        =   10
         Top             =   135
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   4350
         Picture         =   "frmRegistPlan.frx":058A
         Top             =   45
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -150
      TabIndex        =   8
      Top             =   2700
      Width           =   6900
   End
   Begin VB.ComboBox cbo医生 
      Height          =   300
      Left            =   1905
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   2025
   End
   Begin VB.ComboBox cbo诊室 
      Height          =   300
      Left            =   1905
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2010
      Width           =   2025
   End
   Begin VB.ComboBox cbo科室 
      Height          =   300
      ItemData        =   "frmRegistPlan.frx":20CC
      Left            =   1905
      List            =   "frmRegistPlan.frx":20CE
      TabIndex        =   0
      Top             =   1125
      Width           =   2025
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3660
      TabIndex        =   4
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2445
      TabIndex        =   3
      Top             =   2865
      Width           =   1100
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "转诊医生"
      Height          =   180
      Left            =   1125
      TabIndex        =   7
      Top             =   1620
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "转诊诊室"
      Height          =   180
      Left            =   1125
      TabIndex        =   6
      Top             =   2070
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "转诊科室"
      Height          =   180
      Left            =   1125
      TabIndex        =   5
      Top             =   1185
      Width           =   720
   End
End
Attribute VB_Name = "frmRegistPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrNO As String

Private mlng科室ID As Long
Private mstr诊室 As String
Private mstr医生 As String
Private mlng医生ID As Long

Private mstr原诊室 As String
Private mlng原科室ID As Long

Private mlngPreDept As Long
Private mstrLike As String
Private mblnOK As Boolean

Public Function ShowMe(frmParent As Object, ByVal strNO As String, _
    lng科室ID As Long, str诊室 As String, str医生 As String, lng医生ID As Long) As Boolean
'参数：strNO=要转诊的挂号单
'返回：转诊号别,转诊科室,转诊诊室,转诊医生信息
    
    mstrNO = strNO
    Me.Show 1, frmParent
    
    If mblnOK Then
        lng科室ID = mlng科室ID
        str诊室 = mstr诊室
        str医生 = mstr医生
        lng医生ID = mlng医生ID
    End If
    ShowMe = mblnOK
End Function

Private Sub cbo科室_Click()
    If cbo科室.ListIndex <> -1 Then
        If mlngPreDept <> cbo科室.ItemData(cbo科室.ListIndex) Then
            mlngPreDept = cbo科室.ItemData(cbo科室.ListIndex)
            '读取该科室医生、诊室
            Call LoadDoctor
            Call LoadRoom
        End If
    Else
        mlngPreDept = 0
    End If
End Sub

Private Sub cbo科室_GotFocus()
    Call zlControl.TxtSelAll(cbo科室)
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
        
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If cbo科室.Text <> "" Then
            strSql = "Select B.ID,B.编码,B.名称" & _
                " From 部门表 B,部门性质说明 C" & _
                " Where B.ID=C.部门ID And C.服务对象 In(1,3) And C.工作性质='临床'" & _
                " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null)" & _
                " And (B.编码 Like [1] Or Upper(B.简码) Like [2] Or Upper(B.名称) Like [2])" & _
                " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
                " Order by B.编码"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(cbo科室.Text) & "%", mstrLike & UCase(cbo科室.Text) & "%")
            If Not rsTmp.EOF Then
                Call Cbo.SeekIndex(cbo科室, rsTmp!ID)
            Else
                Call Cbo.SeekIndex(cbo科室, mlngPreDept)
            End If
            Call ZLCommFun.PressKey(vbKeyTab)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo医生_Click()
    Call LoadRoom
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '检查设置
    If cbo科室.ListIndex = -1 Then
        MsgBox "请确定要转诊的科室。", vbInformation, gstrSysName
        cbo科室.SetFocus: Exit Sub
    End If
'    If cbo诊室.Text = "" And cbo医生.Text = "" Then
'        MsgBox "请指明转诊诊室或医生。", vbInformation, gstrSysName
'        If cbo诊室.Enabled Then cbo诊室.SetFocus
'        Exit Sub
'    End If
    If cbo科室.ItemData(cbo科室.ListIndex) = mlng原科室ID Then
        If cbo医生.Text = "" Then
            MsgBox "病人在原科室内转诊时，请指明转诊医生。", vbInformation, gstrSysName
            If cbo医生.Enabled Then cbo医生.SetFocus
            Exit Sub
        End If
        If ZLCommFun.GetNeedName(cbo医生.Text) = UserInfo.姓名 Then
            MsgBox "病人在原科室内转诊时，转诊医生应该为其他医生。", vbInformation, gstrSysName
            If cbo医生.Enabled Then cbo医生.SetFocus
            Exit Sub
        End If
    End If
    
    '返回数据
    mlng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    mstr诊室 = cbo诊室.Text
    mstr医生 = ZLCommFun.GetNeedName(cbo医生.Text)
    If cbo医生.ListIndex <> -1 Then
        mlng医生ID = cbo医生.ItemData(cbo医生.ListIndex)
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not Me.ActiveControl Is cbo科室 Then
            KeyAscii = 0
            Call ZLCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    mlng科室ID = 0
    mstr诊室 = ""
    mstr医生 = ""
    mlng医生ID = 0
    mblnOK = False
    mlngPreDept = 0
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "") '输入匹配方式
    
    On Error GoTo errH
    
    '原挂号相关信息
    strSql = "Select 执行部门ID,诊室 From 病人挂号记录 Where NO=[1] And 记录性质=1 And 记录状态=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstrNO)
    mstr原诊室 = Nvl(rsTmp!诊室)
    mlng原科室ID = rsTmp!执行部门ID
    
    '读取门诊科室:缺省为本科室
    strSql = "Select Distinct B.ID,B.编码,B.名称,Decode(B.ID,[1],1,0) as 缺省" & _
        " From 部门表 B,部门性质说明 C" & _
        " Where B.ID=C.部门ID And C.服务对象 In(1,3) And C.工作性质='临床'" & _
        " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null)" & _
        " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
        " Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng原科室ID)
    Do While Not rsTmp.EOF
        cbo科室.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!ID
        If rsTmp!缺省 Then
            cbo科室.ListIndex = cbo科室.NewIndex '主动激活Click
            mlngPreDept = rsTmp!ID
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDoctor()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
                
    cbo医生.Clear
    If cbo科室.ListIndex = -1 Then Exit Sub
    
    strSql = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
        " And C.人员性质='医生' And B.部门ID=[1]" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.简码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, cbo科室.ItemData(cbo科室.ListIndex))
    
    cbo医生.AddItem ""
    Call Cbo.SetIndex(cbo医生.hwnd, 0)
    Do While Not rsTmp.EOF
        cbo医生.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cbo医生.ItemData(cbo医生.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadRoom()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim bytRegistMode As Byte, datRegistTime As Date
    
    On Error GoTo errH
    
    cbo诊室.Clear
    If cbo科室.ListIndex = -1 Then Exit Sub
    
    bytRegistMode = Val(Split(zlDatabase.GetPara("挂号排班模式", glngSys) & "|", "|")(0))
    If Split(zlDatabase.GetPara("挂号排班模式", glngSys) & "|", "|")(1) <> "" Then
        datRegistTime = CDate(Format(Split(zlDatabase.GetPara("挂号排班模式", glngSys) & "|", "|")(1), "yyyy-mm-dd hh:mm:ss"))
    End If
    
    If bytRegistMode = 0 Then
        strSql = "Select ID From 挂号安排 Where 科室ID=[1] And (医生姓名=[2] Or 医生姓名 Is Null Or [2] Is Null)"
        strSql = "Select Distinct 门诊诊室 From 挂号安排诊室 Where 号表ID IN(" & strSql & ") Order by 门诊诊室"
    Else
        If Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") < Format(datRegistTime, "yyyy-mm-dd hh:mm:ss") Then
            strSql = "Select ID From 挂号安排 Where 科室ID=[1] And (医生姓名=[2] Or 医生姓名 Is Null Or [2] Is Null)"
            strSql = "Select Distinct 门诊诊室 From 挂号安排诊室 Where 号表ID IN(" & strSql & ") Order by 门诊诊室"
        Else
            strSql = "Select A.ID From 临床出诊记录 A Where A.科室ID=[1] And (A.医生姓名=[2] Or A.医生姓名 Is Null Or [2] Is Null)"
            strSql = "Select Distinct B.名称 As 门诊诊室 From 临床出诊诊室记录 A,门诊诊室 B Where A.诊室ID=B.ID And A.记录ID IN(" & strSql & ") Order by B.名称"
        End If
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, cbo科室.ItemData(cbo科室.ListIndex), ZLCommFun.GetNeedName(cbo医生.Text))
    
    cbo诊室.AddItem ""
    Call Cbo.SetIndex(cbo诊室.hwnd, 0)
    Do While Not rsTmp.EOF
        cbo诊室.AddItem rsTmp!门诊诊室
        If cbo科室.ItemData(cbo科室.ListIndex) = mlng原科室ID And rsTmp!门诊诊室 = mstr原诊室 Then
            cbo诊室.ListIndex = cbo诊室.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
