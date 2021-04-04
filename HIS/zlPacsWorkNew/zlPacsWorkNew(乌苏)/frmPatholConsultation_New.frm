VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholConsultation_New 
   Caption         =   "会诊申请"
   ClientHeight    =   3420
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   6945
   Icon            =   "frmPatholConsultation_New.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   6945
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cbxConsultationDoctor 
      Height          =   300
      Left            =   4800
      TabIndex        =   2
      Top             =   720
      Width           =   1905
   End
   Begin VB.TextBox txtConsultationUnit 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   1905
   End
   Begin VB.PictureBox picShow 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   3975
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   3975
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
         TabIndex        =   16
         Top             =   120
         Width           =   3735
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   4200
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退 出(&E)"
      Height          =   400
      Left            =   5520
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ComboBox cbxConsultationType 
      Height          =   300
      ItemData        =   "frmPatholConsultation_New.frx":179A
      Left            =   1200
      List            =   "frmPatholConsultation_New.frx":179C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1905
   End
   Begin VB.TextBox txtRequestDoctor 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   240
      Width           =   1905
   End
   Begin VB.TextBox txtDescription 
      Height          =   1095
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   5535
   End
   Begin MSComCtl2.DTPicker dtpStartTime 
      Height          =   300
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   154796035
      CurrentDate     =   40646.4399652778
   End
   Begin MSComCtl2.DTPicker dtpEndTime 
      Height          =   300
      Left            =   4800
      TabIndex        =   4
      Top             =   1200
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   154796035
      CurrentDate     =   40646.4399652778
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "会诊单位："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   18
      Top             =   780
      Width           =   900
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "截止时间："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3840
      TabIndex        =   17
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   240
      Width           =   255
   End
   Begin VB.Label labConsultationDoctor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "会诊医师："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3840
      TabIndex        =   13
      Top             =   780
      Width           =   900
   End
   Begin VB.Label labConsultation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "会诊类型："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   12
      Top             =   300
      Width           =   900
   End
   Begin VB.Label labDescription 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "初步诊断："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   11
      Top             =   1680
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
      Left            =   3840
      TabIndex        =   10
      Top             =   300
      Width           =   900
   End
   Begin VB.Label labRequestTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "会诊时间："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   1260
      Width           =   900
   End
End
Attribute VB_Name = "frmPatholConsultation_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mufgParentGrid As ucFlexGrid

Private mlngPatholAdviceId As Long
Private mlngCurDepartmentId As Long


Public Function ShowConsultationWindow(ufgParentGrid As ucFlexGrid, ByVal lngPatholAdviceId As Long, _
    ByVal lngDepartmentId As Long, owner As Form) As Boolean
'显示制片申请窗口
    Dim curDate As Date
    
    Set mufgParentGrid = ufgParentGrid
    
    mlngPatholAdviceId = lngPatholAdviceId
    mlngCurDepartmentId = lngDepartmentId

    curDate = zlDatabase.Currentdate

    dtpStartTime.value = curDate
    dtpEndTime.value = Format(curDate + 1, "yyyy-mm-dd 23:59:59")
    txtRequestDoctor.Text = UserInfo.姓名

    Call CloseProcessHint
    
    '载入当前科室的医师
    Call LoadConsultationDoctor(lngDepartmentId)

    Call Me.Show(1, owner)
End Function





Private Sub LoadConsultationDoctor(ByVal lngDepartmentId As Long)
'读取会诊医生数据
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select a.姓名 from 人员表 a, 部门人员 b where a.id=b.人员ID and b.部门ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngDepartmentId)
    
    Call cbxConsultationDoctor.Clear
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    Do While Not rsData.EOF
        Call cbxConsultationDoctor.AddItem(rsData!姓名)
        Call rsData.MoveNext
    Loop
    
End Sub


Private Sub LoadConsultationType()
'载入会诊类型
    Call cbxConsultationType.AddItem("0-科内会诊")
    Call cbxConsultationType.AddItem("1-院外会诊")
End Sub


Private Sub ShowProcessHint(ByVal strHint As String)
'显示处理信息
On Error Resume Next

    txtShow.Text = strHint

    picShow.Visible = True
End Sub


Private Sub CloseProcessHint()
'关闭处理提示
    picShow.Visible = False
End Sub


Private Function CheckDataIsValid() As Boolean
'检查数据是否有效
    CheckDataIsValid = True
    
    If cbxConsultationType.Text = "" Then
        CheckDataIsValid = False
        Call ShowProcessHint("请选择合适的会诊类型。")
        
        cbxConsultationType.SetFocus
        
        Exit Function
    End If
End Function



Private Sub SaveConsultationData()
'保存会诊数据
    Dim lngNewRow As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select Zl_病理会诊_新增([1],[2],[3],[4],[5],[6],[7],[8]) as 返回值 from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                            mlngPatholAdviceId, _
                                            txtRequestDoctor.Text, _
                                            txtConsultationUnit.Text, _
                                            cbxConsultationDoctor.Text, _
                                            CDate(dtpStartTime.value), _
                                            CDate(dtpEndTime.value), _
                                            Val(cbxConsultationType.Text), _
                                            txtDescription.Text)
                                            
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "SaveConsultationData", "未成功获取新增后的会诊记录ID,处理失败。")
        Exit Sub
    End If
    
'    mvfgConsultation.mCurFlexGrid.Rows = mvfgConsultation.mCurFlexGrid.Rows + 1
    
    lngNewRow = mufgParentGrid.NewRow
    
    mufgParentGrid.Text(lngNewRow, gstrConsultation_ID) = rsData!返回值
    mufgParentGrid.Text(lngNewRow, gstrConsultation_申请医师) = txtRequestDoctor.Text
    mufgParentGrid.Text(lngNewRow, gstrConsultation_会诊单位) = txtConsultationUnit.Text
    mufgParentGrid.Text(lngNewRow, gstrConsultation_会诊医师) = cbxConsultationDoctor.Text
    mufgParentGrid.Text(lngNewRow, gstrConsultation_会诊时间) = dtpStartTime.value
    mufgParentGrid.Text(lngNewRow, gstrConsultation_截止时间) = dtpEndTime.value
    mufgParentGrid.Text(lngNewRow, gstrConsultation_会诊类型) = GetConsultationTypeValue(Val(cbxConsultationType.Text))
    mufgParentGrid.Text(lngNewRow, gstrConsultation_初步诊断) = txtDescription.Text
    mufgParentGrid.Text(lngNewRow, gstrConsultation_当前状态) = "已申请"
    
    '定位到新增行
    Call mufgParentGrid.LocateRow(lngNewRow)
                                            
End Sub


Private Function GetConsultationTypeValue(ByVal lngConsultationType As Long) As String
'获取会诊类型取值
    Select Case lngConsultationType
        Case 0:
            GetConsultationTypeValue = "科内会诊"
        Case 1:
            GetConsultationTypeValue = "院外会诊"
    End Select

End Function



Private Sub cbxConsultationType_Click()
On Error GoTo errHandle
    txtConsultationUnit.Text = ""
    If cbxConsultationType.Text = "" Then Exit Sub
    
    If Val(cbxConsultationType.Text) = 0 Then txtConsultationUnit.Text = "本科室"
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
    Call Me.Hide
End Sub

Private Sub cmdSure_Click()
'添加会诊记录
On Error GoTo errHandle
    If Not CheckDataIsValid Then Exit Sub
    
    '保存会诊数据
    Call SaveConsultationData
    
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    
    '当编写会诊报告的时候，需要显示在最前面
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶
    
    '载入会诊类型
    Call LoadConsultationType
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub
