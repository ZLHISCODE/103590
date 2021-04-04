VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholRequisition_SupMaterial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "补取材申请"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5340
   Icon            =   "frmPatholRequisition_SupMaterial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtRequestDoctor 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   3945
   End
   Begin VB.TextBox txtDescription 
      Height          =   975
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   2640
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退 出(&E)"
      Height          =   400
      Left            =   3960
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpRequestTime 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   160038915
      CurrentDate     =   40646.4399652778
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
      TabIndex        =   7
      Top             =   1200
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
      Left            =   240
      TabIndex        =   6
      Top             =   780
      Width           =   900
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
      Top             =   300
      Width           =   900
   End
End
Attribute VB_Name = "frmPatholRequisition_SupMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private mufgParentRequest As ucFlexGrid
Private mufgParentContext As ucFlexGrid

Private mlngPatholAdviceId As Long

Private mfrmOwner As Form


Public blnIsOk As Boolean


Public Function ShowSupMaterialWindow(ufgParentRequestGrid As ucFlexGrid, ufgParentContextGrid As ucFlexGrid, _
    ByVal lngPatholAdviceId As Long, owner As Form) As Boolean
'显示补取材申请窗口
    Set mufgParentRequest = ufgParentRequestGrid
    Set mufgParentContext = ufgParentContextGrid
    
    Set mfrmOwner = owner

    mlngPatholAdviceId = lngPatholAdviceId
    
    blnIsOk = False


    dtpRequestTime.value = zlDatabase.Currentdate
    txtRequestDoctor.Text = UserInfo.姓名

    Call Me.Show(1, owner)
End Function



Private Sub cmdExit_Click()
On Error Resume Next
    blnIsOk = False
    Call Me.Hide
End Sub


Private Sub SaveSupMaterialRequest()
'保存补取材申请
    Dim lngNewRow As Long
    
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    

    '添加检查申请信息
    strSql = "select Zl_病理申请_新增([1],[2],[3],[4],[5],[6],[7]) as 返回值 from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                            mlngPatholAdviceId, _
                                            txtRequestDoctor.Text, _
                                            CDate(dtpRequestTime.value), _
                                            4, 0, _
                                            0, _
                                            txtDescription.Text)
                                            
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "SaveSpeExamRequest", "未成功获取新增后的申请ID,处理失败。")
        Exit Sub
    End If

    '设置界面信息
    lngNewRow = mufgParentRequest.NewRow
    
    mufgParentRequest.Text(lngNewRow, gstrRequisition_申请ID) = rsData!返回值
    mufgParentRequest.Text(lngNewRow, gstrRequisition_申请人) = txtRequestDoctor.Text
    mufgParentRequest.Text(lngNewRow, gstrRequisition_申请类型) = "补取材"
    mufgParentRequest.Text(lngNewRow, gstrRequisition_申请时间) = dtpRequestTime.value
    mufgParentRequest.Text(lngNewRow, gstrRequisition_申请描述) = txtDescription.Text
    mufgParentRequest.Text(lngNewRow, gstrRequisition_当前状态) = "已申请"
                                            
    
    '定位到新增行
    Call mufgParentRequest.LocateRow(lngNewRow)

End Sub




Private Sub cmdSure_Click()
'确认申请
On Error GoTo errHandle
    
    '判断申请明细列表是否为制片项目明细列表
    If mufgParentContext.GetColIndex(gstrRequest_Material_取材时间) < 0 Then
        mufgParentContext.ColNames = gstrRequest_Material_Cols
        mufgParentContext.ColConvertFormat = gstrRequest_MaterialConvertFormat
        
        '切换主界面的控制界面
        Call mfrmOwner.ChangeControlFace(4)
    End If
    
    '保存取材申请
    Call SaveSupMaterialRequest
    
    blnIsOk = True
    
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub
