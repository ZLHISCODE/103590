VERSION 5.00
Begin VB.Form frmPatholRequisition_View 
   Caption         =   "申请查看"
   ClientHeight    =   6390
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   9075
   Icon            =   "frmPatholRequisition_View.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   9075
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   7680
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame framRequest 
      Caption         =   "申请记录"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   9128
         GridRows        =   21
         BackColor       =   12648447
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         Editable        =   0
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
End
Attribute VB_Name = "frmPatholRequisition_View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnMoved As Boolean


Public Sub ShowRequestViewWind(ByVal lngPatholAdviceId As Long, ByVal lngRequestType As Long, _
    ByVal blnMoved As Boolean, owner As Form)
'显示申请查看窗口
    mblnMoved = blnMoved
    
    Call LoadRequestViewData(lngPatholAdviceId, lngRequestType)
    
    Select Case lngRequestType
        Case 0
            Me.Caption = "免疫申请查看"
        Case 1
            Me.Caption = "特染申请查看"
        Case 2
            Me.Caption = "分子申请查看"
        Case 3
            Me.Caption = "制片申请查看"
        Case 4
            Me.Caption = "取材申请查看"
    End Select
    
    Call Me.Show(1, owner)
End Sub



Private Sub InitRequestList()
'初始化申请查看列表
     Dim strTemp As String
     

    
    '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("申请查看列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgData.DefaultColNames = gstrRequisitionViewCols
     
    If strTemp = "" Then
        ufgData.ColNames = gstrRequisitionViewCols
    Else
        ufgData.ColNames = strTemp
    End If
         '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.ColConvertFormat = gstrRequisitionConvertFormat
End Sub



Private Sub LoadRequestViewData(ByVal lngPatholAdviceId As Long, ByVal lngRequestType As Long)
'载入申请信息
    Dim strSql As String
    
    strSql = "select 申请ID,申请人,申请时间,申请细目,申请状态,申请描述,补费状态,完成时间 from 病理申请信息 where 病理医嘱ID=[1] and 申请类型=[2] order by 申请时间"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId, lngRequestType)
    
    Call ufgData.RefreshData
End Sub


Private Sub AdjustFace()
    framRequest.Left = 120
    framRequest.Top = 120
    framRequest.Width = Me.Width - 360
    framRequest.Height = Me.Height - cmdSure.Height - 900
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framRequest.Width - 240
    ufgData.Height = framRequest.Height - 360
    
    cmdSure.Left = Me.Width - cmdSure.Width - 240
    cmdSure.Top = Me.Height - cmdSure.Height - 620
End Sub


Private Sub cmdSure_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitRequestList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    '关闭窗口时保存列表配置
     zlDatabase.SetPara "申请查看列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
     
End Sub
