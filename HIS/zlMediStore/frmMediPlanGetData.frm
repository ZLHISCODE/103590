VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMediPlanGetData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "获取数据"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   Icon            =   "frmMediPlanGetData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame fraSetupParams 
      Caption         =   "选项"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin VB.CheckBox chkData 
         Caption         =   "计划单号(L)"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chkData 
         Caption         =   "审核日期(&V)"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkData 
         Caption         =   "计划库房(&W)"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtPlanNO 
         Height          =   270
         Index           =   0
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtPlanNO 
         Height          =   270
         Index           =   1
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cboWH 
         Height          =   300
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpVerifyDate 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   166723585
         CurrentDate     =   40331
         MaxDate         =   402133
         MinDate         =   36526
      End
      Begin MSComCtl2.DTPicker dtpVerifyDate 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   6
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   166723585
         CurrentDate     =   40331
         MaxDate         =   402133
         MinDate         =   36526
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         Caption         =   "提示：全院表示全院类型的计划单"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   2700
      End
   End
   Begin VB.Label lblRemark 
      BackStyle       =   0  'Transparent
      Caption         =   "注意：当前有获取的数据未导入处理，此时再确定获取数据时将会丢失它！"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   4935
   End
End
Attribute VB_Name = "frmMediPlanGetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'从参数表中取药品价格、数量、金额小数位数（计算精度）
Public mintCostDigit As Integer        '成本价小数位数
Public mintPriceDigit As Integer       '售价小数位数
Public mintNumberDigit As Integer      '数量小数位数
Public mintMoneyDigit As Integer       '金额小数位数
Private mintUnit As Integer             '单位系数：1-售价;2-门诊;3-住院;4-药库
Private mstrWhere As String

Property Get SQLWhere() As String
    SQLWhere = mstrWhere
End Property
Property Let SQLWhere(ByVal strVal As String)
    mstrWhere = strVal
End Property



Private Sub chkData_Click(Index As Integer)
    Call SwitchParamsState
End Sub

Private Sub cmdGetData_Click()
    Dim strWhere As String
    
    If chkData(0).Value <> 1 And chkData(1).Value <> 1 And chkData(2).Value <> 1 Then Exit Sub

    If chkData(0).Value = 1 Then
        '补充单据号
        If Len(Trim(txtPlanNO(0).Text)) < 8 And Len(Trim(txtPlanNO(0).Text)) > 0 Then
            txtPlanNO(0).Text = zlCommFun.GetFullNO(txtPlanNO(0).Text, 32, cboWH.ItemData(cboWH.ListIndex))
        End If
        If Len(Trim(txtPlanNO(1).Text)) < 8 And Len(Trim(txtPlanNO(1).Text)) > 0 Then
            txtPlanNO(1).Text = zlCommFun.GetFullNO(txtPlanNO(1).Text, 32, cboWH.ItemData(cboWH.ListIndex))
        End If
    
        If Len(Trim(txtPlanNO(0).Text)) > 0 And Len(Trim(txtPlanNO(1).Text)) > 0 Then
            strWhere = strWhere & " and NO between '" & Trim(txtPlanNO(0).Text) & "' and '" & Trim(txtPlanNO(1).Text) & "'"
        Else
            MsgBox "请录入'计划单号'！", , gstrSysName
            txtPlanNO(0).SetFocus
            Exit Sub
        End If
    End If
    '审核日期
    If chkData(1).Value = 1 Then
        strWhere = strWhere & " and 审核日期 between to_date('" & Me.dtpVerifyDate(0).Value & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss')" _
               & " and to_date('" & Me.dtpVerifyDate(1).Value & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
    End If
    '库房
    If chkData(2).Value = 1 Then
        If cboWH.Text <> "全院" Then
            If cboWH.ListIndex >= 0 Then
                strWhere = strWhere & " and 库房ID=" & cboWH.ItemData(cboWH.ListIndex)
            End If
        Else
            strWhere = strWhere & " and nvl(库房ID,0)=0 "
        End If
    End If
    
    SQLWhere = strWhere
    Me.Hide

End Sub

Private Sub cmdCancel_Click()
    SQLWhere = ""
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    Call SwitchParamsState
    Call SetWarehouse
    dtpVerifyDate(0).Value = Date - 7
    dtpVerifyDate(1).Value = Date
End Sub

Private Sub SwitchParamsState()
'fraSetupParams内控件的状态控制
    Const COLOR_DISABLE = &H8000000F
    Dim i As Integer
    
    For i = Me.chkData.UBound To Me.chkData.LBound Step -1
        Select Case Me.chkData(i).Index
            Case 2
                If Me.chkData(i).Value = 1 Then
                    Me.cboWH.Enabled = True
                    Me.cboWH.BackColor = vbWhite
                Else
                    Me.cboWH.Enabled = False
                    Me.cboWH.BackColor = COLOR_DISABLE
                End If
            Case 1
                If Me.chkData(i).Value = 1 Then
                    Me.dtpVerifyDate(0).Enabled = True
                    Me.dtpVerifyDate(1).Enabled = True
                Else
                    Me.dtpVerifyDate(0).Enabled = False
                    Me.dtpVerifyDate(1).Enabled = False
                End If
            Case Else
                If Me.chkData(i).Value = 1 Then
                    Me.txtPlanNO(0).Enabled = True
                    Me.txtPlanNO(1).Enabled = True
                    Me.txtPlanNO(0).BackColor = vbWhite
                    Me.txtPlanNO(1).BackColor = vbWhite
                Else
                    Me.txtPlanNO(0).Enabled = False
                    Me.txtPlanNO(1).Enabled = False
                    Me.txtPlanNO(0).BackColor = COLOR_DISABLE
                    Me.txtPlanNO(1).BackColor = COLOR_DISABLE
                End If
        End Select
    Next
End Sub

Private Sub SetWarehouse(Optional ByVal bln指定 As Boolean)
'设置库房、ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim cboTmp As ComboBox
    
    On Error GoTo ErrHand
    

    Set cboTmp = cboWH
    strSQL = "select cast(0  as integer) 库房ID, cast('全院' as varchar2(50)) 库房 from dual" _
           & " union all " _
           & "Select distinct a.库房id, b.名称 From 药品采购计划 a, 部门表 b where a.库房id=b.id"

    cboTmp.Clear
    Set rsTmp = zlDataBase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 0 To rsTmp.RecordCount - 1
            cboTmp.AddItem rsTmp!库房
            cboTmp.ItemData(i) = rsTmp!库房id
            rsTmp.MoveNext
        Next
        cboTmp.ListIndex = 0
    End If
    rsTmp.Close
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txtPlanNO_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Index = 0 Then
        If Len(Trim(txtPlanNO(0).Text)) < 8 And Len(Trim(txtPlanNO(0).Text)) > 0 Then
            txtPlanNO(0).Text = zlCommFun.GetFullNO(txtPlanNO(0).Text, 32, cboWH.ItemData(cboWH.ListIndex))
        End If
        
    Else
        If Len(Trim(txtPlanNO(1).Text)) < 8 And Len(Trim(txtPlanNO(1).Text)) > 0 Then
            txtPlanNO(1).Text = zlCommFun.GetFullNO(txtPlanNO(1).Text, 32, cboWH.ItemData(cboWH.ListIndex))
        End If
        
    End If
End Sub

Private Sub txtPlanNO_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub
