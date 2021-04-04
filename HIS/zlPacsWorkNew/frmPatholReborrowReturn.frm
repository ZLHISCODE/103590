VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPatholReborrowReturn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "借阅归还"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10455
   Icon            =   "frmPatholReborrowReturn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtReturnedMoney 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtMemo 
      Height          =   300
      Left            =   1080
      TabIndex        =   19
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox txtEnregMan 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4440
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox rtfAdvice 
      Height          =   855
      Left            =   4320
      TabIndex        =   16
      Top             =   3960
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"frmPatholReborrowReturn.frx":000C
   End
   Begin VB.TextBox txtDoctor 
      Height          =   300
      Left            =   7920
      TabIndex        =   13
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox txtHospital 
      Height          =   300
      Left            =   4320
      TabIndex        =   11
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox txtReturnMoney 
      BackColor       =   &H00E0E0E0&
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
      Height          =   300
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtBorrowMoney 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtHoldMoney 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3720
      TabIndex        =   5
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtpReturnDate 
      Height          =   300
      Left            =   1080
      TabIndex        =   3
      Top             =   3480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   64684035
      CurrentDate     =   40903
   End
   Begin VB.TextBox txtReturnPepole 
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin zl9PACSWork.ucFlexGrid ufgBackDetail 
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   10215
      _ExtentX        =   16325
      _ExtentY        =   5106
      DefaultCols     =   ""
      GridRows        =   201
      IsCopyAdoMode   =   0   'False
      IsEjectConfig   =   -1  'True
      HeadFontCharset =   134
      HeadFontWeight  =   400
      DataFontCharset =   134
      DataFontWeight  =   400
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10200
      TabIndex        =   28
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label Label14 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10200
      TabIndex        =   27
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label13 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6600
      TabIndex        =   26
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label12 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "已退押金："
      Height          =   255
      Left            =   6600
      TabIndex        =   24
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "备    注："
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4020
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "登 记 人："
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "外诊意见："
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   4020
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "外诊医师："
      Height          =   255
      Left            =   6960
      TabIndex        =   14
      Top             =   3540
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "外诊医院："
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   3540
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "需退押金："
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8520
      TabIndex        =   8
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "借阅押金："
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "扣缴押金："
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "归还日期："
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3540
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "归 还 人："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmPatholReborrowReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mufgParentBorrowGrid As ucFlexGrid
Private mlngBorrowId As Long
Public blnIsOk As Boolean


Public Sub ShowBorrowReturnWindow(ufgBorrow As ucFlexGrid, owner As Object)
    If Not ufgBorrow.IsSelectionRow Then
        Call err.Raise(0, "ShowBorrowReturnWindow", "没有选择需要归还的借阅记录。")
        Exit Sub
    End If
    
    Set mufgParentBorrowGrid = ufgBorrow
    
    mlngBorrowId = Val(ufgBorrow.KeyValue(ufgBorrow.SelectionRow))
    
    blnIsOk = False
    
    Call ReadBorrowInfToFace
    Call LoadReturnMaterialData(mlngBorrowId)
    
    Call Me.Show(1, owner)
End Sub

Private Function GetReturnMoney(ByVal lngBorrowId As Long)
'取得已退押金
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetReturnMoney = 0
    
    strSQL = "select sum(退还押金) as 返回值 from 病理归还信息 where 借阅ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBorrowId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetReturnMoney = Val(Nvl(rsData!返回值))
End Function


Private Function CheckDataIsValid() As String
'判断归还的数据是否有效
    Dim i As Long

    CheckDataIsValid = ""
    
    For i = 1 To ufgBackDetail.GridRows - 1
        If ufgBackDetail.GetRowCheck(i) Then
            If Val(ufgBackDetail.Text(i, gstrPatholCol_待还数量)) < Val(ufgBackDetail.Text(i, gstrPatholCol_实还数量)) Then
                CheckDataIsValid = "在材料规划列表中，实还材料数不能大于待还材料数量。"
                
                Call ufgBackDetail.SetFocus
                Call ufgBackDetail.LocateRow(i)
                
                Exit Function
            End If
        End If
    Next i
    
End Function


Private Sub ReadBorrowInfToFace()
'读取相关借阅信息到界面
    txtReturnPepole.Text = mufgParentBorrowGrid.Text(mufgParentBorrowGrid.SelectionRow, gstrPatholCol_借阅人)
    
    txtBorrowMoney.Text = mufgParentBorrowGrid.Text(mufgParentBorrowGrid.SelectionRow, gstrPatholCol_押金)   '借阅押金
    txtReturnedMoney.Text = GetReturnMoney(mlngBorrowId)    '已退押金
    txtHoldMoney.Text = 0   '扣留押金
    txtReturnMoney.Text = Val(txtBorrowMoney.Text) - Val(txtReturnedMoney.Text) '需退押金
    
    txtEnregMan.Text = UserInfo.姓名
    
    dtpReturnDate.value = zlDatabase.Currentdate

End Sub

Private Sub cmdCancel_Click()
'取消归还
On Error Resume Next
    blnIsOk = False
    
    Call Me.Hide
err.Clear
End Sub

Private Sub cmdSure_Click()
On Error GoTo ErrHandle
    Dim strInf As String
    
    If Trim(txtReturnPepole.Text) = "" Then
        Call MsgBoxD(Me, "归还人不能为空。", vbOKOnly, Me.Caption)
        Call txtReturnPepole.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtHospital.Text) = "" Then
        Call MsgBoxD(Me, "外诊医院不能为空。", vbOKOnly, Me.Caption)
        Call txtHospital.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtDoctor.Text) = "" Then
        Call MsgBoxD(Me, "外诊医师不能为空。", vbOKOnly, Me.Caption)
        Call txtDoctor.SetFocus
        
        Exit Sub
    End If
    
    If Trim(rtfAdvice.Text) = "" Then
        Call MsgBoxD(Me, "外诊意见不能为空。", vbOKOnly, Me.Caption)
        Call rtfAdvice.SetFocus
        
        Exit Sub
    End If
    
    If Not ufgBackDetail.IsCheckedRow Then
        If MsgBoxD(Me, "确认不选择任何材料进行归还处理吗？对未勾选的材料，系统将自动作遗失处理。", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    End If
    
    strInf = CheckDataIsValid()
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call MaterialReturnProcess
    
    '更新借阅归还状态
    Call UpdateBorrowReturnState
    
    blnIsOk = True
    
    Call Me.Hide
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub UpdateBorrowReturnState()
'更新借阅归还状态
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim blnFind As Boolean
    Dim chkState As CheckState
    Dim strValue As String
    
    strSQL = "select 归还状态 from 病理借阅信息 where id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngBorrowId)
    
    If rsData.RecordCount > 0 Then
'        Call mufgParentBorrowGrid.GetFieldDisplayText(gstrPatholCol_归还状态, Val(Nvl(rsData!归还状态)), blnFind, chkState, strValue)
        Call mufgParentBorrowGrid.SyncData(mufgParentBorrowGrid.SelectionRow, gstrPatholCol_归还状态, Val(Nvl(rsData!归还状态)), True)
    End If
End Sub

Private Sub MaterialReturnProcess()
'材料归还处理
    Dim i As Long
    Dim lngReturnCount As Long
    
    gcnOracle.BeginTrans
    
    On Error GoTo errTrans
    Call zlDatabase.ExecuteProcedure("ZL_病理归还_新增记录(" & _
                                        mlngBorrowId & ",'" & _
                                        txtReturnPepole.Text & "'," & _
                                        zlStr.To_Date(dtpReturnDate.value) & "," & _
                                        Val(txtReturnMoney.Text) & ",'" & _
                                        txtHospital.Text & "','" & _
                                        txtDoctor.Text & "','" & _
                                        rtfAdvice.Text & "','" & _
                                        txtEnregMan.Text & "','" & _
                                        txtMemo.Text & "')", Me.Caption)
                                        
    For i = 1 To ufgBackDetail.GridRows - 1
        lngReturnCount = IIf(ufgBackDetail.GetRowCheck(i), Val(ufgBackDetail.Text(i, gstrPatholCol_实还数量)), 0)
        
        Call zlDatabase.ExecuteProcedure("ZL_病理归还_材料归还(" & _
                                            mlngBorrowId & "," & _
                                            ufgBackDetail.KeyValue(i) & "," & _
                                            lngReturnCount & "," & _
                                            zlStr.To_Date(dtpReturnDate.value) & ",'" & _
                                            txtEnregMan.Text & "')", Me.Caption)
        
    Next i
    
    gcnOracle.CommitTrans
Exit Sub
errTrans:
    gcnOracle.RollbackTrans
End Sub


Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitBorrowReturnList
    
    blnIsOk = False
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub InitBorrowReturnList()
    '设置行数
    ufgBackDetail.GridRows = glngStandardRowCount
    '设置行高
    ufgBackDetail.RowHeightMin = glngStandardRowHeight
    
    '初始化归还列表
    ufgBackDetail.IsKeepRows = False
    ufgBackDetail.DefaultColNames = gstrMaterialBorrowReturnCols
    ufgBackDetail.ColNames = gstrMaterialBorrowReturnCols
    ufgBackDetail.ColConvertFormat = gstrMaterialBorrowReturnConvertFormat
End Sub


Private Sub LoadReturnMaterialData(ByVal lngBorrowId As Long)
'读取需要归还的材料信息
    Dim strSQL As String
    
    strSQL = "select a.归档id, d.检查类型,d.病理号,c.序号,c.标本名称,c.取材位置, '蜡块' as 材料类别, " & _
             " case when c.申请ID is null then '常规取材' else '补取材' end as 材料明细, " & _
             " (nvl(a.借阅数量, 0) - nvl(a.归还数量, 0)) as 待还数量, (nvl(a.借阅数量, 0) - nvl(a.归还数量, 0)) as 实还数量, a.归还状态, e.档案名称, e.详细地址, " & _
            " '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置 " & _
             " from 病理检查信息 d, 病理取材信息 c, 病理档案信息 e, 病理归档信息 b, 病理借阅关联 a " & _
             " Where c.病理医嘱id = d.病理医嘱id And b.材块id = c.材块id and e.id=b.档案ID And a.归档id = b.ID And b.资料来源 = 1 And a.归还状态<>1 and a.借阅id = [1] " & _
         " Union All " & _
             " select a.归档id, d.检查类型,d.病理号,c.序号,c.标本名称,c.取材位置, '切片' as 材料类别, " & _
             " decode(o.制片方式,0,'正常',1,'重切',2,'深切',3,'连切',4,'白片',5,'重染',6,'薄片','其他') as 材料明细, " & _
             " (nvl(a.借阅数量, 0) - nvl(a.归还数量, 0)) as 待还数量,(nvl(a.借阅数量, 0) - nvl(a.归还数量, 0)) as 实还数量,a.归还状态, e.档案名称, e.详细地址, " & _
            " '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置 " & _
             " from 病理检查信息 d, 病理取材信息 c, 病理制片信息 o, 病理档案信息 e, 病理归档信息 b, 病理借阅关联 a " & _
             " Where c.病理医嘱id = d.病理医嘱id And o.病理医嘱id = c.病理医嘱id " & _
             " and b.制片id = o.id and c.材块id= o.材块id and e.id=b.档案ID and a.归档id=b.id and b.资料来源=2 and a.归还状态<>1 and a.借阅id=[1] " & _
         " Union All " & _
             " select a.归档id, d.检查类型,d.病理号,c.序号,c.标本名称,c.取材位置, " & _
             " decode(o.特检类型,0, '免疫',1,'特染',2,'分子') as 材料类别, " & _
             " decode(o.特检细目,0,decode(o.特检类型,0, '免疫',1,'特染',2,'分子'),1,'鉴别',2,'多耐药',3,'荧光',4,'普通') || '(' || q.抗体名称 || decode(o.制作类型,-1,'-补',0,'','-重' || o.制作类型) || ')' as 材料明细, " & _
             " (nvl(a.借阅数量, 0) - nvl(a.归还数量, 0)) as 待还数量,(nvl(a.借阅数量, 0) - nvl(a.归还数量, 0)) as 实还数量,a.归还状态, e.档案名称, e.详细地址, " & _
            " '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置 " & _
             " from 病理检查信息 d, 病理取材信息 c, 病理抗体信息 q, 病理特检信息 o, 病理档案信息 e, 病理归档信息 b, 病理借阅关联 a " & _
             " Where c.病理医嘱id = d.病理医嘱id And q.抗体ID = o.抗体ID And o.病理医嘱id = c.病理医嘱id " & _
             " and b.特检id = o.id and e.id=b.档案ID and a.归档id=b.id and b.资料来源=3 and a.归还状态<>1 and a.借阅id=[1] "
             
    Set ufgBackDetail.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBorrowId)
    Call ufgBackDetail.RefreshData
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
err.Clear
End Sub

Private Sub txtHoldMoney_Change()
On Error Resume Next
    txtReturnMoney.Text = Val(txtBorrowMoney.Text) - Val(txtReturnedMoney.Text) - Val(txtHoldMoney.Text) '需退押金
err.Clear
End Sub


Private Sub ufgBackDetail_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    If Col <> ufgBackDetail.GetColIndex(gstrPatholCol_实还数量) Then Exit Sub
    
    If Val(ufgBackDetail.Text(Row, gstrPatholCol_实还数量)) < Val(ufgBackDetail.Text(Row, gstrPatholCol_待还数量)) And _
        Val(ufgBackDetail.Text(Row, gstrPatholCol_实还数量)) > 0 Then
        Call ufgBackDetail.SetRowCheck(Row, True)
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

