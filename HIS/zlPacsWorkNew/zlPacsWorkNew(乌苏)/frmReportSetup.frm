VERSION 5.00
Begin VB.Form frmReportSetup 
   BorderStyle     =   0  'None
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraReportSetup 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.Frame fraEditorSetUp 
         Caption         =   "报告文档编辑器设置"
         Height          =   5535
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   7695
         Begin VB.Frame Frame8 
            Caption         =   "查看历史报告"
            Height          =   1215
            Left            =   240
            TabIndex        =   35
            Top             =   480
            Width           =   7215
            Begin VB.OptionButton optHistoryReportEditor 
               Caption         =   "PACS报告编辑器"
               Height          =   255
               Index           =   1
               Left            =   4080
               TabIndex        =   37
               Top             =   600
               Width           =   1695
            End
            Begin VB.OptionButton optHistoryReportEditor 
               Caption         =   "电子病历编辑器"
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   36
               Top             =   600
               Value           =   -1  'True
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "报告编辑器"
         Height          =   615
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   7730
         Begin VB.OptionButton optReportEditor 
            Caption         =   "报告文档编辑器"
            Height          =   255
            Index           =   2
            Left            =   5640
            TabIndex        =   33
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optReportEditor 
            Caption         =   "电子病历编辑器"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   31
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optReportEditor 
            Caption         =   "PACS报告编辑器"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   30
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "报告设置"
         Height          =   4575
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   7695
         Begin VB.CheckBox chkUntreadPrinted 
            Caption         =   "审核打印后允许回退"
            Height          =   180
            Left            =   600
            TabIndex        =   32
            Top             =   1800
            Width           =   2175
         End
         Begin VB.CheckBox chkSpecialContent 
            Caption         =   "显示专科报告内容："
            Height          =   180
            Left            =   600
            TabIndex        =   28
            Top             =   2280
            Width           =   2055
         End
         Begin VB.ComboBox cboSpecialContent 
            Height          =   300
            Left            =   600
            TabIndex        =   27
            Text            =   "Combo1"
            Top             =   2760
            Width           =   6495
         End
         Begin VB.CheckBox chkExitAfterPrint 
            Caption         =   "打印后退出"
            Height          =   180
            Left            =   600
            TabIndex        =   26
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            Caption         =   "报告文本段名称"
            Height          =   1335
            Left            =   3960
            TabIndex        =   19
            Top             =   960
            Width           =   3255
            Begin VB.TextBox txtAdvice 
               Height          =   270
               Left            =   1560
               TabIndex        =   22
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox txtResult 
               Height          =   270
               Left            =   1560
               TabIndex        =   21
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox txtCheckView 
               Height          =   270
               Left            =   1560
               TabIndex        =   20
               Top             =   225
               Width           =   1335
            End
            Begin VB.Label Label3 
               Caption         =   "建    议："
               Height          =   255
               Left            =   360
               TabIndex        =   25
               Top             =   975
               Width           =   975
            End
            Begin VB.Label Label2 
               Caption         =   "诊断意见："
               Height          =   255
               Left            =   360
               TabIndex        =   24
               Top             =   615
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "检查所见："
               Height          =   255
               Left            =   360
               TabIndex        =   23
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.CheckBox chkShowVideoCapture 
            Caption         =   "显示视频采集区域"
            Height          =   180
            Left            =   600
            TabIndex        =   18
            Top             =   840
            Width           =   2055
         End
         Begin VB.Frame frmShowBigImg 
            Height          =   735
            Left            =   480
            TabIndex        =   14
            Top             =   3480
            Width           =   6735
            Begin VB.OptionButton optBigImgAction 
               Caption         =   "鼠标移动时显示大图，放大倍数为："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Value           =   -1  'True
               Width           =   3255
            End
            Begin VB.OptionButton optBigImgAction 
               Caption         =   "单击显示大图窗口"
               Height          =   180
               Index           =   2
               Left            =   120
               TabIndex        =   16
               Top             =   480
               Width           =   1815
            End
            Begin VB.ComboBox cboZoom 
               Height          =   300
               ItemData        =   "frmReportSetup.frx":0000
               Left            =   3360
               List            =   "frmReportSetup.frx":0010
               TabIndex        =   15
               Text            =   "1"
               Top             =   200
               Width           =   855
            End
         End
         Begin VB.CheckBox chkShowBigImg 
            Caption         =   "显示大图："
            Height          =   300
            Left            =   600
            TabIndex        =   13
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtMinImageCount 
            Height          =   270
            Left            =   6240
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "8"
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox chkShowImage 
            Caption         =   "显示报告图像区域                   报告缩略图显示数量："
            Height          =   180
            Left            =   600
            TabIndex        =   11
            Top             =   420
            Width           =   5415
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "报告词句双击后"
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   5640
         Width           =   2415
         Begin VB.OptionButton optWordDblClick 
            Caption         =   "直接写入报告"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optWordDblClick 
            Caption         =   "打开词句编辑窗口"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   480
            Width           =   1750
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "缩略图双击后"
         Height          =   855
         Left            =   2520
         TabIndex        =   4
         Top             =   5640
         Width           =   2895
         Begin VB.OptionButton optImageDblClick 
            Caption         =   "打开图片编辑窗口"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   6
            Top             =   480
            Width           =   1750
         End
         Begin VB.OptionButton optImageDblClick 
            Caption         =   "直接写入报告"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "词句模板显示"
         Height          =   855
         Left            =   5400
         TabIndex        =   1
         Top             =   5640
         Width           =   2415
         Begin VB.OptionButton optShowWord 
            Caption         =   "双击标题"
            Height          =   180
            Index           =   1
            Left            =   360
            TabIndex        =   3
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optShowWord 
            Caption         =   "直接显示"
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmReportSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDeptID As Long   '科室ID
Private mblnRefreshed As Boolean

Public Sub zlRefresh(lngDeptID As Long)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngTemp As Long
    
    mblnRefreshed = True            '数据被刷新过了，可以保存
    
    mlngDeptID = lngDeptID
    optReportEditor(0).value = True '默认使用电子病历编辑器编辑报告
    chkShowImage.value = 0          '默认不显示图像区域
    chkShowVideoCapture.value = 0   '默认不显示视频采集区域
    chkShowBigImg.value = 0         '默认鼠标移动时不显示大图
    optBigImgAction(1).value = True '默认鼠标移动的时候显示大图
    frmShowBigImg.Enabled = False   '默认鼠标移动时不显示大图
    
    chkSpecialContent.value = 0     '默认不显示专科报告
    cboSpecialContent.Enabled = False
    cboZoom.Text = 1                '默认放大倍数为1
    chkExitAfterPrint.value = 0     '默认打印后不退出
    optWordDblClick(0).value = True '默认双击词句后直接写入报告
    optImageDblClick(0).value = True '默认报告缩略图双击后直接写入报告
    txtCheckView.Text = "检查所见"  '默认为检查所见
    txtResult.Text = "诊断意见"     '默认为诊断意见
    txtAdvice.Text = "建议"         '默认为建议
    optShowWord(0).value = True     '默认为直接显示词句模板
    chkUntreadPrinted.value = 0     '默认为审核打印后不允许回退
     
    On Error GoTo err
    strSql = "select ID ,科室ID,参数名,参数值 from 影像流程参数 where 科室ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    
    While Not rsTemp.EOF
        Select Case rsTemp!参数名
            Case "报告编辑器"
                If Nvl(rsTemp!参数值, 0) = 0 Then
                    optReportEditor(0).value = True
                ElseIf Nvl(rsTemp!参数值, 0) = 1 Then
                    optReportEditor(1).value = True
                Else
                    optReportEditor(2).value = True
                End If
            Case "查看历史报告"
                If Nvl(rsTemp!参数值, 0) = 0 Then
                    optHistoryReportEditor(0).value = True
                Else
                    optHistoryReportEditor(1).value = True
                End If
                
            Case "显示报告图像"
                chkShowImage.value = Nvl(rsTemp!参数值, 0)
            Case "报告缩略图数量"
                txtMinImageCount.Text = Nvl(rsTemp!参数值, "8")
            Case "显示视频采集"
                chkShowVideoCapture.value = Nvl(rsTemp!参数值, 0)
            Case "打印后退出"
                chkExitAfterPrint.value = Nvl(rsTemp!参数值, 0)
            Case "报告中显示大图"
                lngTemp = Nvl(rsTemp!参数值, 0)
                If lngTemp = 0 Then
                    chkShowBigImg.value = 0
                ElseIf lngTemp = 1 Then
                    chkShowBigImg.value = 1
                    optBigImgAction(1).value = True
                Else
                    chkShowBigImg.value = 1
                    optBigImgAction(2).value = True
                End If
                frmShowBigImg.Enabled = IIf(chkShowBigImg.value = 1, True, False)
            Case "报告大图放大倍数"
                cboZoom.Text = Nvl(rsTemp!参数值, 1)
                If Val(cboZoom.Text) = 0 Then cboZoom.Text = 1
            Case "显示专科报告"
                chkSpecialContent.value = Nvl(rsTemp!参数值, 0)
                cboSpecialContent.Enabled = IIf(chkSpecialContent.value = 1, True, False)
            Case "专科报告页"
                cboSpecialContent.Text = Nvl(rsTemp!参数值)
            Case "报告词句双击操作"
                If Nvl(rsTemp!参数值, 0) = 0 Then
                    optWordDblClick(0).value = True
                Else
                    optWordDblClick(1).value = True
                End If
            Case "缩略图双击操作"
                If Nvl(rsTemp!参数值, 0) = 0 Then
                    optImageDblClick(0).value = True
                Else
                    optImageDblClick(1).value = True
                End If
            Case "检查所见名称"
                txtCheckView.Text = Nvl(rsTemp!参数值, "检查所见")
            Case "诊断意见名称"
                txtResult.Text = Nvl(rsTemp!参数值, "诊断意见")
            Case "建议名称"
                txtAdvice.Text = Nvl(rsTemp!参数值, "建议")
            Case "显示词句示范"
                If Nvl(rsTemp!参数值, 0) = 0 Then
                    optShowWord(0).value = True
                Else
                    optShowWord(1).value = True
                End If
            Case "审核打印后允许回退"
                chkUntreadPrinted.value = Nvl(rsTemp!参数值, 0)
        End Select
        rsTemp.MoveNext
    Wend
    
    If optReportEditor(2).value Then
        fraEditorSetUp.Visible = True
        
    Else
        fraEditorSetUp.Visible = False
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub


Public Sub zlSave()
    Dim intMatch As Integer
    Dim strSql As String
    
    On Error GoTo errHand
    
    If mblnRefreshed = False Then Exit Sub          '数据没有被刷新，所以不保存
    
    If optReportEditor(0).value = True Then         '电子病历编辑器
        intMatch = 0
    ElseIf optReportEditor(1).value = True Then     'PACS报告编辑器
        intMatch = 1
    ElseIf optReportEditor(2).value = True Then     '报告文档编辑器
        intMatch = 2
    End If
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '报告编辑器','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '显示报告图像','" & chkShowImage.value & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '报告缩略图数量','" & txtMinImageCount.Text & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '显示视频采集','" & chkShowVideoCapture.value & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '打印后退出','" & chkExitAfterPrint.value & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    If chkShowBigImg.value = 0 Then
        intMatch = 0
    ElseIf optBigImgAction(1).value = True Then
        intMatch = 1
    Else
        intMatch = 2
    End If
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '报告中显示大图','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption

    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '报告大图放大倍数','" & IIf(Val(cboZoom.Text) = 0, 1, Val(cboZoom.Text)) & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '显示专科报告','" & chkSpecialContent.value & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '专科报告页','" & cboSpecialContent.Text & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    If optWordDblClick(0).value = True Then         '报告词句双击后直接写入报告
        intMatch = 0
    ElseIf optWordDblClick(1).value = True Then     '报告词句双击后打开编辑窗口
        intMatch = 1
    End If
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '报告词句双击操作','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    If optImageDblClick(0).value = True Then         '缩略图双击后直接写入报告
        intMatch = 0
    ElseIf optImageDblClick(1).value = True Then     '缩略图双击后打开图像编辑窗口
        intMatch = 1
    End If
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '缩略图双击操作','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '检查所见名称','" & txtCheckView.Text & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '诊断意见名称','" & txtResult.Text & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '建议名称','" & txtAdvice.Text & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    If optShowWord(0).value = True Then         '直接显示词句示范
        intMatch = 0
    ElseIf optShowWord(1).value = True Then     '双击标题后显示词句示范
        intMatch = 1
    End If
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '显示词句示范','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '审核打印后允许回退','" & chkUntreadPrinted.value & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    If optReportEditor(2) Then
        strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '查看历史报告','" & IIf(optHistoryReportEditor(0).value, 0, 1) & "')"
        zlDatabase.ExecuteProcedure strSql, Me.Caption
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub chkShowBigImg_Click()
    frmShowBigImg.Enabled = IIf(chkShowBigImg.value = 1, True, False)
End Sub

Private Sub chkSpecialContent_Click()
    If chkSpecialContent.value = 1 Then
        cboSpecialContent.Enabled = True
    Else
        cboSpecialContent.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    mblnRefreshed = False
    '装载专科报告名称
    cboSpecialContent.Clear
    cboSpecialContent.AddItem (Report_Form_frmReportES)
    cboSpecialContent.AddItem (Report_Form_frmReportPathology)
    cboSpecialContent.AddItem (Report_Form_frmReportUS)
    cboSpecialContent.AddItem (Report_Form_frmReportCustom)
End Sub

Private Sub Form_Resize()
    fraReportSetup.Left = (Me.ScaleWidth - fraReportSetup.Width) / 2
End Sub

Private Sub optBigImgAction_Click(Index As Integer)
    If frmShowBigImg.Enabled = True Then
        cboZoom.Enabled = IIf(Index = 1, True, False)
    Else
        cboZoom.Enabled = False
    End If
End Sub

Private Sub optReportEditor_Click(Index As Integer)
    fraEditorSetUp.Visible = Index = 2
End Sub
