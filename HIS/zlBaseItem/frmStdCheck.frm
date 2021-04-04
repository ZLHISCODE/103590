VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStdCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "标准核查"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   Icon            =   "frmStdCheck.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6660
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd 
      Caption         =   "详细"
      Height          =   350
      Index           =   2
      Left            =   5460
      TabIndex        =   10
      Top             =   2325
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   5235
      TabIndex        =   16
      Top             =   3945
      Width           =   1100
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "检查(&K)"
      Height          =   350
      Left            =   3960
      TabIndex        =   14
      Top             =   3945
      Width           =   1100
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "打印(&P)..."
      Enabled         =   0   'False
      Height          =   350
      Left            =   165
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3975
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "详细"
      Height          =   350
      Index           =   3
      Left            =   5460
      TabIndex        =   13
      Top             =   3075
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmd 
      Caption         =   "详细"
      Height          =   350
      Index           =   1
      Left            =   5460
      TabIndex        =   7
      Top             =   1620
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmd 
      Caption         =   "详细"
      Height          =   350
      Index           =   0
      Left            =   5460
      TabIndex        =   4
      Top             =   915
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CheckBox chkCaption 
      Caption         =   "核查医院不符标准医价的项目"
      Height          =   300
      Index           =   3
      Left            =   420
      TabIndex        =   11
      Top             =   3120
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CheckBox chkCaption 
      Caption         =   "核查医院停用而标准医价未注销的项目"
      Height          =   300
      Index           =   2
      Left            =   420
      TabIndex        =   8
      Top             =   2370
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkCaption 
      Caption         =   "核查医院在用但标准医价已经注销的项目"
      Height          =   300
      Index           =   1
      Left            =   420
      TabIndex        =   5
      Top             =   1665
      Value           =   1  'Checked
      Width           =   3570
   End
   Begin VB.CheckBox chkCaption 
      Caption         =   "核查医院未正确对应标准医价的项目"
      Height          =   300
      Index           =   0
      Left            =   420
      TabIndex        =   2
      Top             =   945
      Value           =   1  'Checked
      Width           =   3180
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   1
      Left            =   15
      TabIndex        =   15
      Top             =   3780
      Width           =   6675
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   630
      Width           =   6675
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2100
      Index           =   1
      Left            =   660
      TabIndex        =   19
      Top             =   1980
      Visible         =   0   'False
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   3704
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2100
      Index           =   0
      Left            =   660
      TabIndex        =   18
      Top             =   1275
      Visible         =   0   'False
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   3704
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2280
      Index           =   2
      Left            =   660
      TabIndex        =   20
      Top             =   30
      Visible         =   0   'False
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   4022
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3030
      Index           =   3
      Left            =   660
      TabIndex        =   21
      Top             =   30
      Visible         =   0   'False
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   5345
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      Height          =   180
      Index           =   3
      Left            =   4110
      TabIndex        =   12
      Top             =   3165
      Width           =   1200
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      Height          =   180
      Index           =   2
      Left            =   4110
      TabIndex        =   9
      Top             =   2430
      Width           =   1200
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      Height          =   180
      Index           =   1
      Left            =   4110
      TabIndex        =   6
      Top             =   1710
      Width           =   1200
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      Height          =   180
      Index           =   0
      Left            =   4110
      TabIndex        =   3
      Top             =   1005
      Width           =   1200
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "请选择您要核查的内容"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   420
      TabIndex        =   0
      Top             =   210
      Width           =   2550
   End
   Begin VB.Menu mnuPop 
      Caption         =   "弹出打印"
      Visible         =   0   'False
      Begin VB.Menu mnuPopExcel 
         Caption         =   "输出到(&E)xcel"
      End
      Begin VB.Menu mnuPopPreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuPopPrint 
         Caption         =   "输出到打印机(&P)"
      End
   End
End
Attribute VB_Name = "frmStdCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrCol As String = "编码,1200,0,2;名称,1500,0,0;单位,1000,0,0"
Private Const mstrCol1 As String = "编码,1200,0,2;名称,1500,0,0;单位,1000,0,0;现价,1000,1,0;最高,1000,1,0;最低,1000,1,0"
Private mintColumn As Integer, mintColumn1 As Integer, mintColumn2 As Integer, mintColumn3 As Integer

Private Sub chkCaption_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim i As Long
    For i = Me.lvw.LBound To Me.lvw.UBound
        If i = Index Then
            Me.lvw(Index).Visible = Not Me.lvw(Index).Visible
            If Me.lvw(Index).Visible Then
                Me.lvw(Index).ZOrder
                Me.lvw(Index).SetFocus
                Me.cmdReport.Enabled = True
                Me.cmdReport.Tag = i
            Else
                Me.cmdReport.Enabled = False
            End If
        Else
            Me.lvw(i).Visible = False
        End If
    Next
End Sub

Private Sub cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCheck_Click()
On Error GoTo errHandle
    Dim i As Long
    Dim strSQL As String
    Dim strTmp As String
    Dim ObjItem As ListItem
    Dim rsTmp As New ADODB.Recordset
    Dim blnHave As Boolean
    
    Me.cmdClose.Enabled = False
    Me.cmdCheck.Enabled = False
    zlCommFun.ShowFlash "开始检查数据..."
    '先初始化
    For i = Me.lvw.LBound To Me.lvw.UBound
        Me.lvw(i).ListItems.Clear
        Me.cmd(i).Visible = False
        If Me.chkCaption(i).value = 1 And blnHave = False Then
            blnHave = True
        End If
        Me.lvw(i).Sorted = False
        Me.lvw(i).Visible = False
        Me.lblCaption(i).Caption = ""
    Next
    Me.cmdReport.Enabled = False
    If blnHave = False Then
        MsgBox "请选择要核查的内容！", vbInformation, gstrSysName
        Me.chkCaption(0).SetFocus
        Me.cmdClose.Enabled = True
        Me.cmdCheck.Enabled = True
        Exit Sub
    End If
    '根据设置进行核查
    strSQL = ""
    For i = Me.chkCaption.LBound To Me.chkCaption.UBound
        If Me.chkCaption(i).value = 1 Then
            Select Case True
                Case i = 0   '未正确对应标准医价项目的
                    strTmp = " SELECT 1 测试类型, '未正确对应' 说明,A.ID,A.编码,A.名称,A.计算单位 单位,B.最高限价,B.最低限价,0 现价 " & vbCrLf & _
                           " FROM 收费项目目录 A, 标准医价规范 B  " & vbCrLf & _
                           " WHERE A.类别 <> '5' AND A.类别 <> '6' AND A.类别 <> '7' " & vbCrLf & _
                           " AND A.标识主码=B.项目编码(+) AND B.项目编码 IS NULL "
                Case i = 1   '医院在用但标准中已经注销的
                    strTmp = "SELECT 2 测试类型, '医院在用标准中注销' 说明,A.ID,A.编码,A.名称,A.计算单位 单位,B.最高限价,B.最低限价,0 现价 " & vbCrLf & _
                            "  FROM 收费项目目录 A, 标准医价规范 B " & vbCrLf & _
                            "  WHERE A.类别 <> '5' AND A.类别 <> '6' AND A.类别 <> '7' AND A.标识主码=B.项目编码  " & vbCrLf & _
                            "    AND NVL(A.撤档时间,TO_DATE('3000-01-01','YYYY-MM-DD'))=TO_DATE('3000-01-01','YYYY-MM-DD') AND LTRIM(RTRIM(NVL(B.注销标志,'0')))='1' "
                Case i = 2   '医院停用但标准医价未注销的
                    strTmp = "SELECT 3 测试类型, '医院中停用标准中在用' 说明,A.ID,A.编码,A.名称,A.计算单位 单位,B.最高限价,B.最低限价,0 现价  " & vbCrLf & _
                            "  FROM 收费项目目录 A, 标准医价规范 B " & vbCrLf & _
                            "  WHERE A.类别 <> '5' AND A.类别 <> '6' AND A.类别 <> '7' AND A.标识主码=B.项目编码  " & vbCrLf & _
                            "    AND NVL(A.撤档时间,TO_DATE('3000-01-01','YYYY-MM-DD'))<>TO_DATE('3000-01-01','YYYY-MM-DD') AND NOT (LTRIM(RTRIM(NVL(B.注销标志,'0')))='1') "
                Case i = 3   '医院价格与标准医价不符合的
                    strTmp = " SELECT 4 测试类型, '最高与最低价不符' 说明,A.ID,A.编码,A.名称,A.计算单位 单位,B.最高限价,B.最低限价,0 现价  " & vbCrLf & _
                            "   FROM 收费项目目录 A, 标准医价规范 B " & vbCrLf & _
                            "   WHERE A.类别 <> '5' AND A.类别 <> '6' AND A.类别 <> '7' AND A.标识主码=B.项目编码  " & vbCrLf & _
                            "     AND (A.最高限价<>B.最高限价 OR A.最低限价<>B.最低限价) " & vbCrLf & _
                            " UNION ALL " & vbCrLf & _
                            " SELECT 5 测试类型, '当前价格不符' 说明,C.ID,C.编码,C.名称,C.计算单位 单位,B.最高限价,B.最低限价,SUM(A.现价) 现价 " & vbCrLf & _
                            "   FROM 收费价目 A,标准医价规范 B,  收费项目目录 C  " & vbCrLf & _
                            "  WHERE A.收费细目ID = C.ID  AND NVL(C.是否变价,0)=0  AND  C.标识主码=B.项目编码  " & vbCrLf & _
                            "    AND A.执行日期<=SYSDATE AND (A.终止日期>=SYSDATE OR A.终止日期 IS NULL)   " & vbCrLf & _
                            " GROUP BY C.ID,C.编码,C.名称,C.计算单位 ,B.最高限价,B.最低限价,A.价格等级 " & vbCrLf & _
                            " HAVING NOT (SUM(A.现价) >=B.最低限价 AND SUM(A.现价)<=B.最高限价) "
            End Select
            strSQL = strSQL & " UNION ALL " & vbCrLf & strTmp
        End If
    Next
    strSQL = Mid(strSQL, 11)
    strSQL = "select * from (" & strSQL & ")  order by 测试类型,编码"
    Call zldatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        If Me.chkCaption(0).value = 1 Then
            '未正确对应标准医价项目的
            rsTmp.Filter = "测试类型=1"
            If rsTmp.RecordCount > 0 Then
                zlCommFun.ShowFlash "正在处理“核查医院未正确对应标准医价的项目”..."
                rsTmp.MoveFirst
                For i = 1 To rsTmp.RecordCount
                    Set ObjItem = Me.lvw(0).ListItems.Add(, , zlCommFun.Nvl(rsTmp!编码))
                    ObjItem.SubItems(1) = zlCommFun.Nvl(rsTmp!名称)
                    ObjItem.SubItems(2) = zlCommFun.Nvl(rsTmp!单位)
                    rsTmp.MoveNext
                Next
            End If
        End If
        If Me.chkCaption(1).value = 1 Then
            '医院在用但标准中已经注销的
            rsTmp.Filter = "测试类型=2"
            If rsTmp.RecordCount > 0 Then
                zlCommFun.ShowFlash "正在处理“核查医院在用但标准医价已经注销的项目”..."
                rsTmp.MoveFirst
                For i = 1 To rsTmp.RecordCount
                    Set ObjItem = Me.lvw(1).ListItems.Add(, , zlCommFun.Nvl(rsTmp!编码))
                    ObjItem.SubItems(1) = zlCommFun.Nvl(rsTmp!名称)
                    ObjItem.SubItems(2) = zlCommFun.Nvl(rsTmp!单位)
                    rsTmp.MoveNext
                Next
            End If
        End If
        If Me.chkCaption(2).value = 1 Then
            '医院停用但标准医价未注销的
            rsTmp.Filter = "测试类型=3"
            If rsTmp.RecordCount > 0 Then
                zlCommFun.ShowFlash "正在处理“核查医院停用而标准医价未注销的项目”..."
                rsTmp.MoveFirst
                For i = 1 To rsTmp.RecordCount
                    Set ObjItem = Me.lvw(2).ListItems.Add(, , zlCommFun.Nvl(rsTmp!编码))
                    ObjItem.SubItems(1) = zlCommFun.Nvl(rsTmp!名称)
                    ObjItem.SubItems(2) = zlCommFun.Nvl(rsTmp!单位)
                    rsTmp.MoveNext
                Next
            End If
        End If
        If Me.chkCaption(3).value = 1 Then
            '修正最高与最低限价
            rsTmp.Filter = "测试类型=4"
            If rsTmp.RecordCount > 0 Then
                zlCommFun.ShowFlash "正在处理“核查医院不符标准医价的最高与最低限价的项目”..."
                rsTmp.MoveFirst
                For i = 1 To rsTmp.RecordCount
                    strSQL = "ZL_收费细目标准限价_UPDATE(" & rsTmp!ID & "," & rsTmp!最高限价 & "," & rsTmp!最低限价 & ")"
                    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
                    rsTmp.MoveNext
                Next
            End If
            '显示不符合医价的项目
            rsTmp.Filter = "测试类型=5"
            If rsTmp.RecordCount > 0 Then
                zlCommFun.ShowFlash "正在处理“核查医院不符标准医价的项目”..."
                rsTmp.MoveFirst
                For i = 1 To rsTmp.RecordCount
                    Set ObjItem = Me.lvw(3).ListItems.Add(, , zlCommFun.Nvl(rsTmp!编码))
                    ObjItem.SubItems(1) = zlCommFun.Nvl(rsTmp!名称)
                    ObjItem.SubItems(2) = zlCommFun.Nvl(rsTmp!单位)
                    ObjItem.SubItems(3) = CStr(Format(zlCommFun.Nvl(rsTmp!现价, 0), "0.00"))
                    ObjItem.SubItems(4) = CStr(Format(zlCommFun.Nvl(rsTmp!最高限价, 0), "0.00"))
                    ObjItem.SubItems(5) = CStr(Format(zlCommFun.Nvl(rsTmp!最低限价, 0), "0.00"))
                    rsTmp.MoveNext
                Next
            End If
        End If
    End If
    '设置显示状态
    For i = Me.lvw.LBound To Me.lvw.UBound
        If Me.lvw(i).ListItems.Count > 0 Then
            Me.cmd(i).Visible = True
            Me.lblCaption(i).Caption = "(" & Me.lvw(i).ListItems.Count & " 条)"
        Else
            If Me.chkCaption(i).value = 1 Then
                Me.lblCaption(i).Caption = "(无)"
            Else
                Me.lblCaption(i).Caption = ""
            End If
        End If
    Next
    Me.cmdClose.Enabled = True
    Me.cmdCheck.Enabled = True
    zlCommFun.ShowFlash
    Exit Sub
errHandle:
    If ERRCENTER() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.cmdClose.Enabled = True
    Me.cmdCheck.Enabled = True
    zlCommFun.ShowFlash
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdReport_Click()
    PopupMenu mnuPop
End Sub

Private Sub Form_Load()
    Dim i As Long
    For i = Me.lvw.LBound To Me.lvw.UBound
        If i = Me.lvw.UBound Then
            zlControl.LvwSelectColumns Me.lvw(i), mstrCol1, True
        Else
            zlControl.LvwSelectColumns Me.lvw(i), mstrCol, True
        End If
    Next
End Sub

Private Sub lvw_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvw(Index).Sorted = True
    If Choose(Index + 1, mintColumn, mintColumn1, mintColumn2, mintColumn3) = ColumnHeader.Index - 1 Then
        lvw(Index).SortOrder = IIF(lvw(Index).SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Select Case Index
            Case 0
                mintColumn = ColumnHeader.Index - 1
            Case 1
                mintColumn1 = ColumnHeader.Index - 1
            Case 2
                mintColumn2 = ColumnHeader.Index - 1
            Case 3
                mintColumn3 = ColumnHeader.Index - 1
        End Select
        lvw(Index).SortKey = ColumnHeader.Index - 1
        lvw(Index).SortOrder = lvwAscending
    End If
End Sub

Private Sub subPrint(ByVal Index As Long, bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    
    objPrint.Title.Text = "标准价格核查"
    Set objPrint.Body.objData = Me.lvw(Index)
    Select Case Index
        Case 0
            objPrint.UnderAppItems.Add "核查医院未正确对应标准医价的项目"
        Case 1
            objPrint.UnderAppItems.Add "核查医院在用但标准医价已经注销的项目"
        Case 2
            objPrint.UnderAppItems.Add "核查医院停用而标准医价未注销的项目"
        Case 3
            objPrint.UnderAppItems.Add "核查医院不符标准医价的项目"
    End Select
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zldatabase.Currentdate, "yyyy年MM月dd日")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Private Sub mnuPopExcel_Click()
    '输出到Excel
    Call subPrint(cmdReport.Tag, 3)
End Sub

Private Sub mnuPopPreview_Click()
    '打印预览
    Call subPrint(cmdReport.Tag, 2)
End Sub

Private Sub mnuPopPrint_Click()
    '输出到打印机
    Call subPrint(cmdReport.Tag, 1)
End Sub
